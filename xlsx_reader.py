"""Read an .xlsx workbook into a plain cell grid - just the text of every cell.

Deliberately not openpyxl: offline/core.js has to read the same workbooks in
the browser with no library, and the two must see exactly the same cell text
or the parity test can't hold them to the same answer. So both sides use the
same small regex reader over the SpreadsheetML parts: cached values only
(a formula with no cached result reads as empty), numbers as Excel stored
them, shared/inline/rich strings flattened to their text.

Result: [{"name": sheet name, "rows": [{"r": row number, "cells": [[col, text], ...]}]}]
with 1-based row/column numbers and empty cells left out.
"""
from __future__ import annotations

import io
import re
import zipfile

# Per-part cap on decompressed size. A real license sheet is a few hundred KB;
# this only exists so a zip bomb can't take the worker down.
MAX_PART_BYTES = 50 * 1024 * 1024
MAX_COL = 16384

_P = r"(?:[A-Za-z_][\w.-]*:)?"  # optional namespace prefix (Open XML SDK writes x:c, x:row ...)
SHEET_RE = re.compile(r"<" + _P + r"sheet\b([^>]*)>")
REL_RE = re.compile(r"<" + _P + r"Relationship\b([^>]*)>")
SI_RE = re.compile(r"<" + _P + r"si\b[^>]*?(?:/>|>(.*?)</" + _P + r"si>)", re.S)
T_RE = re.compile(r"<" + _P + r"t\b[^>]*?(?:/>|>(.*?)</" + _P + r"t>)", re.S)
RPH_RE = re.compile(r"<" + _P + r"rPh\b.*?</" + _P + r"rPh>", re.S)
ROW_RE = re.compile(r"<" + _P + r"row\b([^>]*?)(?:/>|>(.*?)</" + _P + r"row>)", re.S)
CELL_RE = re.compile(r"<" + _P + r"c\b([^>]*?)(?:/>|>(.*?)</" + _P + r"c>)", re.S)
V_RE = re.compile(r"<" + _P + r"v\b[^>]*?(?:/>|>(.*?)</" + _P + r"v>)", re.S)
IS_RE = re.compile(r"<" + _P + r"is\b[^>]*?(?:/>|>(.*?)</" + _P + r"is>)", re.S)
REF_RE = re.compile(r"^([A-Za-z]{1,3})(\d+)$")
ENTITY_RE = re.compile(r"&(#[xX][0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);")
OOXML_ESC_RE = re.compile(r"_x([0-9a-fA-F]{4})_")
ENTITIES = {"amp": "&", "lt": "<", "gt": ">", "quot": '"', "apos": "'"}


def _attr(attrs: str, name: str) -> str | None:
    m = re.search(r"(?:^|\s)" + _P + name + r"\s*=\s*\"([^\"]*)\"", attrs)
    if not m:
        m = re.search(r"(?:^|\s)" + _P + name + r"\s*=\s*'([^']*)'", attrs)
    return _unescape(m.group(1)) if m else None


def _entity(m: re.Match) -> str:
    e = m.group(1)
    if e[0] == "#":
        try:
            cp = int(e[2:], 16) if e[1] in "xX" else int(e[1:])
            return chr(cp) if 0 < cp <= 0x10FFFF else ""
        except ValueError:
            return ""
    return ENTITIES[e]


def _unescape(s: str) -> str:
    s = ENTITY_RE.sub(_entity, s)
    return OOXML_ESC_RE.sub(lambda m: chr(int(m.group(1), 16)), s)


def _text_of(xml: str) -> str:
    """All <t> runs of a shared/inline string, phonetic (rPh) runs excluded."""
    xml = RPH_RE.sub("", xml)
    return "".join(_unescape(m.group(1) or "") for m in T_RE.finditer(xml))


def _col_index(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _resolve(target: str) -> str:
    """A workbook relationship target, relative to xl/ unless absolute."""
    if target.startswith("/"):
        return target[1:]
    parts = []
    for p in ("xl/" + target).split("/"):
        if p == "..":
            if parts:
                parts.pop()
        elif p and p != ".":
            parts.append(p)
    return "/".join(parts)


def _read(zf: zipfile.ZipFile, name: str) -> str | None:
    try:
        info = zf.getinfo(name)
    except KeyError:
        return None
    with zf.open(info) as f:
        data = f.read(MAX_PART_BYTES + 1)
    if len(data) > MAX_PART_BYTES:
        raise ValueError("That workbook is too large to read.")
    return data.decode("utf-8", errors="replace")


def parse_sheet_xml(xml: str, shared: list[str]) -> list[dict]:
    rows = []
    next_row = 1
    for rm in ROW_RE.finditer(xml):
        r_attr = _attr(rm.group(1), "r")
        rnum = int(r_attr) if r_attr and r_attr.isdigit() else next_row
        next_row = rnum + 1
        cells = []
        next_col = 1
        for cm in CELL_RE.finditer(rm.group(2) or ""):
            attrs, inner = cm.group(1), cm.group(2) or ""
            ref = REF_RE.match(_attr(attrs, "r") or "")
            col = _col_index(ref.group(1)) if ref else next_col
            next_col = col + 1
            if col > MAX_COL:
                continue
            t = _attr(attrs, "t") or "n"
            if t == "inlineStr":
                im = IS_RE.search(inner)
                text = _text_of(im.group(1) or "") if im else ""
            else:
                vm = V_RE.search(inner)
                raw = (vm.group(1) or "") if vm else ""
                text = _unescape(raw)
                if t == "s":
                    idx = raw.strip()
                    text = shared[int(idx)] if idx.isdigit() and int(idx) < len(shared) else ""
                elif t == "b":
                    text = "TRUE" if raw.strip() == "1" else "FALSE" if raw.strip() == "0" else raw
            text = text.strip()
            if text:
                cells.append([col, text])
        if cells:
            rows.append({"r": rnum, "cells": cells})
    return rows


def read_xlsx_grid(data: bytes) -> list[dict]:
    try:
        zf = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile as exc:
        raise ValueError("Could not read that Excel file - it isn't a valid .xlsx workbook.") from exc
    with zf:
        workbook = _read(zf, "xl/workbook.xml")
        if workbook is None:
            raise ValueError("Not an Excel workbook (.xlsx).")
        rels = {}
        shared_path = "xl/sharedStrings.xml"
        for m in REL_RE.finditer(_read(zf, "xl/_rels/workbook.xml.rels") or ""):
            rid, target, typ = _attr(m.group(1), "Id"), _attr(m.group(1), "Target"), _attr(m.group(1), "Type") or ""
            if rid and target:
                rels[rid] = _resolve(target)
                if typ.endswith("/sharedStrings"):
                    shared_path = rels[rid]
        shared = [_text_of(m.group(1) or "") for m in SI_RE.finditer(_read(zf, shared_path) or "")]

        sheets = []
        for m in SHEET_RE.finditer(workbook):
            path = rels.get(_attr(m.group(1), "id") or "")
            if not path:
                continue
            xml = _read(zf, path)
            if xml is None:
                continue
            sheets.append({"name": _attr(m.group(1), "name") or "Sheet", "rows": parse_sheet_xml(xml, shared)})
        return sheets
