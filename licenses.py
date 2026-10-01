"""Parse a Waters Licensing Wizard PDF into a cleaned, categorized license list.

Input: the text of the PDF (extracted with poppler's pdftotext -layout).
Every interesting line looks like:

    [Empower 3 System Control License Pack] System licenses: 2 Serial No: G22L32477W

Cleanup rules (all requested by Matt, 2026-08-24):
  1. Strip a trailing -001 / -002 / ... pack-instance suffix from serials.
  2. A pack whose serial equals the base license's serial (e.g. the Named User
     pack that the base key bundles) is folded into the base row.
  3. No duplicate (license, serial) rows anywhere - system control, third-party
     instrument control, options. The first occurrence's quantity is kept.
Every dropped line is recorded with a reason so the Excel "Removed" sheet can
show it.
"""
from __future__ import annotations

import re
import subprocess
from dataclasses import dataclass, field

LINE_RE = re.compile(
    r"^\s*\[(?P<name>[^\]]+)\]\s*(?:(?P<qtylabel>[A-Za-z ]+?licenses):\s*(?P<qty>\d+)\s*)?"
    r"Serial No:\s*(?P<serial>\S+)\s*$"
)
SUFFIX_RE = re.compile(r"-\d{3}$")
HEADER_RE = re.compile(r"Waters Licensing Wizard\s*:\s*(?P<install>.+?)\s*$")
PRINTED_RE = re.compile(r"Date Printed:\s*(?P<date>.+?)\s*$")

# Checksum .txt grammar (Waters "Checksum_<date>.txt" report):
#   Option Properly Installed - 5 Named User License(s)          Serial Number - W3SAP2033M-001
CHK_LINE_RE = re.compile(
    r"^\s*Option Properly Installed\s*-\s*(?P<desc>.+?)\s{2,}Serial Number\s*-\s*(?P<serial>\S+)\s*$"
)
CHK_QTY_RE = re.compile(r"^(?P<qty>\d+)\s+(?P<rest>.+)$")
CHK_META = {
    "company": re.compile(r"^\s*Company Name\s*-\s*(?P<v>.+?)\s*$"),
    "support_id": re.compile(r"^\s*Support Plan ID\s*-\s*(?P<v>.+?)\s*$"),
    "installation": re.compile(r"^\s*Computer Name\s*-\s*(?P<v>.+?)\s*$"),
    "printed": re.compile(r"^\s*Current Date and Time\s*-\s*(?P<v>.+?)\s*$"),
}
# Third-party instrument control, by vendor in the name (the PDF says
# "Instrument control licenses:", the checksum file just says e.g.
# "Shimadzu LC Control" / "Shimadzu LC License(s)").
INSTRUMENT_RE = re.compile(r"\b(Agilent|Shimadzu|Thermo|Hitachi|PerkinElmer|Perkin Elmer|Bruker|Dionex)\b", re.I)

CATEGORY_ORDER = [
    "Base License",
    "User Licenses",
    "System Control",
    "Instrument Control (3rd Party)",
    "Options",
]


@dataclass
class License:
    name: str
    serial: str
    qty: int | None
    qty_label: str | None
    raw: str
    category: str = ""


@dataclass
class Removed:
    raw: str
    reason: str


@dataclass
class Result:
    installation: str | None
    printed: str | None
    company: str | None = None
    support_id: str | None = None
    licenses: list[License] = field(default_factory=list)
    removed: list[Removed] = field(default_factory=list)
    unparsed: list[str] = field(default_factory=list)


def categorize(name: str, qty_label: str | None) -> str:
    n = name.lower()
    if "base license" in n or "base package" in n:
        return "Base License"
    if "named user" in n or "user license" in n:
        return "User Licenses"
    if "system control" in n or n.startswith("system license"):
        return "System Control"
    if qty_label and "instrument" in qty_label.lower():
        return "Instrument Control (3rd Party)"
    if INSTRUMENT_RE.search(name):
        return "Instrument Control (3rd Party)"
    return "Options"


# Option licenses the Wizard prints without a count but which are really one
# seat each - shown (and exported) as Quantity 1.
DEFAULT_QTY_ONE = ("system suitability", "gpc", "sec", "dissolution")
# Whole words only. A plain substring test made "sec" match "Empower
# Security Option", which the Wizard genuinely prints without a count -
# and the workbook then claimed a quantity of 1 that Waters never stated.
# \b still matches inside punctuation, so "GPC/SEC Option" is unaffected.
DEFAULT_QTY_ONE_RE = re.compile(
    r"\b(?:" + "|".join(re.escape(k) for k in DEFAULT_QTY_ONE) + r")\b", re.I)


def default_qty(name: str, qty: int | None) -> int | None:
    if qty is None and DEFAULT_QTY_ONE_RE.search(name):
        return 1
    return qty


def clean_serial(serial: str) -> str:
    return SUFFIX_RE.sub("", serial.strip())


def parse_text(text: str) -> Result:
    res = Result(installation=None, printed=None)
    parsed: list[License] = []
    for line in text.splitlines():
        if not line.strip():
            continue
        m = HEADER_RE.search(line)
        if m and res.installation is None:
            res.installation = m.group("install")
            continue
        m = PRINTED_RE.search(line)
        if m:
            res.printed = m.group("date")
            continue
        m = LINE_RE.match(line)
        if not m:
            if line.strip().startswith("The following licenses"):
                continue
            res.unparsed.append(line.strip())
            continue
        lic = License(
            name=m.group("name").strip(),
            serial=clean_serial(m.group("serial")),
            qty=int(m.group("qty")) if m.group("qty") else None,
            qty_label=(m.group("qtylabel") or "").strip() or None,
            raw=line.strip(),
        )
        lic.category = categorize(lic.name, lic.qty_label)
        lic.qty = default_qty(lic.name, lic.qty)
        parsed.append(lic)

    return _finalize(res, parsed)


def parse_checksum_text(text: str) -> Result:
    """Parse a Waters Checksum_*.txt report into the same Result shape."""
    res = Result(installation=None, printed=None)
    parsed: list[License] = []
    for line in text.splitlines():
        if not line.strip():
            continue
        for key, rx in CHK_META.items():
            m = rx.match(line)
            if m and getattr(res, key) is None:
                setattr(res, key, m.group("v"))
                break
        else:
            m = CHK_LINE_RE.match(line)
            if not m:
                # Most of this file is CRC noise and listing it would bury
                # the signal. A line that announces an installed option but
                # doesn't fit the grammar is a different thing entirely: it
                # IS a license, and dropping it without a word left the
                # workbook quietly short a row - the exact failure the
                # Removed sheet exists to make visible. The grammar wants
                # two or more spaces before "Serial Number", so a report
                # printed with a tab or a single space lands here.
                if "Option Properly Installed" in line:
                    res.unparsed.append(line.strip())
                continue
            desc = m.group("desc").strip()
            qty, label = None, None
            qm = CHK_QTY_RE.match(desc)
            if qm:
                qty, desc = int(qm.group("qty")), qm.group("rest").strip()
                low = desc.lower()
                label = "User licenses" if "user" in low else "System licenses" if "system" in low else None
            lic = License(name=desc, serial=clean_serial(m.group("serial")), qty=qty, qty_label=label, raw=line.strip())
            lic.category = categorize(lic.name, lic.qty_label)
            if lic.category == "Instrument Control (3rd Party)" and lic.qty is None:
                lic.qty = 1
            lic.qty = default_qty(lic.name, lic.qty)
            parsed.append(lic)
    return _finalize(res, parsed)


# ---- Excel workbooks (.xlsx) -------------------------------------------------
# People hand these over in whatever shape they built: Wizard lines pasted one
# per cell (with a stray "[" lost here and there and extra columns of notes
# beside them), or a proper table with License / Qty / Serial columns. Each row
# is tried in that order:
#   1. any cell (or any line inside a multi-line cell) that reads like a Wizard
#      or Checksum line - the name, count and key all come from that text;
#   2. otherwise a table row: the serial from the column headed "Serial"/"Key"
#      (or, with no such header, a cell that looks like a Waters key), the
#      name from the column headed License/Description/... (or the longest
#      license-sounding cell), the count from a Qty/Count column.
# A row with a key-looking value that can't be paired with a license name
# goes to the Removed sheet as unrecognized, never silently dropped.
XL_LINE_RE = re.compile(
    r"^\s*\[?(?P<name>[^\[\]]+?)\s*\]\s*(?:(?P<qtylabel>[A-Za-z ]+?[Ll]icenses)\s*:\s*(?P<qty>\d+)\s*)?"
    r"[Ss]erial\s*(?:[Nn]o|[Nn]umber)\.?\s*[:\-]?\s*(?P<serial>[A-Za-z0-9]\S*)\s*$"
)
# Same line with the brackets gone entirely: the count label then has to be
# there to show where the name ends.
XL_BARE_RE = re.compile(
    r"^\s*(?P<name>[^\s\[\]][^\[\]]*?)\s+(?P<qtylabel>[A-Za-z]+(?: [Cc]ontrol)? [Ll]icenses)\s*:\s*(?P<qty>\d+)\s*"
    r"[Ss]erial\s*(?:[Nn]o|[Nn]umber)\.?\s*[:\-]?\s*(?P<serial>[A-Za-z0-9]\S*)\s*$"
)
XL_SERIAL_CELL_RE = re.compile(r"^(?:[Ss]erial\s*(?:[Nn]o|[Nn]umber)\.?\s*[:\-]?\s*)?(?P<serial>[A-Za-z0-9][A-Za-z0-9-]{3,29})$")
# A Waters key with no header to vouch for it: 10 or 18 upper-case letters and
# digits (both present), optional -NNN pack suffix. Shorter codes such as
# database names (SBEP0573) don't qualify.
WATERS_KEY_RE = re.compile(r"^(?=[A-Z0-9]*[A-Z])(?=[A-Z0-9]*\d)(?:[A-Z0-9]{18}|[A-Z0-9]{10})(?:-\d{3})?$")
XL_SERIAL_HDR_RE = re.compile(r"serial|\bkeys?\b", re.I)
XL_NAME_HDR_RE = re.compile(r"licen[cs]e|description|option|product|\bname\b|\btype\b|\bitem\b|module", re.I)
XL_QTY_HDR_RE = re.compile(r"\b(?:qty|quantity|counts?|seats?|units?)\b", re.I)
XL_LICENSE_WORD_RE = re.compile(
    r"licen[cs]e|empower|option|control|\bpack\b|\buser|\bsystem\b|" + INSTRUMENT_RE.pattern[3:-3], re.I)
XL_INT_RE = re.compile(r"^\d+$")


def _xl_ref(col: int, row: int) -> str:
    s = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        s = chr(65 + rem) + s
    return f"{s}{row}"


def _xl_line(line: str) -> tuple[License, bool] | None:
    """(license, came-from-a-Checksum-style-line) or None."""
    m = XL_LINE_RE.match(line) or XL_BARE_RE.match(line)
    if m:
        return License(
            name=m.group("name").strip(), serial=clean_serial(m.group("serial")),
            qty=int(m.group("qty")) if m.group("qty") else None,
            qty_label=(m.group("qtylabel") or "").strip() or None, raw=line.strip()), False
    m = CHK_LINE_RE.match(line)
    if m:
        desc, qty, label = m.group("desc").strip(), None, None
        qm = CHK_QTY_RE.match(desc)
        if qm:
            qty, desc = int(qm.group("qty")), qm.group("rest").strip()
            low = desc.lower()
            label = "User licenses" if "user" in low else "System licenses" if "system" in low else None
        return License(name=desc, serial=clean_serial(m.group("serial")), qty=qty, qty_label=label, raw=line.strip()), True
    return None


def _xl_header(cells: list) -> dict | None:
    """Column roles if this row is a table header, else None."""
    if any(WATERS_KEY_RE.match(t) for _, t in cells):
        return None
    # A label ("Serial Number", "License key"), not a broken data line that
    # happens to say "Serial No:" - those have a count or key digits in them.
    serial = next((c for c, t in cells if XL_SERIAL_HDR_RE.search(t) and len(t) <= 40 and not re.search(r"\d", t)), None)
    if serial is None:
        return None
    name = next((c for c, t in cells if c != serial and XL_NAME_HDR_RE.search(t)), None)
    qty = next((c for c, t in cells if c not in (serial, name) and XL_QTY_HDR_RE.search(t)), None)
    return {"serial": serial, "name": name, "qty": qty}


def parse_grid(sheets: list[dict]) -> Result:
    """Scan an Excel workbook's cells (xlsx_reader.read_xlsx_grid) for licenses."""
    res = Result(installation=None, printed=None)
    parsed: list[License] = []
    for sheet in sheets:
        hdr = None
        for row in sheet["rows"]:
            r, cells = row["r"], row["cells"]
            by_col = dict((c, t) for c, t in cells)
            qty_cell = by_col.get(hdr["qty"]) if hdr and hdr["qty"] else None
            found = []
            for col, text in cells:
                for line in text.splitlines():
                    hit = _xl_line(line)
                    if hit is not None:
                        hit[0].raw = f"{sheet['name']}!{_xl_ref(col, r)}: {hit[0].raw}"
                        found.append(hit)
            # A count column only speaks for its row when the row holds one
            # license - a cell with three pasted lines has one count beside it.
            if len(found) == 1 and found[0][0].qty is None and qty_cell and XL_INT_RE.match(qty_cell):
                found[0][0].qty = int(qty_cell)
            for lic, chk in found:
                lic.category = categorize(lic.name, lic.qty_label)
                if lic.category == "Instrument Control (3rd Party)" and lic.qty is None and chk:
                    lic.qty = 1  # same default as parse_checksum_text
                lic.qty = default_qty(lic.name, lic.qty)
                parsed.append(lic)
            if found:
                continue
            h = _xl_header(cells)
            if h:
                hdr = h
                continue

            raw = f"{sheet['name']}!{r}: " + " | ".join(t for _, t in cells)
            serial_col = serial = None
            if hdr and hdr["serial"] in by_col:
                m = XL_SERIAL_CELL_RE.match(by_col[hdr["serial"]])
                if m:
                    serial_col, serial = hdr["serial"], m.group("serial")
            if serial is None:
                for col, text in cells:
                    m = XL_SERIAL_CELL_RE.match(text)
                    if m and WATERS_KEY_RE.match(m.group("serial")):
                        serial_col, serial = col, m.group("serial")
                        break
            if serial is None:
                # Not a license row - unless it plainly talks about one.
                if re.search(r"[Ss]erial\s*(?:[Nn]o|[Nn]umber)\b", raw) or any(
                        WATERS_KEY_RE.match(w) for _, t in cells for w in t.split()):
                    res.unparsed.append(raw)
                continue
            name = None
            if hdr and hdr["name"] and hdr["name"] != serial_col and hdr["name"] in by_col:
                name = by_col[hdr["name"]]
            if name is None:
                words = [t for c, t in cells if c != serial_col and not XL_INT_RE.match(t)
                         and XL_LICENSE_WORD_RE.search(t)]
                name = max(words, key=len) if words else None
            if not name:
                res.unparsed.append(raw)
                continue
            q = by_col.get(hdr["qty"]) if hdr and hdr["qty"] else None
            lic = License(name=name, serial=clean_serial(serial),
                          qty=int(q) if q and XL_INT_RE.match(q) else None, qty_label=None, raw=raw)
            lic.category = categorize(lic.name, None)
            lic.qty = default_qty(lic.name, lic.qty)
            parsed.append(lic)
    return _finalize(res, parsed)


def _finalize(res: Result, parsed: list[License]) -> Result:
    base_serials = {l.serial for l in parsed if l.category == "Base License"}
    seen: set[str] = set()
    bases = {l.serial: l for l in parsed if l.category == "Base License"}
    for lic in parsed:
        if lic.category != "Base License" and lic.serial in base_serials:
            base = bases[lic.serial]
            if base.qty is None and lic.qty is not None:
                base.qty, base.qty_label = lic.qty, lic.qty_label  # checksum file prints the count on the folded line
            res.removed.append(Removed(lic.raw, "Included with base license (same serial)"))
            continue
        # Dedupe on serial alone: one key = one row, whatever label the report
        # printed next to it (e.g. "Shimadzu LC Control" + "Shimadzu LC License(s)").
        if lic.serial in seen:
            res.removed.append(Removed(lic.raw, "Duplicate serial"))
            continue
        seen.add(lic.serial)
        res.licenses.append(lic)

    order = {c: i for i, c in enumerate(CATEGORY_ORDER)}
    res.licenses.sort(key=lambda l: (order[l.category], l.name.lower(), l.serial))
    return res


def remove_sqt(res: Result) -> Result:
    """Drop every license with 'SQT' in its name (SystemsQT, Software SQT ...),
    recording each one on the Removed list."""
    keep = []
    for lic in res.licenses:
        if "sqt" in lic.name.lower():
            res.removed.append(Removed(lic.raw, "SQT removed (checkbox)"))
        else:
            keep.append(lic)
    res.licenses = keep
    return res


def remove_zero_qty(res: Result) -> Result:
    """Drop every license whose printed quantity is 0."""
    keep = []
    for lic in res.licenses:
        if lic.qty == 0:
            res.removed.append(Removed(lic.raw, "Zero quantity removed (checkbox)"))
        else:
            keep.append(lic)
    res.licenses = keep
    return res


def pdf_to_text(pdf_bytes: bytes) -> str:
    try:
        proc = subprocess.run(
            ["pdftotext", "-layout", "-", "-"],
            input=pdf_bytes, capture_output=True, timeout=60,
        )
    except subprocess.TimeoutExpired as exc:
        # ValueError, not the raw TimeoutExpired: app._get_result() turns a
        # ValueError into a 400 with the message on it, while anything else
        # escapes as a 500 and tells the user nothing. A PDF that takes
        # poppler more than a minute is a bad upload, not a broken server.
        raise ValueError(
            "That PDF took too long to read - it may be corrupt or very "
            "large. Try re-printing it from the Licensing Wizard.") from exc
    if proc.returncode != 0:
        raise ValueError("Could not read that PDF: " + proc.stderr.decode(errors="replace")[:200])
    return proc.stdout.decode("utf-8", errors="replace")


def parse_pdf(pdf_bytes: bytes) -> Result:
    return parse_text(pdf_to_text(pdf_bytes))


# The PDF spec puts the %PDF header at the start of the file but tolerates
# it anywhere in the first 1024 bytes, and readers follow suit - a byte-order
# mark or a stray banner ahead of it is common enough in files that have been
# through an email gateway. Requiring it at offset 0 sent those to the text
# decoder, which then told the user their PDF was not a PDF.
PDF_HEADER_WINDOW = 1024


def parse_xlsx(data: bytes) -> Result:
    from xlsx_reader import read_xlsx_grid
    return parse_grid(read_xlsx_grid(data))


def parse_upload(data: bytes) -> Result:
    """PDF (Licensing Wizard printout), .txt (Checksum report) or .xlsx
    (any workbook with the lines/keys in it) - sniffed by content."""
    if b"%PDF" in data[:PDF_HEADER_WINDOW]:
        return parse_pdf(data)
    if data[:4] == b"PK\x03\x04":
        return parse_xlsx(data)
    if data[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        raise ValueError("That's an old-style .xls (or password-protected) Excel file - "
                         "open it in Excel and Save As .xlsx, then try again.")
    text = data.decode("utf-8", errors="replace")
    if "Option Properly Installed" in text or "Waters File Verification" in text:
        return parse_checksum_text(text)
    if "Serial No:" in text:
        return parse_text(text)
    raise ValueError("Not a Licensing Wizard PDF, a Checksum .txt report or an Excel .xlsx workbook.")
