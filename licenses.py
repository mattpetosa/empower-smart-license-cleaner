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


def default_qty(name: str, qty: int | None) -> int | None:
    if qty is None and any(k in name.lower().split() or k in name.lower() for k in DEFAULT_QTY_ONE):
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
                continue  # everything else in this file is CRC noise, not worth listing
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
    proc = subprocess.run(
        ["pdftotext", "-layout", "-", "-"],
        input=pdf_bytes, capture_output=True, timeout=60,
    )
    if proc.returncode != 0:
        raise ValueError("Could not read that PDF: " + proc.stderr.decode(errors="replace")[:200])
    return proc.stdout.decode("utf-8", errors="replace")


def parse_pdf(pdf_bytes: bytes) -> Result:
    return parse_text(pdf_to_text(pdf_bytes))


def parse_upload(data: bytes) -> Result:
    """PDF (Licensing Wizard printout) or .txt (Checksum report) - sniffed by content."""
    if data.startswith(b"%PDF"):
        return parse_pdf(data)
    text = data.decode("utf-8", errors="replace")
    if "Option Properly Installed" in text or "Waters File Verification" in text:
        return parse_checksum_text(text)
    if "Serial No:" in text:
        return parse_text(text)
    raise ValueError("Not a Licensing Wizard PDF or a Checksum .txt report.")
