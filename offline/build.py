#!/usr/bin/env python3
"""Assemble the single-file offline ESLC.html: template + site CSS + core.js +
pdf.js (main + worker, both embedded as JS string literals and loaded from Blob
URLs at runtime, so the file has zero external dependencies)."""
import json, subprocess, datetime
from pathlib import Path

here = Path(__file__).resolve().parent
root = here.parent
pdfjs = here / "node_modules/pdfjs-dist/legacy/build"

tpl = (here / "template.html").read_text()
css = (root / "www/static/app.css").read_text()
core = (here / "core.js").read_text()
main = (pdfjs / "pdf.min.mjs").read_text()
worker = (pdfjs / "pdf.worker.min.mjs").read_text()
build = (root / "VERSION").read_text().strip()

def js_string(s: str) -> str:
    # JSON string literal is a valid JS string literal; keep </script> from ending the tag.
    return json.dumps(s).replace("</", "<\\/")

assert "</script" not in core and "</style" not in css
out = (tpl.replace("/*__CSS__*/", css)
          .replace("/*__CORE__*/", core)
          .replace("/*__PDFJS__*/", js_string(main))
          .replace("/*__WORKER__*/", js_string(worker))
          .replace("__BUILD__", build))
dest = root / "www/offline/ESLC.html"
dest.parent.mkdir(parents=True, exist_ok=True)
dest.write_text(out)
print(f"wrote {dest} ({dest.stat().st_size/1024/1024:.2f} MB, build {build})")
