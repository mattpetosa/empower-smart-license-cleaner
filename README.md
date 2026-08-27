# Empower Smart License Cleaner (ESLC)

**Empower Smart License Cleaner (ESLC)** — live at https://licenses.mhpwebserver.com, or download the single-file offline version from the site footer. Sample inputs in `fixtures/` are synthetic (fake serials, placeholder company).

Upload a Waters Licensing Wizard PDF printout, get back an Excel workbook with
the licenses categorized and cleaned up.

Cleanup rules (`licenses.py`):
1. `-001`/`-002`… pack-instance suffixes are stripped from serial numbers.
2. Any pack whose serial equals the base license serial (the bundled Named User
   pack) is folded into the base license row.
3. Duplicate (license, serial) rows are removed — system control, third-party
   instrument control, options alike.

Workbook sheets: **Licenses** (Category / License / Quantity / Type / Serial),
**Summary** (installation, print date, per-category counts), **Removed** (every
line dropped by the rules above, with the reason).

Stateless: the PDF is parsed in memory with poppler `pdftotext`; nothing is
written to disk.

## Dev
    ./venv/bin/pytest
    ./venv/bin/flask --app app run --port 8797

## Deploy
    sudo cp deploy/licenses-backend.service /etc/systemd/system/
    sudo cp deploy/licenses.mhpwebserver.com.conf /etc/nginx/sites-available/
    sudo ln -s /etc/nginx/sites-available/licenses.mhpwebserver.com.conf /etc/nginx/sites-enabled/
    sudo systemctl daemon-reload && sudo systemctl enable --now licenses-backend
    sudo nginx -t && sudo systemctl reload nginx

## Releasing
Bump `VERSION`, `python3 offline/build.py`, commit, then `git tag -a vX.Y.Z`, push, `cp www/offline/ESLC.html /tmp/EmpowerSmartLicenseCleaner_vX.Y.Z.html` and `gh release create vX.Y.Z /tmp/EmpowerSmartLicenseCleaner_vX.Y.Z.html` — the release asset is the offline file. (The `file#label` form only changes the display label; the asset keeps the on-disk filename, so copy it first.)

## Versioning
Bump `VERSION` (semver, shown in both footers), then `python3 offline/build.py` so the offline file carries it.

## Offline version
`www/offline/ESLC.html` is a single self-contained file (served at
`/offline/ESLC.html`, linked in the footer): the same parser ported to JS
(`offline/core.js`), a hand-rolled XLSX/zip writer, and pdf.js embedded inline.
Runs entirely in the browser; nothing leaves the machine.

    cd offline && npm install          # once, pulls pdfjs-dist
    node offline/parity_test.mjs       # JS core must match Python row-for-row on both fixtures
    python3 offline/build.py           # regenerate www/offline/ESLC.html (do this after ANY parser change)

`licenses.py` and `offline/core.js` are deliberately kept in lock-step; the
parity test is what stops them drifting.
