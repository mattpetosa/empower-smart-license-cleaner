/* ESLC core: parser + XLSX/CSV writers. Mirrors licenses.py / excel.py exactly.
   Plain script: works in the browser (window.ESLC) and in Node (module.exports)
   so tests can diff it against the Python implementation. */
(function (root) {
  'use strict';

  const CATEGORY_ORDER = ['Base License', 'User Licenses', 'System Control', 'Instrument Control (3rd Party)', 'Options'];
  const LINE_RE = /^\s*\[([^\]]+)\]\s*(?:([A-Za-z ]+?licenses):\s*(\d+)\s*)?Serial No:\s*(\S+)\s*$/;
  const SUFFIX_RE = /-\d{3}$/;
  const HEADER_RE = /Waters Licensing Wizard\s*:\s*(.+?)\s*$/;
  const PRINTED_RE = /Date Printed:\s*(.+?)\s*$/;
  const CHK_LINE_RE = /^\s*Option Properly Installed\s*-\s*(.+?)\s{2,}Serial Number\s*-\s*(\S+)\s*$/;
  const CHK_QTY_RE = /^(\d+)\s+(.+)$/;
  const CHK_META = {
    company: /^\s*Company Name\s*-\s*(.+?)\s*$/,
    support_id: /^\s*Support Plan ID\s*-\s*(.+?)\s*$/,
    installation: /^\s*Computer Name\s*-\s*(.+?)\s*$/,
    printed: /^\s*Current Date and Time\s*-\s*(.+?)\s*$/,
  };
  const INSTRUMENT_RE = /\b(Agilent|Shimadzu|Thermo|Hitachi|PerkinElmer|Perkin Elmer|Bruker|Dionex)\b/i;
  const DEFAULT_QTY_ONE = ['system suitability', 'gpc', 'sec', 'dissolution'];
  // Whole words only, matching licenses.py's DEFAULT_QTY_ONE_RE. A plain
  // substring test made 'sec' match 'Empower Security Option', which the
  // Wizard genuinely prints without a count - and the workbook then claimed
  // a quantity of 1 that Waters never stated. \b still matches inside
  // punctuation, so 'GPC/SEC Option' is unaffected.
  const DEFAULT_QTY_ONE_RE = new RegExp('\\b(?:' + DEFAULT_QTY_ONE.join('|') + ')\\b', 'i');

  function categorize(name, qtyLabel) {
    const n = name.toLowerCase();
    if (n.includes('base license') || n.includes('base package')) return 'Base License';
    if (n.includes('named user') || n.includes('user license')) return 'User Licenses';
    if (n.includes('system control') || n.startsWith('system license')) return 'System Control';
    if (qtyLabel && qtyLabel.toLowerCase().includes('instrument')) return 'Instrument Control (3rd Party)';
    if (INSTRUMENT_RE.test(name)) return 'Instrument Control (3rd Party)';
    return 'Options';
  }
  function defaultQty(name, qty) {
    if (qty === null && DEFAULT_QTY_ONE_RE.test(name)) return 1;
    return qty;
  }
  const cleanSerial = s => s.trim().replace(SUFFIX_RE, '');

  function newResult() { return { installation: null, printed: null, company: null, support_id: null, licenses: [], removed: [], unparsed: [] }; }

  function parseWizardText(text) {
    const res = newResult(), parsed = [];
    for (const line of text.split(/\r?\n|\r/)) {
      if (!line.trim()) continue;
      let m = HEADER_RE.exec(line);
      if (m && res.installation === null) { res.installation = m[1]; continue; }
      m = PRINTED_RE.exec(line);
      if (m) { res.printed = m[1]; continue; }
      m = LINE_RE.exec(line);
      if (!m) { if (!line.trim().startsWith('The following licenses')) res.unparsed.push(line.trim()); continue; }
      const lic = { name: m[1].trim(), serial: cleanSerial(m[4]), qty: m[3] ? parseInt(m[3], 10) : null, qty_label: (m[2] || '').trim() || null, raw: line.trim() };
      lic.category = categorize(lic.name, lic.qty_label);
      lic.qty = defaultQty(lic.name, lic.qty);
      parsed.push(lic);
    }
    return finalize(res, parsed);
  }

  function parseChecksumText(text) {
    const res = newResult(), parsed = [];
    for (const line of text.split(/\r?\n|\r/)) {
      if (!line.trim()) continue;
      let hit = false;
      for (const [k, rx] of Object.entries(CHK_META)) { const m = rx.exec(line); if (m && res[k] === null) { res[k] = m[1]; hit = true; break; } }
      if (hit) continue;
      const m = CHK_LINE_RE.exec(line);
      // Most of this file is CRC noise and listing it would bury the signal.
      // A line that announces an installed option but doesn't fit the
      // grammar is a different thing entirely: it IS a license, and dropping
      // it without a word left the workbook quietly short a row. Mirrors
      // licenses.parse_checksum_text.
      if (!m) { if (line.includes('Option Properly Installed')) res.unparsed.push(line.trim()); continue; }
      let desc = m[1].trim(), qty = null, label = null;
      const qm = CHK_QTY_RE.exec(desc);
      if (qm) { qty = parseInt(qm[1], 10); desc = qm[2].trim(); const low = desc.toLowerCase(); label = low.includes('user') ? 'User licenses' : low.includes('system') ? 'System licenses' : null; }
      const lic = { name: desc, serial: cleanSerial(m[2]), qty, qty_label: label, raw: line.trim() };
      lic.category = categorize(lic.name, lic.qty_label);
      if (lic.category === 'Instrument Control (3rd Party)' && lic.qty === null) lic.qty = 1;
      lic.qty = defaultQty(lic.name, lic.qty);
      parsed.push(lic);
    }
    return finalize(res, parsed);
  }

  function finalize(res, parsed) {
    const bases = {}; for (const l of parsed) if (l.category === 'Base License') bases[l.serial] = l;
    const seen = new Set();
    for (const lic of parsed) {
      if (lic.category !== 'Base License' && bases[lic.serial]) {
        const b = bases[lic.serial];
        if (b.qty === null && lic.qty !== null) { b.qty = lic.qty; b.qty_label = lic.qty_label; }
        res.removed.push({ raw: lic.raw, reason: 'Included with base license (same serial)' }); continue;
      }
      if (seen.has(lic.serial)) { res.removed.push({ raw: lic.raw, reason: 'Duplicate serial' }); continue; }
      seen.add(lic.serial); res.licenses.push(lic);
    }
    const order = Object.fromEntries(CATEGORY_ORDER.map((c, i) => [c, i]));
    res.licenses.sort((a, b) => (order[a.category] - order[b.category]) || cmp(a.name.toLowerCase(), b.name.toLowerCase()) || cmp(a.serial, b.serial));
    return res;
  }
  const cmp = (a, b) => (a < b ? -1 : a > b ? 1 : 0);

  /* ---- Excel workbooks: grid -> licenses (mirrors licenses.parse_grid) ---- */
  const XL_LINE_RE = /^\s*\[?([^\[\]]+?)\s*\]\s*(?:([A-Za-z ]+?[Ll]icenses)\s*:\s*(\d+)\s*)?[Ss]erial\s*(?:[Nn]o|[Nn]umber)\.?\s*[:\-]?\s*([A-Za-z0-9]\S*)\s*$/;
  const XL_BARE_RE = /^\s*([^\s\[\]][^\[\]]*?)\s+([A-Za-z]+(?: [Cc]ontrol)? [Ll]icenses)\s*:\s*(\d+)\s*[Ss]erial\s*(?:[Nn]o|[Nn]umber)\.?\s*[:\-]?\s*([A-Za-z0-9]\S*)\s*$/;
  const XL_SERIAL_CELL_RE = /^(?:[Ss]erial\s*(?:[Nn]o|[Nn]umber)\.?\s*[:\-]?\s*)?([A-Za-z0-9][A-Za-z0-9-]{3,29})$/;
  const WATERS_KEY_RE = /^(?=[A-Z0-9]*[A-Z])(?=[A-Z0-9]*\d)(?:[A-Z0-9]{18}|[A-Z0-9]{10})(?:-\d{3})?$/;
  const XL_SERIAL_HDR_RE = /serial|\bkeys?\b/i;
  const XL_NAME_HDR_RE = /licen[cs]e|description|option|product|\bname\b|\btype\b|\bitem\b|module/i;
  const XL_QTY_HDR_RE = /\b(?:qty|quantity|counts?|seats?|units?)\b/i;
  const XL_LICENSE_WORD_RE = /licen[cs]e|empower|option|control|\bpack\b|\buser|\bsystem\b|Agilent|Shimadzu|Thermo|Hitachi|PerkinElmer|Perkin Elmer|Bruker|Dionex/i;
  const XL_INT_RE = /^\d+$/;
  const XL_MENTIONS_SERIAL_RE = /[Ss]erial\s*(?:[Nn]o|[Nn]umber)\b/;
  // Python's str.splitlines() boundaries, so a multi-line cell splits the same way.
  const splitLines = t => t.split(/\r\n|[\n\r\v\f\x1c-\x1e\x85\u2028\u2029]/);

  function xlRef(col, row) { return colL(col) + row; }

  function xlLine(line) {
    let m = XL_LINE_RE.exec(line) || XL_BARE_RE.exec(line);
    if (m) return { name: m[1].trim(), serial: cleanSerial(m[4]), qty: m[3] ? parseInt(m[3], 10) : null, qty_label: (m[2] || '').trim() || null, raw: line.trim() };
    m = CHK_LINE_RE.exec(line);
    if (m) {
      let desc = m[1].trim(), qty = null, label = null;
      const qm = CHK_QTY_RE.exec(desc);
      if (qm) { qty = parseInt(qm[1], 10); desc = qm[2].trim(); const low = desc.toLowerCase(); label = low.includes('user') ? 'User licenses' : low.includes('system') ? 'System licenses' : null; }
      return { name: desc, serial: cleanSerial(m[2]), qty, qty_label: label, raw: line.trim(), chk: true };
    }
    return null;
  }

  function xlHeader(cells) {
    if (cells.some(([, t]) => WATERS_KEY_RE.test(t))) return null;
    const find = pred => { const c = cells.find(pred); return c ? c[0] : null; };
    // A label ("Serial Number", "License key"), not a broken data line that
    // happens to say "Serial No:" - those have a count or key digits in them.
    const serial = find(([, t]) => XL_SERIAL_HDR_RE.test(t) && t.length <= 40 && !/\d/.test(t));
    if (serial === null) return null;
    const name = find(([c, t]) => c !== serial && XL_NAME_HDR_RE.test(t));
    const qty = find(([c, t]) => c !== serial && c !== name && XL_QTY_HDR_RE.test(t));
    return { serial, name, qty };
  }

  function parseGrid(sheets) {
    const res = newResult(), parsed = [];
    for (const sheet of sheets) {
      let hdr = null;
      for (const row of sheet.rows) {
        const r = row.r, cells = row.cells;
        const byCol = new Map(cells.map(([c, t]) => [c, t]));
        const qtyCell = hdr && hdr.qty ? byCol.get(hdr.qty) : undefined;
        const found = [];
        for (const [col, text] of cells) {
          for (const line of splitLines(text)) {
            const lic = xlLine(line);
            if (lic) { lic.raw = `${sheet.name}!${xlRef(col, r)}: ${lic.raw}`; found.push(lic); }
          }
        }
        // A count column only speaks for its row when the row holds one license.
        if (found.length === 1 && found[0].qty === null && qtyCell && XL_INT_RE.test(qtyCell)) found[0].qty = parseInt(qtyCell, 10);
        for (const lic of found) {
          lic.category = categorize(lic.name, lic.qty_label);
          if (lic.category === 'Instrument Control (3rd Party)' && lic.qty === null && lic.chk) lic.qty = 1; // same default as parseChecksumText
          delete lic.chk;
          lic.qty = defaultQty(lic.name, lic.qty);
          parsed.push(lic);
        }
        if (found.length) continue;
        const h = xlHeader(cells);
        if (h) { hdr = h; continue; }

        const raw = `${sheet.name}!${r}: ` + cells.map(([, t]) => t).join(' | ');
        let serialCol = null, serial = null;
        if (hdr && byCol.has(hdr.serial)) {
          const m = XL_SERIAL_CELL_RE.exec(byCol.get(hdr.serial));
          if (m) { serialCol = hdr.serial; serial = m[1]; }
        }
        if (serial === null) {
          for (const [col, text] of cells) {
            const m = XL_SERIAL_CELL_RE.exec(text);
            if (m && WATERS_KEY_RE.test(m[1])) { serialCol = col; serial = m[1]; break; }
          }
        }
        if (serial === null) {
          if (XL_MENTIONS_SERIAL_RE.test(raw) || cells.some(([, t]) => t.split(/\s+/).some(w => w && WATERS_KEY_RE.test(w)))) res.unparsed.push(raw);
          continue;
        }
        let name = null;
        if (hdr && hdr.name && hdr.name !== serialCol && byCol.has(hdr.name)) name = byCol.get(hdr.name);
        if (name === null) {
          const words = cells.filter(([c, t]) => c !== serialCol && !XL_INT_RE.test(t) && XL_LICENSE_WORD_RE.test(t)).map(([, t]) => t);
          name = words.length ? words.reduce((a, b) => (b.length > a.length ? b : a)) : null;
        }
        if (!name) { res.unparsed.push(raw); continue; }
        const q = hdr && hdr.qty ? byCol.get(hdr.qty) : undefined;
        const lic = { name, serial: cleanSerial(serial), qty: q && XL_INT_RE.test(q) ? parseInt(q, 10) : null, qty_label: null, raw };
        lic.category = categorize(lic.name, null);
        lic.qty = defaultQty(lic.name, lic.qty);
        parsed.push(lic);
      }
    }
    return finalize(res, parsed);
  }

  /* ---- .xlsx -> cell grid (mirrors xlsx_reader.py; same regexes, same output) ---- */
  const MAX_PART_BYTES = 50 * 1024 * 1024, MAX_COL = 16384;
  const P = '(?:[A-Za-z_][\\w.-]*:)?';
  const xr = (src, f) => new RegExp(src.replace(/~/g, P), f);
  const SHEET_RE = xr('<~sheet\\b([^>]*)>', 'g');
  const REL_RE = xr('<~Relationship\\b([^>]*)>', 'g');
  const SI_RE = xr('<~si\\b[^>]*?(?:/>|>(.*?)</~si>)', 'gs');
  const T_RE = xr('<~t\\b[^>]*?(?:/>|>(.*?)</~t>)', 'gs');
  const RPH_RE = xr('<~rPh\\b.*?</~rPh>', 'gs');
  const ROW_RE = xr('<~row\\b([^>]*?)(?:/>|>(.*?)</~row>)', 'gs');
  const CELL_RE = xr('<~c\\b([^>]*?)(?:/>|>(.*?)</~c>)', 'gs');
  const V_RE = xr('<~v\\b[^>]*?(?:/>|>(.*?)</~v>)', 's');
  const IS_RE = xr('<~is\\b[^>]*?(?:/>|>(.*?)</~is>)', 's');
  const REF_RE = /^([A-Za-z]{1,3})(\d+)$/;
  const ENTITY_RE = /&(#[xX][0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g;
  const OOXML_ESC_RE = /_x([0-9a-fA-F]{4})_/g;
  const ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

  function xmlAttr(attrs, name) {
    const m = new RegExp('(?:^|\\s)' + P + name + '\\s*=\\s*"([^"]*)"').exec(attrs) || new RegExp('(?:^|\\s)' + P + name + "\\s*=\\s*'([^']*)'").exec(attrs);
    return m ? unescapeXml(m[1]) : null;
  }
  function unescapeXml(s) {
    s = s.replace(ENTITY_RE, (_, e) => {
      if (e[0] !== '#') return ENTITIES[e];
      const cp = (e[1] === 'x' || e[1] === 'X') ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return cp > 0 && cp <= 0x10FFFF ? String.fromCodePoint(cp) : '';
    });
    return s.replace(OOXML_ESC_RE, (_, h) => String.fromCharCode(parseInt(h, 16)));
  }
  const textOf = xml => [...xml.replace(RPH_RE, '').matchAll(T_RE)].map(m => unescapeXml(m[1] || '')).join('');
  const colIndex = letters => [...letters.toUpperCase()].reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0);
  function resolveTarget(target) {
    if (target.startsWith('/')) return target.slice(1);
    const parts = [];
    for (const p of ('xl/' + target).split('/')) { if (p === '..') parts.pop(); else if (p && p !== '.') parts.push(p); }
    return parts.join('/');
  }

  function parseSheetXml(xml, shared) {
    const rows = []; let nextRow = 1;
    for (const rm of xml.matchAll(ROW_RE)) {
      const ra = xmlAttr(rm[1], 'r');
      const rnum = ra && /^\d+$/.test(ra) ? parseInt(ra, 10) : nextRow;
      nextRow = rnum + 1;
      const cells = []; let nextCol = 1;
      for (const cm of (rm[2] || '').matchAll(CELL_RE)) {
        const attrs = cm[1], inner = cm[2] || '';
        const ref = REF_RE.exec(xmlAttr(attrs, 'r') || '');
        const col = ref ? colIndex(ref[1]) : nextCol;
        nextCol = col + 1;
        if (col > MAX_COL) continue;
        const t = xmlAttr(attrs, 't') || 'n';
        let text;
        if (t === 'inlineStr') { const im = IS_RE.exec(inner); text = im ? textOf(im[1] || '') : ''; }
        else {
          const vm = V_RE.exec(inner), raw = vm ? (vm[1] || '') : '';
          text = unescapeXml(raw);
          if (t === 's') { const idx = raw.trim(); text = /^\d+$/.test(idx) && parseInt(idx, 10) < shared.length ? shared[parseInt(idx, 10)] : ''; }
          else if (t === 'b') text = raw.trim() === '1' ? 'TRUE' : raw.trim() === '0' ? 'FALSE' : raw;
        }
        text = text.trim();
        if (text) cells.push([col, text]);
      }
      if (cells.length) rows.push({ r: rnum, cells });
    }
    return rows;
  }

  // Minimal zip reader: central directory -> stored or deflated entries,
  // inflated with the platform's DecompressionStream (browsers and Node 18+).
  function zipEntries(u8) {
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    if (eocd < 0) throw new Error("Could not read that Excel file - it isn't a valid .xlsx workbook.");
    const count = dv.getUint16(eocd + 10, true); let p = dv.getUint32(eocd + 16, true);
    const dec = new TextDecoder(), out = new Map();
    for (let i = 0; i < count; i++) {
      if (p + 46 > u8.length || dv.getUint32(p, true) !== 0x02014b50) throw new Error("Could not read that Excel file - it isn't a valid .xlsx workbook.");
      const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
      const nlen = dv.getUint16(p + 28, true), elen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true), loff = dv.getUint32(p + 42, true);
      out.set(dec.decode(u8.subarray(p + 46, p + 46 + nlen)), { method, csize, loff });
      p += 46 + nlen + elen + clen;
    }
    return out;
  }
  async function zipRead(u8, entries, name) {
    const e = entries.get(name);
    if (!e) return null;
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    const start = e.loff + 30 + dv.getUint16(e.loff + 26, true) + dv.getUint16(e.loff + 28, true);
    const comp = u8.subarray(start, start + e.csize);
    let data;
    if (e.method === 0) data = comp;
    else if (e.method === 8) {
      const reader = new Blob([comp]).stream().pipeThrough(new DecompressionStream('deflate-raw')).getReader();
      const chunks = []; let n = 0;
      for (;;) {
        const { done, value } = await reader.read();
        if (done) break;
        n += value.length;
        if (n > MAX_PART_BYTES) { reader.cancel(); throw new Error('That workbook is too large to read.'); }
        chunks.push(value);
      }
      data = new Uint8Array(n); let o = 0; for (const c of chunks) { data.set(c, o); o += c.length; }
    } else throw new Error("Could not read that Excel file - it isn't a valid .xlsx workbook.");
    if (data.length > MAX_PART_BYTES) throw new Error('That workbook is too large to read.');
    return new TextDecoder().decode(data);
  }

  async function readXlsxGrid(bytes) {
    const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    const entries = zipEntries(u8);
    const workbook = await zipRead(u8, entries, 'xl/workbook.xml');
    if (workbook === null) throw new Error('Not an Excel workbook (.xlsx).');
    const rels = {}; let sharedPath = 'xl/sharedStrings.xml';
    for (const m of ((await zipRead(u8, entries, 'xl/_rels/workbook.xml.rels')) || '').matchAll(REL_RE)) {
      const rid = xmlAttr(m[1], 'Id'), target = xmlAttr(m[1], 'Target'), typ = xmlAttr(m[1], 'Type') || '';
      if (rid && target) { rels[rid] = resolveTarget(target); if (typ.endsWith('/sharedStrings')) sharedPath = rels[rid]; }
    }
    const shared = [...((await zipRead(u8, entries, sharedPath)) || '').matchAll(SI_RE)].map(m => textOf(m[1] || ''));
    const sheets = [];
    for (const m of workbook.matchAll(SHEET_RE)) {
      const path = rels[xmlAttr(m[1], 'id') || ''];
      if (!path) continue;
      const xml = await zipRead(u8, entries, path);
      if (xml === null) continue;
      sheets.push({ name: xmlAttr(m[1], 'name') || 'Sheet', rows: parseSheetXml(xml, shared) });
    }
    return sheets;
  }

  // The offline page keeps whatever it read: text (PDF/txt) or an xlsx grid.
  const parseSource = src => (typeof src === 'string' ? parseText(src) : parseGrid(src));

  function parseText(text) {
    if (text.includes('Option Properly Installed') || text.includes('Waters File Verification')) return parseChecksumText(text);
    if (text.includes('Serial No:')) return parseWizardText(text);
    throw new Error('Not a Licensing Wizard PDF, a Checksum .txt report or an Excel .xlsx workbook.');
  }
  function removeSqt(res) {
    const keep = [];
    for (const l of res.licenses) { if (l.name.toLowerCase().includes('sqt')) res.removed.push({ raw: l.raw, reason: 'SQT removed (checkbox)' }); else keep.push(l); }
    res.licenses = keep; return res;
  }
  function removeZeroQty(res) {
    const keep = [];
    for (const l of res.licenses) { if (l.qty === 0) res.removed.push({ raw: l.raw, reason: 'Zero quantity removed (checkbox)' }); else keep.push(l); }
    res.licenses = keep; return res;
  }

  /* ---- pdf.js text items -> layout lines (approximates `pdftotext -layout`) ---- */
  function itemsToLines(items) {
    const rows = [];
    for (const it of items) {
      if (!it.str || !it.transform) continue;
      const y = Math.round(it.transform[5]), x = it.transform[4];
      let row = rows.find(r => Math.abs(r.y - y) <= 2);
      if (!row) { row = { y, parts: [] }; rows.push(row); }
      row.parts.push({ x, s: it.str });
    }
    rows.sort((a, b) => b.y - a.y);
    return rows.map(r => r.parts.sort((a, b) => a.x - b.x).map(p => p.s).join(' ').replace(/\s+/g, ' ').trim());
  }

  /* ---- CSV (portal format) ---- */
  // Mirrors excel.inert(): Excel and Sheets both evaluate a field starting
  // with =, + , - or @ on import, and everything here came out of a Waters
  // report rather than out of this file. A leading apostrophe is the
  // standard escape that keeps the cell inert text.
  const inert = v => (typeof v === 'string' && '=+-@'.includes(v.charAt(0))) ? "'" + v : v;
  function buildCsv(res, details) {
    const out = ['Serial_Numbers'];
    const seen = new Set();
    for (const l of res.licenses) if (!seen.has(l.serial)) { seen.add(l.serial); out.push(csvCell(inert(l.serial))); }
    // Serials only: the Waters portal rejects a file with anything below the
    // list. Details live in the filename and the workbook (mirrors excel.build_csv).
    return out.join('\r\n') + '\r\n';
  }
  const csvCell = v => /[",\r\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;

  /* ---- XLSX: minimal writer (stored zip, inline strings, small style set) ---- */
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const colL = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - 1 - m) / 26; } return s; };
  // style ids: 0 normal, 1 header, 2 group (fill+bold), 3 group-cell (fill), 4 bold, 5 muted, 6 bordered, 7 centered bordered, 8 group centered
  const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="4"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font><font><sz val="11"/><color rgb="FF8A94A6"/><name val="Calibri"/></font></fonts>
<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F3A5F"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE8EEF6"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFC9D1D9"/></left><right style="thin"><color rgb="FFC9D1D9"/></right><top style="thin"><color rgb="FFC9D1D9"/></top><bottom style="thin"><color rgb="FFC9D1D9"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="9">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;

  function sheetXml(rows, opts) {
    // rows: array of arrays of {v, s} | primitive | null
    const widths = {};
    let xml = '';
    rows.forEach((row, ri) => {
      let cells = '';
      row.forEach((cell, ci) => {
        if (cell === null || cell === undefined) return;
        const v = (cell && typeof cell === 'object') ? cell.v : cell, s = (cell && typeof cell === 'object' && cell.s) || 0;
        if (v === null || v === undefined || v === '') { if (s) cells += `<c r="${colL(ci + 1)}${ri + 1}" s="${s}"/>`; return; }
        widths[ci] = Math.max(widths[ci] || 0, String(v).length);
        const ref = `${colL(ci + 1)}${ri + 1}`, sa = s ? ` s="${s}"` : '';
        cells += typeof v === 'number' ? `<c r="${ref}"${sa}><v>${v}</v></c>` : `<c r="${ref}"${sa} t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
      });
      xml += `<row r="${ri + 1}">${cells}</row>`;
    });
    const maxW = opts.maxWidth || 60;
    const cols = Object.keys(widths).map(ci => `<col min="${+ci + 1}" max="${+ci + 1}" width="${Math.max(10, Math.min(maxW, widths[ci] + 2))}" customWidth="1"/>`).join('');
    const freeze = opts.freezeRow ? `<sheetViews><sheetView workbookViewId="0"><pane ySplit="${opts.freezeRow}" topLeftCell="A${opts.freezeRow + 1}" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>` : '';
    const filter = opts.filter ? `<autoFilter ref="${opts.filter}"/>` : '';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${freeze}${cols ? `<cols>${cols}</cols>` : ''}<sheetData>${xml}</sheetData>${filter}</worksheet>`;
  }

  function buildWorkbook(res, details) {
    details = details || [];
    const lic = [];
    for (const [label, value] of details) lic.push([{ v: label, s: 4 }, inert(value)]);
    if (details.length) lic.push([]);
    const hdr = lic.length + 1;
    lic.push(['Category', 'License', 'Quantity', 'Quantity Type', 'Serial No'].map(v => ({ v, s: 1 })));
    let last = null;
    for (const l of res.licenses) {
      const first = l.category !== last; last = l.category;
      lic.push([
        { v: l.category, s: first ? 2 : 5 }, { v: inert(l.name), s: first ? 3 : 6 },
        { v: l.qty, s: first ? 8 : 7 }, { v: inert(l.qty_label), s: first ? 3 : 6 },
        { v: inert(l.serial), s: first ? 3 : 6 },
      ]);
    }
    const licXml = sheetXml(lic, { freezeRow: hdr, filter: `A${hdr}:E${lic.length}` });

    const sum = [];
    for (const [label, value] of details) sum.push([{ v: label, s: 4 }, inert(value)]);
    sum.push([{ v: 'Installation', s: 4 }, inert(res.installation || '')], [{ v: 'Date Printed', s: 4 }, inert(res.printed || '')],
      [{ v: 'Licenses (after cleanup)', s: 4 }, res.licenses.length], [{ v: 'Lines removed', s: 4 }, res.removed.length], []);
    sum.push(['Category', 'Licenses', 'Total Quantity'].map(v => ({ v, s: 1 })));
    for (const c of CATEGORY_ORDER) {
      const ls = res.licenses.filter(l => l.category === c);
      if (ls.length) sum.push([{ v: c, s: 4 }, ls.length, ls.reduce((n, l) => n + (l.qty || 0), 0)]);
    }
    const sumXml = sheetXml(sum, {});

    const rem = [['Reason', 'Original Line'].map(v => ({ v, s: 1 }))];
    for (const r of res.removed) rem.push([r.reason, inert(r.raw)]);
    for (const u of res.unparsed) rem.push(['Unrecognized line (not a license)', inert(u)]);
    const remXml = sheetXml(rem, { freezeRow: 1, maxWidth: 110 });

    const files = [
      ['[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`],
      ['_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
      ['xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Licenses" sheetId="1" r:id="rId1"/><sheet name="Summary" sheetId="2" r:id="rId2"/><sheet name="Removed" sheetId="3" r:id="rId3"/></sheets><definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Licenses!$A$${hdr}:$E$${lic.length}</definedName></definedNames></workbook>`],
      ['xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`],
      ['xl/styles.xml', STYLES],
      ['xl/worksheets/sheet1.xml', licXml], ['xl/worksheets/sheet2.xml', sumXml], ['xl/worksheets/sheet3.xml', remXml],
    ];
    return zipStore(files);
  }

  /* ---- tiny ZIP (store only) ---- */
  const CRC_TABLE = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  function crc32(u8) { let c = 0xFFFFFFFF; for (let i = 0; i < u8.length; i++) c = CRC_TABLE[(c ^ u8[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  function zipStore(files) {
    const enc = new TextEncoder(), parts = [], central = []; let offset = 0;
    const u16 = n => [n & 255, (n >> 8) & 255], u32 = n => [n & 255, (n >> 8) & 255, (n >> 16) & 255, (n >>> 24) & 255];
    for (const [name, content] of files) {
      const nameB = enc.encode(name), data = enc.encode(content), crc = crc32(data);
      const local = new Uint8Array([...u32(0x04034b50), ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0), ...u32(crc), ...u32(data.length), ...u32(data.length), ...u16(nameB.length), ...u16(0), ...nameB]);
      parts.push(local, data);
      central.push(new Uint8Array([...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0), ...u32(crc), ...u32(data.length), ...u32(data.length), ...u16(nameB.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(offset), ...nameB]));
      offset += local.length + data.length;
    }
    const cdSize = central.reduce((n, c) => n + c.length, 0);
    const end = new Uint8Array([...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(files.length), ...u16(files.length), ...u32(cdSize), ...u32(offset), ...u16(0)]);
    const total = offset + cdSize + end.length, out = new Uint8Array(total); let p = 0;
    for (const c of [...parts, ...central, end]) { out.set(c, p); p += c.length; }
    return out;
  }

  const safe = s => String(s || '').trim().replace(/[^A-Za-z0-9.-]+/g, '_').replace(/^_+|_+$/g, '');
  function baseName(fieldValues, originalName) {
    const parts = fieldValues.map(safe).filter(Boolean);
    if (parts.length) return parts.join('_');
    return safe(originalName.replace(/\.(pdf|txt|xlsx|xlsm)$/i, '')) || 'licenses';
  }

  const api = { parseText, parseWizardText, parseChecksumText, parseGrid, readXlsxGrid, parseSource, removeSqt, removeZeroQty, itemsToLines, buildCsv, buildWorkbook, baseName, CATEGORY_ORDER };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.ESLC = api;
})(typeof window !== 'undefined' ? window : globalThis);
