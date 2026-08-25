(() => {
  const drop = document.getElementById('drop');
  const input = document.getElementById('file');
  const errBox = document.getElementById('error');
  const result = document.getElementById('result');
  let currentFile = null;
  // Every checkbox toggle re-uploads and re-renders. Two quick toggles put
  // two requests in flight, and the slower one finishing last painted the
  // table for options the page no longer has ticked. Each request claims a
  // ticket; only the newest is allowed to render.
  let previewSeq = 0;

  const opts = { remove_sqt: document.getElementById('remove_sqt'), remove_zero: document.getElementById('remove_zero') };
  const fields = { company: document.getElementById('company'), support_id: document.getElementById('support_id'), sold_to: document.getElementById('sold_to'), order_number: document.getElementById('order_number') };
  const form = (file) => { const fd = new FormData(); fd.append('pdf', file); for (const [k, el] of Object.entries(opts)) if (el.checked) fd.append(k, '1'); for (const [k, el] of Object.entries(fields)) if (el.value.trim()) fd.append(k, el.value.trim()); return fd; };
  Object.values(opts).forEach(el => el.addEventListener('change', () => { if (currentFile) handle(currentFile); }));
  const showError = (m) => { errBox.textContent = m; errBox.hidden = !m; };

  drop.addEventListener('click', () => input.click());
  drop.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); } });
  input.addEventListener('change', () => input.files[0] && handle(input.files[0]));
  ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', (e) => { const f = e.dataTransfer.files[0]; if (f) handle(f); });
  document.addEventListener('paste', (e) => { const f = [...(e.clipboardData?.files || [])][0]; if (f) handle(f); });

  document.getElementById('reset').addEventListener('click', () => {
    currentFile = null; input.value = ''; result.hidden = true; drop.hidden = false; showError('');
  });
  document.querySelectorAll('[data-dl]').forEach(b => b.addEventListener('click', () => download(b)));

  async function handle(file) {
    showError('');
    if (!/\.(pdf|txt)$/i.test(file.name) && !['application/pdf', 'text/plain'].includes(file.type)) return showError('Please choose a PDF or .txt file.');
    drop.classList.add('busy');
    const fd = form(file);
    const ticket = ++previewSeq;
    try {
      const r = await fetch('/api/preview', { method: 'POST', body: fd });
      const data = await r.json().catch(() => ({ error: 'Unexpected server response.' }));
      if (ticket !== previewSeq) return;          // a newer toggle already won
      if (!r.ok) throw new Error(data.error || `HTTP ${r.status}`);
      currentFile = file;
      render(data);
    } catch (e) {
      if (ticket === previewSeq) showError(e.message);
    } finally {
      if (ticket === previewSeq) drop.classList.remove('busy');
    }
  }

  function render(d) {
    for (const [k, v] of Object.entries(d.prefill || {})) if (v && fields[k] && !fields[k].value.trim()) fields[k].value = v;
    document.getElementById('install').textContent = d.installation || currentFile.name;
    const qty = d.licenses.reduce((n, l) => n + (l.qty || 0), 0);
    document.getElementById('meta').textContent =
      `${d.licenses.length} licenses · ${qty} total seats/units` + (d.printed ? ` · printed ${d.printed}` : '');
    const tb = document.querySelector('#table tbody'); tb.innerHTML = '';
    let last = null;
    for (const l of d.licenses) {
      const tr = document.createElement('tr');
      if (l.category !== last) { tr.className = 'first'; last = l.category; }
      tr.innerHTML = `<td class="cat">${esc(l.category)}</td><td>${esc(l.name)}</td>` +
        `<td class="qty">${l.qty ?? '—'}</td><td class="serial">${esc(l.serial)}</td>`;
      tb.appendChild(tr);
    }
    const removed = [...d.removed, ...d.unparsed.map(u => ({ reason: 'Unrecognized line', raw: u }))];
    const rb = document.querySelector('#removed tbody'); rb.innerHTML = '';
    for (const r of removed) {
      const tr = document.createElement('tr');
      tr.innerHTML = `<td>${esc(r.reason)}</td><td>${esc(r.raw)}</td>`;
      rb.appendChild(tr);
    }
    document.getElementById('removedcount').textContent = removed.length;
    document.getElementById('removedbox').hidden = removed.length === 0;
    drop.hidden = true; result.hidden = false;
  }

  async function download(btn) {
    if (!currentFile) return;
    const url = btn.dataset.dl, fallback = url.includes('csv') ? 'Simple_ESLC.csv' : 'Detailed_ESLC.xlsx';
    const fd = form(currentFile);
    btn.disabled = true;
    try {
      const r = await fetch(url, { method: 'POST', body: fd });
      if (!r.ok) { const d = await r.json().catch(() => ({})); throw new Error(d.error || `HTTP ${r.status}`); }
      const name = (r.headers.get('Content-Disposition') || '').match(/filename="?([^";]+)"?/)?.[1] || fallback;
      const blobUrl = URL.createObjectURL(await r.blob());
      const a = Object.assign(document.createElement('a'), { href: blobUrl, download: name });
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(blobUrl), 5000);
    } catch (e) { showError(e.message); } finally { btn.disabled = false; }
  }

  fetch('/api/version').then(r => r.json()).then(d => { document.getElementById('version').textContent = d.version; const o = document.getElementById('offline-link'); if (o) o.download = `EmpowerSmartLicenseCleaner_v${d.version}.html`; }).catch(() => {});

  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
})();
