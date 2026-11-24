(function () {
  const $ = (s, r) => (r || document).querySelector(s), $$ = (s, r) => [...(r || document).querySelectorAll(s)];
  const KEYS = ['date', 'customer', 'contactPerson', 'contactNo', 'designation', 'site1', 'site2', 'spec', 'slno', 'make', 'mva', 'dInst', 'dComm', 'incident', 'remarks', 'engineer'];
  const today = () => { const t = new Date(), p = (n) => String(n).padStart(2, '0'); return t.getFullYear() + '-' + p(t.getMonth() + 1) + '-' + p(t.getDate()); };
  const defaults = () => ({ date: today(), make: 'T&R', points: [''] });   // "T&R" is already written in the company form
  let currentId = null;
  let photos = [];   // [{data: jpeg data-url, w, h, cap}]

  // browser storage can be blocked (private window): never let it break the form
  const store = {
    get(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; } }
  };

  // photos live in IndexedDB (too big for localStorage); the text of the report stays in localStorage
  const idb = {
    db: null,
    open() { return this.db ? Promise.resolve(this.db) : new Promise((ok, no) => { try { const r = indexedDB.open('mk16', 1); r.onupgradeneeded = () => r.result.createObjectStore('photos'); r.onsuccess = () => { this.db = r.result; ok(this.db); }; r.onerror = () => no(r.error); } catch (e) { no(e); } }); },
    async get(k) { try { const db = await this.open(); return await new Promise((ok) => { const q = db.transaction('photos').objectStore('photos').get(k); q.onsuccess = () => ok(q.result || []); q.onerror = () => ok([]); }); } catch (e) { return []; } },
    async set(k, v) { try { const db = await this.open(); await new Promise((ok, no) => { const t = db.transaction('photos', 'readwrite'); t.objectStore('photos').put(v, k); t.oncomplete = ok; t.onerror = () => no(t.error); }); return true; } catch (e) { return false; } },
    async forget(k) { try { const db = await this.open(); await new Promise((ok) => { const t = db.transaction('photos', 'readwrite'); t.objectStore('photos')['delete'](k); t.oncomplete = ok; t.onerror = ok; }); } catch (e) { /* ignore */ } }
  };
  let phTimer;
  const keepPhotos = () => { clearTimeout(phTimer); phTimer = setTimeout(() => idb.set('draft', photos), 400); };

  // a phone photo is 3-6 MB: it is shrunk to at most 2000 px on the long side (JPEG), enough for a sharp full page
  async function shrink(file) {
    const url = URL.createObjectURL(file);
    try {
      const img = new Image(); img.src = url; await img.decode();
      const s = Math.min(1, 2000 / Math.max(img.naturalWidth, img.naturalHeight)), w = Math.round(img.naturalWidth * s), h = Math.round(img.naturalHeight * s);
      const c = document.createElement('canvas'); c.width = w; c.height = h; const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, w, h); g.drawImage(img, 0, 0, w, h);
      return { data: c.toDataURL('image/jpeg', 0.85), w, h, cap: '' };
    } finally { URL.revokeObjectURL(url); }
  }
  function renderPhotos() {
    const box = $('#photos'); box.innerHTML = '';
    photos.forEach((p, i) => {
      const row = document.createElement('div'); row.className = 'ph';
      row.innerHTML = '<img alt=""><div class="side"><div class="pn"></div><input placeholder="Caption (optional)"><div class="acts"><button type="button" class="up" aria-label="Move up">↑</button><button type="button" class="dn" aria-label="Move down">↓</button><button type="button" class="x" aria-label="Remove photo">Remove</button></div></div>';
      row.querySelector('img').src = p.data; row.querySelector('.pn').textContent = 'Photo ' + (i + 1) + ' (own page after the form)';
      const cap = row.querySelector('input'); cap.value = p.cap || ''; cap.addEventListener('input', () => { p.cap = cap.value; keepPhotos(); });
      const move = (d) => { const j = i + d; if (j < 0 || j >= photos.length) return; [photos[i], photos[j]] = [photos[j], photos[i]]; renderPhotos(); keepPhotos(); };
      row.querySelector('.up').addEventListener('click', () => move(-1)); row.querySelector('.dn').addEventListener('click', () => move(1));
      row.querySelector('.x').addEventListener('click', () => { photos.splice(i, 1); renderPhotos(); keepPhotos(); });
      box.appendChild(row);
    });
  }
  async function addPhotos(files) {
    let bad = 0;
    for (const f of files) { try { photos.push(await shrink(f)); renderPhotos(); } catch (e) { bad++; } }
    keepPhotos(); if (bad) toast(bad + ' photo(s) could not be read');
  }

  /* ---------- form <-> data ---------- */
  function read() {
    const d = {};
    KEYS.forEach((k) => { d[k] = $('[data-k="' + k + '"]').value; });
    d.points = $$('#points textarea').map((t) => t.value);
    return d;
  }
  function growText(t) { t.style.height = 'auto'; t.style.height = Math.max(46, t.scrollHeight + 2) + 'px'; }
  function renderPoints(list) {
    const box = $('#points'); box.innerHTML = '';
    (list && list.length ? list : ['']).forEach((v, i) => {
      const row = document.createElement('div'); row.className = 'pt';
      row.innerHTML = '<div class="n"></div><textarea rows="1" placeholder="Service point"></textarea><button type="button" class="x" title="Remove this point" aria-label="Remove point">×</button>';
      const ta = row.querySelector('textarea'); ta.value = v; box.appendChild(row);
      ta.addEventListener('input', () => { growText(ta); draft(); });
      row.querySelector('.x').addEventListener('click', () => { if ($$('#points .pt').length > 1) { row.remove(); number(); draft(); } else { ta.value = ''; draft(); } });
      growText(ta);
    });
    number();
  }
  function number() { $$('#points .pt .n').forEach((n, i) => { n.textContent = i + 1; }); }
  function write(d) {
    KEYS.forEach((k) => { $('[data-k="' + k + '"]').value = d[k] || ''; });
    renderPoints(d.points);
  }
  const draft = () => store.set('mk16.draft', { id: currentId, data: read() });

  /* ---------- small helpers ---------- */
  let tt;
  function toast(t) { const e = $('#toast'); e.textContent = t; e.classList.add('on'); clearTimeout(tt); tt = setTimeout(() => e.classList.remove('on'), 2200); }
  const open = (id) => $(id).classList.add('on'), close = (id) => $(id).classList.remove('on');
  function fileBase(d) {
    const clean = (s) => String(s || '').replace(/[^\w\-]+/g, '_').replace(/^_+|_+$/g, '').slice(0, 40);
    return 'MK16_' + (clean(d.customer) || 'report') + '_' + (d.date || today());
  }
  function save(blob, name) {
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
  }

  /* ---------- save / reset / saved list ---------- */
  function title(d) { return [d.customer, d.site1].filter(Boolean).join(' · ') || 'Untitled report'; }
  function saveReport() {
    const d = read(), list = store.get('mk16.saves', []);
    const rec = { id: currentId || ('r' + Date.now()), title: title(d), savedAt: Date.now(), data: d };
    const i = list.findIndex((x) => x.id === rec.id);
    if (i >= 0) list[i] = rec; else list.unshift(rec);
    currentId = rec.id;
    idb.set('save:' + rec.id, photos);
    if (store.set('mk16.saves', list)) { draft(); toast('Saved ✓  (' + rec.title + ')'); } else toast('Could not save in this browser');
  }
  function resetForm() {
    if (!confirm('Clear everything and start a new report?')) return;
    currentId = null; photos = []; renderPhotos(); idb.forget('draft'); write(defaults()); draft(); window.scrollTo({ top: 0, behavior: 'smooth' }); toast('New blank report');
  }
  function renderSaved() {
    const box = $('#savedList'), list = store.get('mk16.saves', []);
    if (!list.length) { box.innerHTML = '<p>No saved reports yet. Fill a report and tap <b>Save</b>.</p>'; return; }
    box.innerHTML = '';
    list.forEach((r) => {
      const row = document.createElement('div'); row.className = 'saved';
      const when = new Date(r.savedAt).toLocaleString('en-IN', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' });
      row.innerHTML = '<div class="t"><b></b><small>Saved ' + when + '</small></div><button class="btn ok" type="button">Open</button><button class="btn bad" type="button">Delete</button>';
      row.querySelector('b').textContent = r.title;
      row.querySelector('.ok').addEventListener('click', async () => { currentId = r.id; write(r.data); photos = await idb.get('save:' + r.id); renderPhotos(); draft(); keepPhotos(); close('#ovList'); window.scrollTo({ top: 0 }); toast('Opened: ' + r.title); });
      row.querySelector('.bad').addEventListener('click', () => { if (!confirm('Delete "' + r.title + '"?')) return; store.set('mk16.saves', store.get('mk16.saves', []).filter((x) => x.id !== r.id)); idb.forget('save:' + r.id); if (currentId === r.id) currentId = null; renderSaved(); });
      box.appendChild(row);
    });
  }

  /* ---------- download ---------- */
  async function download(kind) {
    const st = $('#dlStatus'), btns = $$('#ovDown button'); btns.forEach((b) => { b.disabled = true; });
    const d = Object.assign(read(), { photos });
    try {
      st.textContent = kind === 'pdf' ? 'Making the PDF…' : 'Making the Word file…';
      if (kind === 'pdf') save(await MK16Paper.buildPdf(d), fileBase(d) + '.pdf');
      else save(await MK16.buildDocx(d), fileBase(d) + '.docx');
      st.textContent = 'Done ✓ — check your downloads';
      saveReport();
      setTimeout(() => { st.textContent = ''; close('#ovDown'); }, 1600);
    } catch (e) {
      console.error(e); st.textContent = 'Could not make the file: ' + (e && e.message || e);
    } finally { btns.forEach((b) => { b.disabled = false; }); }
  }
  async function preview() {
    const box = $('#prevBox'); box.innerHTML = ''; open('#ovPrev');
    await MK16Paper.fontsReady();
    const dd = Object.assign(read(), { photos });
    const papers = MK16Paper.buildPapers(dd, box).concat(photos.map((p) => { const e = MK16Paper.photoPaper(p); box.appendChild(e); return e; }));
    const s = Math.min(1, (box.clientWidth - 16) / (612 * 96 / 72));
    papers.forEach((p) => {
      const w = document.createElement('div'); w.className = 'pwrap';
      w.style.width = (612 * 96 / 72 * s) + 'px'; w.style.height = (792 * 96 / 72 * s) + 'px';
      p.style.transform = 'scale(' + s + ')'; box.replaceChild(w, p); w.appendChild(p);
    });
  }

  /* ---------- start ---------- */
  document.addEventListener('DOMContentLoaded', () => {
    const dr = store.get('mk16.draft', null);
    if (dr && dr.data) { currentId = dr.id || null; write(Object.assign(defaults(), dr.data)); } else write(defaults());
    idb.get('draft').then((p) => { if (p && p.length && !photos.length) { photos = p; renderPhotos(); } });
    $('#addPh').addEventListener('click', () => $('#phFile').click());
    $('#phFile').addEventListener('change', (e) => { const fs = [...e.target.files]; e.target.value = ''; addPhotos(fs); });
    $$('#f [data-k]').forEach((i) => { i.addEventListener('input', draft); i.addEventListener('change', draft); });
    $('#addPt').addEventListener('click', () => { const l = read().points; l.push(''); renderPoints(l); const t = $$('#points textarea').pop(); t.focus(); draft(); });
    $('#bSave').addEventListener('click', saveReport);
    $('#bReset').addEventListener('click', resetForm);
    $('#bList').addEventListener('click', () => { renderSaved(); open('#ovList'); });
    $('#bDown').addEventListener('click', () => { $('#dlStatus').textContent = ''; open('#ovDown'); });
    $('#gPdf').addEventListener('click', () => download('pdf'));
    $('#gDoc').addEventListener('click', () => download('docx'));
    $('#bPrev').addEventListener('click', preview);
    $('#dlClose').addEventListener('click', () => close('#ovDown'));
    $('#listClose').addEventListener('click', () => close('#ovList'));
    $('#prevClose').addEventListener('click', () => close('#ovPrev'));
    $$('.ov').forEach((o) => o.addEventListener('click', (e) => { if (e.target === o) o.classList.remove('on'); }));
    window.addEventListener('beforeunload', draft);
  });
})();
