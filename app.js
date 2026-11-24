(function () {
  const $ = (s, r) => (r || document).querySelector(s), $$ = (s, r) => [...(r || document).querySelectorAll(s)];
  const KEYS = ['date', 'customer', 'contactPerson', 'contactNo', 'designation', 'site1', 'site2', 'spec', 'slno', 'make', 'mva', 'dInst', 'dComm', 'incident', 'remarks', 'engineer'];
  const today = () => { const t = new Date(), p = (n) => String(n).padStart(2, '0'); return t.getFullYear() + '-' + p(t.getMonth() + 1) + '-' + p(t.getDate()); };
  const defaults = () => ({ date: today(), make: 'T&R', points: [''] });   // "T&R" is already written in the company form
  let currentId = null;

  // browser storage can be blocked (private window): never let it break the form
  const store = {
    get(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; } }
  };

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
    if (store.set('mk16.saves', list)) { draft(); toast('Saved ✓  (' + rec.title + ')'); } else toast('Could not save in this browser');
  }
  function resetForm() {
    if (!confirm('Clear everything and start a new report?')) return;
    currentId = null; write(defaults()); draft(); window.scrollTo({ top: 0, behavior: 'smooth' }); toast('New blank report');
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
      row.querySelector('.ok').addEventListener('click', () => { currentId = r.id; write(r.data); draft(); close('#ovList'); window.scrollTo({ top: 0 }); toast('Opened: ' + r.title); });
      row.querySelector('.bad').addEventListener('click', () => { if (!confirm('Delete "' + r.title + '"?')) return; store.set('mk16.saves', store.get('mk16.saves', []).filter((x) => x.id !== r.id)); if (currentId === r.id) currentId = null; renderSaved(); });
      box.appendChild(row);
    });
  }

  /* ---------- download ---------- */
  async function download(kind) {
    const st = $('#dlStatus'), btns = $$('#ovDown button'); btns.forEach((b) => { b.disabled = true; });
    const d = read();
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
    const papers = MK16Paper.buildPapers(read(), box);
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
