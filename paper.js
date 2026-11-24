/* The paper: draws the MK16 form exactly like the company file (letter page, same margins, same column widths, row heights and 10 pt font),
   splits it into pages when the report is longer than one page, and turns the pages into a PDF. */
(function () {
  const PAGE_W = 612, PAGE_H = 792;                 // pt (US Letter, same as the company file)
  const TABLE_LEFT = 57.75, TABLE_TOP = 62.35;      // pt, measured from the company file
  const BOTTOM = PAGE_H - 18;                       // bottom margin 360 twips
  const TAIL_H = 13;                                // the "Checked By Service engineer" line
  const PX = 96 / 72;

  const el = (tag, cls, txt) => { const e = document.createElement(tag); if (cls) e.className = cls; if (txt != null) e.textContent = txt; return e; };
  const pt = (tw) => (tw / 20) + 'pt';

  function rowEl(r) {
    const row = el('div', 'mkrow');
    row.style.minHeight = (r.h / 20 + 0.5) + 'pt';
    let x = 0;
    r.cells.forEach((c) => {
      const cell = el('div', 'mkcell');
      cell.style.width = pt(c.w);
      const para = el('div', 'mkpara' + (c.center ? ' ctr' : '') + (c.bold ? ' b' : ''));
      if (c.before) para.style.paddingTop = pt(c.before);
      let tabIndex = -1;
      c.runs.forEach((run, i) => {
        if (run.tab) { tabIndex = i; para.appendChild(el('span', 'mktab')); return; }
        const s = el('span', run.b ? 'b' : '', run.t); para.appendChild(s);
      });
      para.dataset.tab = c.tabs ? c.tabs[0] : '';
      cell.appendChild(para); row.appendChild(cell);
      x += c.w;
    });
    return row;
  }

  // a tab in Word: the next text starts at the tab position (or at the next default tab when the first text is longer)
  function applyTabs(root) {
    root.querySelectorAll('.mkpara').forEach((p) => {
      const sp = p.querySelector('.mktab'); if (!sp) return;
      const pos = Number(p.dataset.tab) / 20 * PX;                 // px
      const first = p.firstElementChild; const w = first ? first.getBoundingClientRect().width : 0;
      const def = 36 * PX;                                          // default tab stop 720 twips
      const target = w <= pos ? pos : Math.ceil((w + 0.5) / def) * def;
      sp.style.display = 'inline-block'; sp.style.width = Math.max(1, target - w) + 'px';
    });
  }

  function paperShell(pageNo) {
    const paper = el('div', 'paper');
    const head = el('div', 'mkhead', 'MK 16'); paper.appendChild(head);
    const tbl = el('div', 'mktable'); paper.appendChild(tbl);
    return { paper, tbl };
  }

  // builds every page of the report inside `host` (the page DOM must be attached so that the text can be measured)
  function buildPapers(d, host) {
    const rows = MK16.rows(d);
    // 1. measure each row at its real size
    const probe = el('div', 'mktable probe'); const holder = el('div', 'paper measuring'); holder.appendChild(probe); host.appendChild(holder);
    const els = rows.map((r) => { const e = rowEl(r); probe.appendChild(e); return e; });
    applyTabs(probe);
    const heights = els.map((e) => e.getBoundingClientRect().height / PX);   // pt
    host.removeChild(holder);
    // 2. fill the pages
    const pages = []; let cur = [], used = 0;
    const limit = BOTTOM - TABLE_TOP - 0.5;
    heights.forEach((h, i) => { if (used + h > limit && cur.length) { pages.push({ rows: cur, h: used }); cur = []; used = 0; } cur.push(i); used += h; });
    pages.push({ rows: cur, h: used });
    // the closing line goes on the last page when it fits, else on a page of its own
    const lastUsed = pages[pages.length - 1].h;
    const tailOnOwn = TABLE_TOP + 0.5 + lastUsed + TAIL_H > BOTTOM;
    const out = [];
    pages.forEach((pg, k) => {
      const { paper, tbl } = paperShell(k + 1);
      pg.rows.forEach((i) => tbl.appendChild(rowEl(rows[i])));
      if (!tailOnOwn && k === pages.length - 1) { const t = el('div', 'mktail', MK16.tailText(d)); t.style.top = (TABLE_TOP + 0.5 + pg.h) + 'pt'; paper.appendChild(t); }
      out.push(paper);
    });
    if (tailOnOwn) { const { paper } = paperShell(pages.length + 1); const t = el('div', 'mktail', MK16.tailText(d)); t.style.top = (TABLE_TOP + 0.5) + 'pt'; paper.appendChild(t); out.push(paper); }
    out.forEach((p) => host.appendChild(p));
    out.forEach((p) => applyTabs(p));
    return out;
  }

  async function fontsReady() {
    try { await document.fonts.load('10pt "Bahnschrift SemiBold"'); await document.fonts.load('700 10pt "MK Fallback"'); await document.fonts.load('400 10pt "MK Fallback"'); await document.fonts.ready; } catch (e) { /* fonts are best effort */ }
  }

  async function buildPdf(d) {
    await fontsReady();
    const host = el('div', 'pdfhost'); document.body.appendChild(host);
    try {
      const papers = buildPapers(d, host);
      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF({ unit: 'pt', format: 'letter', compress: true });
      for (let i = 0; i < papers.length; i++) {
        const canvas = await html2canvas(papers[i], { scale: 3, backgroundColor: '#ffffff', logging: false, useCORS: true, width: PAGE_W * PX, height: PAGE_H * PX, windowWidth: 1400 });
        if (i) pdf.addPage('letter', 'portrait');
        pdf.addImage(canvas.toDataURL('image/jpeg', 0.93), 'JPEG', 0, 0, PAGE_W, PAGE_H, undefined, 'FAST');
      }
      return pdf.output('blob');
    } finally { host.remove(); }
  }

  window.MK16Paper = { buildPapers, buildPdf, fontsReady, PAGE_W, PAGE_H };
})();
