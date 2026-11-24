/* MK16 format: ONE description of the company form (rows, cells, text, widths, heights) used for both the Word file and the PDF.
   Everything here is copied from the company's own MK16 Word file: Bahnschrift SemiBold 10 pt, row heights, column widths, tab positions.
   Units: twips (1/20 pt) like Word. */
(function () {
  const COLS = [3755, 2540, 666, 3479];            // table grid
  const W = { full: 10440, a: 6295, b: 4145, c1: 3755, c2: 3206, c3: 3479 };
  const TAB_DATE = 6269;                            // where "Date :" starts in the Reference row (measured from the company file)
  const TAB_SLNO = 3226;                            // where "OLTC Sl.No :" starts in the OLTC row
  const MIN_DESC = 2, MIN_POINTS = 11, MIN_REMARKS = 5;   // the blank form has 2 incident lines, 11 service lines, 5 closing lines

  const fmtDate = (v) => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v || ''); return m ? m[3] + '-' + m[2] + '-' + m[1] : (v || ''); };
  const lines = (s) => String(s || '').replace(/\r/g, '').split('\n');

  // a cell: { w, span, runs:[{t, b}|{tab:true}], center, before, tabs:[pos] }
  const cell = (w, span, runs, o) => Object.assign({ w, span, runs }, o || {});
  const T = (t, b) => ({ t: t, b: !!b });
  const TAB = { tab: true };

  function rows(d) {
    d = d || {};
    const R = [];
    const row = (h, cells) => R.push({ h, cells });
    row(382, [cell(W.full, 4, [T('OLTC Service Report', 1)], { center: 1 })]);
    row(382, [cell(W.full, 4, [T('EMR Tap Changers (p) Ltd', 1)], { center: 1 })]);
    row(382, [cell(W.full, 4, [T('Reference :', 1), TAB, T('Date : ' + fmtDate(d.date), 1)], { tabs: [TAB_DATE], bold: 1 })]);
    row(382, [cell(W.full, 4, [T('Customer name : ' + (d.customer || ''))])]);
    row(611, [
      cell(W.c1, 1, [T('Contact Person: ' + (d.contactPerson || ''))], { before: 240 }),
      cell(W.c2, 2, [T('Contact N0: ' + (d.contactNo || ''))]),
      cell(W.c3, 1, [T('Designation: ' + (d.designation || ''))])
    ]);
    row(382, [cell(W.full, 4, [T('Site : ' + (d.site1 || ''))])]);
    row(382, [cell(W.full, 4, [T(d.site2 || '')])]);
    row(382, [cell(W.full, 4, [T('OLTC Specification : ' + (d.spec || ''), 1), TAB, T('OLTC Sl.No : ' + (d.slno || ''), 1)], { tabs: [TAB_SLNO], bold: 1 })]);
    row(382, [cell(W.a, 2, [T('Transformer Make: ' + (d.make || ''))]), cell(W.b, 2, [T('MVA & KV: ' + (d.mva || ''))])]);
    row(382, [cell(W.a, 2, [T('Date of installation: ' + fmtDate(d.dInst))]), cell(W.b, 2, [T('Date of Commissioning: ' + fmtDate(d.dComm))])]);
    row(382, [cell(W.full, 4, [T('')], { bold: 1 })]);
    row(469, [cell(W.full, 4, [T('Description of incident:', 1)])]);
    const desc = lines(d.incident); while (desc.length < MIN_DESC) desc.push('');
    desc.forEach((t) => row(437, [cell(W.full, 4, [T(t)])]));
    row(437, [cell(W.full, 4, [T('Service Details', 1)])]);
    const pts = (d.points || []).map((x) => String(x || '').trim()).filter(Boolean).map((x, i) => (i + 1) + '. ' + x);
    while (pts.length < MIN_POINTS) pts.push('');
    pts.forEach((t, i) => row(i === 0 ? 437 : 382, [cell(W.full, 4, [T(t)])]));
    row(382, [cell(W.full, 4, [T('')], { bold: 1 })]);
    const rem = lines(d.remarks); while (rem.length < MIN_REMARKS) rem.push('');
    rem.forEach((t, i) => row(i === rem.length - 1 ? 403 : 382, [cell(W.full, 4, [T(t)])]));
    return R;
  }

  const tailText = (d) => 'Checked By Service engineer' + (d && d.engineer && d.engineer.trim() ? ' : ' + d.engineer.trim() : '');

  /* ---------------- Word (.docx) ---------------- */
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const FONT = '<w:rFonts w:hint="default" w:ascii="Bahnschrift SemiBold" w:hAnsi="Bahnschrift SemiBold" w:cs="Bahnschrift SemiBold" w:eastAsiaTheme="majorEastAsia"/>';
  const rpr = (b) => '<w:rPr>' + FONT + (b ? '<w:b/>' : '') + '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>';
  function runXml(r, cellBold) {
    if (r.tab) return '<w:r>' + rpr(cellBold) + '<w:tab/></w:r>';
    if (!r.t) return '';
    const parts = String(r.t).split('\n');
    return '<w:r>' + rpr(r.b) + parts.map((p, i) => (i ? '<w:br/>' : '') + '<w:t xml:space="preserve">' + esc(p) + '</w:t>').join('') + '</w:r>';
  }
  function cellXml(c) {
    const tcPr = '<w:tcPr><w:tcW w:w="' + c.w + '" w:type="dxa"/>' + (c.span > 1 ? '<w:gridSpan w:val="' + c.span + '"/>' : '') + '<w:shd w:val="clear" w:color="auto" w:fill="auto"/><w:vAlign w:val="bottom"/></w:tcPr>';
    const pPr = '<w:pPr><w:pStyle w:val="9"/>' + (c.tabs ? '<w:tabs>' + c.tabs.map((p) => '<w:tab w:val="left" w:pos="' + p + '"/>').join('') + '</w:tabs>' : '') + (c.before ? '<w:spacing w:before="' + c.before + '"/>' : '') + (c.center ? '<w:jc w:val="center"/>' : '') + rpr(c.bold || c.center) + '</w:pPr>';
    return '<w:tc>' + tcPr + '<w:p>' + pPr + c.runs.map((r) => runXml(r, c.bold)).join('') + '</w:p></w:tc>';
  }
  function documentXml(d) {
    const T0 = window.MK16_TPL;
    const body = rows(d).map((r) => '<w:tr>' + T0.pex + '<w:trPr><w:trHeight w:val="' + r.h + '" w:hRule="atLeast"/></w:trPr>' + r.cells.map(cellXml).join('') + '</w:tr>').join('');
    // the closing paragraph of the company file: "Checked By Service engineer" (+ the engineer's name when typed)
    let tail = T0.tail;
    const name = d && d.engineer && d.engineer.trim();
    if (name) tail = tail.replace('<w:bookmarkStart', '<w:r>' + rpr(true) + '<w:t xml:space="preserve"> : ' + esc(name) + '</w:t></w:r><w:bookmarkStart');
    return T0.head + body + '</w:tbl>' + tail;
  }
  async function buildDocx(d) {
    const zip = await JSZip.loadAsync(window.MK16_TPL.rest, { base64: true });
    zip.file('word/document.xml', documentXml(d));
    return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', compression: 'DEFLATE' });
  }

  window.MK16 = { COLS, W, rows, tailText, fmtDate, buildDocx, documentXml, MIN_POINTS };
})();
