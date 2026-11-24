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
  function documentXml(d, extra) {
    const T0 = window.MK16_TPL;
    const body = rows(d).map((r) => '<w:tr>' + T0.pex + '<w:trPr><w:trHeight w:val="' + r.h + '" w:hRule="atLeast"/></w:trPr>' + r.cells.map(cellXml).join('') + '</w:tr>').join('');
    // the closing paragraph of the company file: "Checked By Service engineer" (+ the engineer's name when typed)
    let tail = T0.tail;
    const name = d && d.engineer && d.engineer.trim();
    if (name) tail = tail.replace('<w:bookmarkStart', '<w:r>' + rpr(true) + '<w:t xml:space="preserve"> : ' + esc(name) + '</w:t></w:r><w:bookmarkStart');
    if (extra) tail = tail.replace('<w:sectPr>', extra + '<w:sectPr>');
    return T0.head + body + '</w:tbl>' + tail;
  }
  /* ---- photos: every photo on its own page, as big as the page allows (same box in the PDF) ---- */
  const PHOTO = { x: 62.5, y: 70, w: 500, h: 640 };          // pt: where a photo may go on its page
  const photoFit = (p) => { const s = Math.min(PHOTO.w / p.w, PHOTO.h / p.h); return { w: p.w * s, h: p.h * s }; };
  const EMU = (pt) => Math.round(pt * 12700);
  function photoXml(p, rid, n) {
    const f = photoFit(p), cx = EMU(f.w), cy = EMU(f.h);
    const pic = '<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="' + cx + '" cy="' + cy + '"/><wp:docPr id="' + (100 + n) + '" name="Photo ' + n + '"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="' + (100 + n) + '" name="photo' + n + '.jpg"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="' + rid + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
    const cap = (p.cap || '').trim();
    return pic + (cap ? '<w:p><w:pPr><w:pStyle w:val="9"/><w:spacing w:before="120"/><w:jc w:val="center"/>' + rpr(false) + '</w:pPr>' + runXml({ t: cap, b: false }, false) + '</w:p>' : '');
  }
  const bytesOf = (url) => { const s = atob(String(url).split(',')[1] || ''), u = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) u[i] = s.charCodeAt(i); return u; };

  async function buildDocx(d) {
    const zip = await JSZip.loadAsync(window.MK16_TPL.rest, { base64: true });
    const photos = (d && d.photos) || [];
    let extra = '';
    if (photos.length) {
      let rels = '';
      photos.forEach((p, i) => {
        const n = i + 1, rid = 'rIdPhoto' + n;
        zip.file('word/media/photo' + n + '.jpg', bytesOf(p.data), { binary: true });
        rels += '<Relationship Id="' + rid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/photo' + n + '.jpg"/>';
        extra += photoXml(p, rid, n);
      });
      const rp = 'word/_rels/document.xml.rels';
      zip.file(rp, (await zip.file(rp).async('string')).replace('</Relationships>', rels + '</Relationships>'));
      let ct = await zip.file('[Content_Types].xml').async('string');
      if (!/Extension="jpg"/i.test(ct)) ct = ct.replace('<Default Extension="xml"', '<Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="xml"');
      zip.file('[Content_Types].xml', ct);
    }
    zip.file('word/document.xml', documentXml(d, extra));
    return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', compression: 'DEFLATE' });
  }

  window.MK16 = { COLS, W, rows, tailText, fmtDate, buildDocx, documentXml, MIN_POINTS, PHOTO, photoFit };
})();
