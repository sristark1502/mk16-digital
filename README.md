# MK16 Digital Report

Digital version of the EMR Tap Changers MK 16 (OLTC Service Report) form.
Engineers fill the details on the page, then download the report as **PDF** or **Word (.docx)** in the company format.

- Works in any browser (phone or laptop). Nothing is sent to a server: the details stay on the device (saved reports are kept in the browser).
- Word file: built from the company MK16 Word file (same page, fonts, widths and row heights).
- PDF: the same form drawn page by page (more than one page when the report is long).
- The company font is Bahnschrift SemiBold (Windows). On devices without it the PDF uses a similar free font (Barlow).

Rebuild the Word template data: `node tools/build-template.js` (kept in the working folder, not needed to run the page).
