/**
 * docx-lite.js — 輕量 DOCX 產生器（純 JavaScript，無外部依賴）
 * 使用 ZIP STORE 格式打包 Office Open XML (DOCX)
 */
(function (global) {
  'use strict';

  // ── CRC-32 ─────────────────────────────────────────────────────────
  const CRC_TABLE = (function () {
    const t = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let j = 0; j < 8; j++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[i] = c;
    }
    return t;
  })();

  function crc32(data) {
    let crc = 0xFFFFFFFF;
    for (let i = 0; i < data.length; i++)
      crc = CRC_TABLE[(crc ^ data[i]) & 0xFF] ^ (crc >>> 8);
    return (crc ^ 0xFFFFFFFF) >>> 0;
  }

  // ── ZIP（STORE，不壓縮）───────────────────────────────────────────
  function makeZip(files) {
    const te = new TextEncoder();
    const locals = [], centrals = [];
    let offset = 0;

    for (const file of files) {
      const name = (typeof file.name === 'string') ? te.encode(file.name) : file.name;
      const data = file.data;
      const crc  = crc32(data);
      const size = data.length;

      const lh = new Uint8Array(30 + name.length);
      const lv = new DataView(lh.buffer);
      lv.setUint32( 0, 0x04034B50, true);
      lv.setUint16( 4, 20, true); lv.setUint16(6, 0, true); lv.setUint16(8, 0, true);
      lv.setUint16(10, 0,    true); lv.setUint16(12, 0, true);
      lv.setUint32(14, crc,  true); lv.setUint32(18, size, true); lv.setUint32(22, size, true);
      lv.setUint16(26, name.length, true); lv.setUint16(28, 0, true);
      lh.set(name, 30);
      locals.push({ lh, data, offset, crc, size, nameLen: name.length });

      const ch = new Uint8Array(46 + name.length);
      const cv = new DataView(ch.buffer);
      cv.setUint32( 0, 0x02014B50, true);
      cv.setUint16( 4, 20, true); cv.setUint16(6, 20, true); cv.setUint16(8, 0, true);
      cv.setUint16(10, 0, true); cv.setUint16(12, 0, true); cv.setUint16(14, 0, true);
      cv.setUint32(16, crc,  true); cv.setUint32(20, size, true); cv.setUint32(24, size, true);
      cv.setUint16(28, name.length, true); cv.setUint16(30, 0, true); cv.setUint16(32, 0, true);
      cv.setUint16(34, 0, true); cv.setUint16(36, 0, true); cv.setUint32(38, 0, true);
      cv.setUint32(42, offset, true);
      ch.set(name, 46);
      centrals.push(ch);

      offset += lh.length + size;
    }

    const cdOffset = offset;
    const cdSize = centrals.reduce((s, c) => s + c.length, 0);
    const eocd = new Uint8Array(22);
    const ev = new DataView(eocd.buffer);
    ev.setUint32( 0, 0x06054B50, true); ev.setUint16(4, 0, true); ev.setUint16(6, 0, true);
    ev.setUint16( 8, files.length, true); ev.setUint16(10, files.length, true);
    ev.setUint32(12, cdSize, true); ev.setUint32(16, cdOffset, true); ev.setUint16(20, 0, true);

    const parts = [];
    for (const f of locals) { parts.push(f.lh); parts.push(f.data); }
    for (const c of centrals) parts.push(c);
    parts.push(eocd);

    const total = parts.reduce((s, p) => s + p.length, 0);
    const out = new Uint8Array(total);
    let pos = 0;
    for (const p of parts) { out.set(p, pos); pos += p.length; }
    return out;
  }

  // ── XML 工具 ────────────────────────────────────────────────────────
  const TE = new TextEncoder();
  function u8(str) { return TE.encode(str); }
  function xe(s) {
    return String(s == null ? '—' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  // ── 固定 XML 片段 ────────────────────────────────────────────────────
  const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

  const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

  const DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

  const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:sz w:val="20"/><w:szCs w:val="20"/>
      <w:lang w:val="zh-TW" w:eastAsia="zh-TW"/>
    </w:rPr></w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:pPr><w:spacing w:after="80"/></w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:jc w:val="center"/><w:spacing w:before="200" w:after="160"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="52"/><w:szCs w:val="52"/><w:color w:val="1C1410"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="0"/><w:spacing w:before="320" w:after="140"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/><w:color w:val="3A1A6E"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="1"/><w:spacing w:before="200" w:after="100"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/><w:color w:val="5A2D82"/></w:rPr>
  </w:style>
</w:styles>`;

  // 頁面可用寬度（A4, 邊距各 1080 dxa）
  const PAGE_W = 9746;
  const BORDER = '<w:top w:val="single" w:sz="4" w:space="0" w:color="AAAAAA"/>'
    + '<w:left w:val="single" w:sz="4" w:space="0" w:color="AAAAAA"/>'
    + '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="AAAAAA"/>'
    + '<w:right w:val="single" w:sz="4" w:space="0" w:color="AAAAAA"/>'
    + '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="AAAAAA"/>'
    + '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="AAAAAA"/>';
  const TBL_PR = `<w:tblPr><w:tblW w:w="${PAGE_W}" w:type="dxa"/><w:tblBorders>${BORDER}</w:tblBorders></w:tblPr>`;

  function pXml(text, style, bold, center) {
    const pParts = [];
    if (style)  pParts.push(`<w:pStyle w:val="${style}"/>`);
    if (center) pParts.push('<w:jc w:val="center"/>');
    const pPr = pParts.length ? `<w:pPr>${pParts.join('')}</w:pPr>` : '';
    const rPr = bold ? '<w:rPr><w:b/></w:rPr>' : '';
    return `<w:p>${pPr}<w:r>${rPr}<w:t xml:space="preserve">${xe(text)}</w:t></w:r></w:p>`;
  }

  function tcXml(text, bold, shade, wDxa) {
    const wp = wDxa ? `<w:tcW w:w="${wDxa}" w:type="dxa"/>` : '';
    const sh = shade ? '<w:shd w:val="clear" w:color="auto" w:fill="EDE8D0"/>' : '';
    const mar = '<w:tcMar><w:top w:w="55" w:type="dxa"/><w:bottom w:w="55" w:type="dxa"/>'
      + '<w:left w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tcMar>';
    const rPr = bold ? '<w:rPr><w:b/></w:rPr>' : '';
    return `<w:tc><w:tcPr>${wp}${sh}${mar}</w:tcPr>`
      + `<w:p><w:r>${rPr}<w:t xml:space="preserve">${xe(text)}</w:t></w:r></w:p></w:tc>`;
  }

  function kvTblXml(rows) {
    const lw = Math.round(PAGE_W * 0.36), vw = PAGE_W - lw;
    return `<w:tbl>${TBL_PR}`
      + rows.map(([l, v]) => `<w:tr>${tcXml(l, true, true, lw)}${tcXml(v, false, false, vw)}</w:tr>`).join('')
      + '</w:tbl>';
  }

  function multiTblXml(headers, dataRows) {
    const n = headers.length;
    const cw = Math.floor(PAGE_W / n);
    const cws = headers.map((_, i) => i === n - 1 ? PAGE_W - cw * (n - 1) : cw);
    const hRow = `<w:tr>${headers.map((h, i) => tcXml(h, true, true, cws[i])).join('')}</w:tr>`;
    const bRows = dataRows.map(row =>
      `<w:tr>${row.map((c, i) => tcXml(c, false, false, cws[i])).join('')}</w:tr>`
    ).join('');
    return `<w:tbl>${TBL_PR}${hRow}${bRows}</w:tbl>`;
  }

  // ── DocxBuilder ─────────────────────────────────────────────────────
  function DocxBuilder() { this._p = []; }

  DocxBuilder.prototype.addTitle      = function (t)      { this._p.push(pXml(t, 'Title', true, true)); };
  DocxBuilder.prototype.addPara       = function (t, b, c){ this._p.push(pXml(t, null, !!b, !!c)); };
  DocxBuilder.prototype.addH1         = function (t)      { this._p.push(pXml(t, 'Heading1')); };
  DocxBuilder.prototype.addH2         = function (t)      { this._p.push(pXml(t, 'Heading2')); };
  DocxBuilder.prototype.addSpacer     = function ()       { this._p.push('<w:p/>'); };
  DocxBuilder.prototype.addKvTable    = function (rows)   { this._p.push(kvTblXml(rows)); };
  DocxBuilder.prototype.addMultiTable = function (h, rows){ this._p.push(multiTblXml(h, rows)); };

  DocxBuilder.prototype.toBlob = function () {
    const WNS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
      + `<w:document ${WNS}><w:body>`
      + this._p.join('')
      + `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>`
      + `<w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080"/></w:sectPr>`
      + `</w:body></w:document>`;

    const zipBytes = makeZip([
      { name: '[Content_Types].xml',          data: u8(CONTENT_TYPES) },
      { name: '_rels/.rels',                  data: u8(ROOT_RELS) },
      { name: 'word/document.xml',            data: u8(docXml) },
      { name: 'word/_rels/document.xml.rels', data: u8(DOC_RELS) },
      { name: 'word/styles.xml',              data: u8(STYLES) },
    ]);

    return new Blob([zipBytes],
      { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  };

  // ── 下載輔助 ────────────────────────────────────────────────────────
  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  global.DocxLite = { DocxBuilder, downloadBlob };

})(typeof window !== 'undefined' ? window : this);
