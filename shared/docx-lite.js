/**
 * docx-lite.js — 輕量 DOCX 產生器（純 JavaScript，無外部依賴）
 * 使用 ZIP STORE 格式打包 Office Open XML (DOCX)
 * v2.0 — 增強版：支援執行摘要頁、風面卡片、對照表等進階排版
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

  // ── 色彩定義 ────────────────────────────────────────────────────────
  // 主色系（藍綠、深灰、白）
  const COLOR = {
    ink:       '1F2B3A',   // 深藍灰（標題底色）
    inkLight:  '2E4057',   // 稍淺的深藍
    paper:     'FAFAFA',   // 白底
    cardBg:    'FFFFFF',   // 卡片背景
    cardAlt:   'F0F4F8',   // 卡片交替列
    accent:    '1A6B8A',   // 礦青色（主調）
    accentPos: '1A5276',   // 正壓藍色
    accentNeg: 'C0392B',   // 負壓橘紅
    accentMod: '922B21',   // 修正值深紅
    accentAmb: 'D4A017',   // 琥珀色（控制值）
    success:   '1A7A4A',   // 成功綠
    gold:      'C9A84C',   // 舊金（標籤）
    muted:     '5D6D7E',   // 輔助文字灰
    border:    'BDC3C7',   // 邊框灰
    headerTxt: 'FFFFFF',   // 標題列文字
    subTxt:    'D6DBDF',   // 副標文字
    rowShade:  'EAF2FF',   // 每10m底色
  };

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

  // ── 增強版 STYLES ────────────────────────────────────────────────────
  const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft JhengHei" w:cs="Arial"/>
      <w:sz w:val="20"/><w:szCs w:val="20"/>
      <w:lang w:val="zh-TW" w:eastAsia="zh-TW"/>
    </w:rPr></w:rPrDefault>
  </w:docDefaults>

  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:pPr><w:spacing w:after="80"/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft JhengHei"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="ReportTitle">
    <w:name w:val="ReportTitle"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:jc w:val="center"/>
      <w:spacing w:before="400" w:after="120"/>
      <w:shd w:val="clear" w:color="auto" w:fill="${COLOR.ink}"/>
      <w:pBdr>
        <w:bottom w:val="single" w:sz="8" w:space="4" w:color="${COLOR.accent}"/>
      </w:pBdr>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft JhengHei"/>
      <w:b/><w:sz w:val="40"/><w:szCs w:val="40"/>
      <w:color w:val="${COLOR.headerTxt}"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:jc w:val="center"/><w:spacing w:before="200" w:after="160"/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft JhengHei"/>
      <w:b/><w:sz w:val="52"/><w:szCs w:val="52"/><w:color w:val="1C1410"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:outlineLvl w:val="0"/>
      <w:spacing w:before="320" w:after="140"/>
      <w:pBdr>
        <w:bottom w:val="single" w:sz="4" w:space="2" w:color="${COLOR.accent}"/>
      </w:pBdr>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft JhengHei"/>
      <w:b/><w:sz w:val="28"/><w:szCs w:val="28"/><w:color w:val="${COLOR.ink}"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:outlineLvl w:val="1"/>
      <w:spacing w:before="200" w:after="100"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft JhengHei"/>
      <w:b/><w:sz w:val="24"/><w:szCs w:val="24"/><w:color w:val="${COLOR.accent}"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="SectionLabel">
    <w:name w:val="SectionLabel"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:spacing w:before="240" w:after="80"/>
      <w:pBdr>
        <w:left w:val="single" w:sz="12" w:space="4" w:color="${COLOR.accent}"/>
      </w:pBdr>
      <w:ind w:left="160"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft JhengHei"/>
      <w:b/><w:sz w:val="22"/><w:szCs w:val="22"/><w:color w:val="${COLOR.inkLight}"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="CardLabel">
    <w:name w:val="CardLabel"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:spacing w:before="80" w:after="40"/>
      <w:shd w:val="clear" w:color="auto" w:fill="${COLOR.accent}"/>
      <w:ind w:left="120" w:right="120"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft JhengHei"/>
      <w:b/><w:sz w:val="20"/><w:szCs w:val="20"/>
      <w:color w:val="${COLOR.headerTxt}"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="AppendixNote">
    <w:name w:val="AppendixNote"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="60" w:after="60"/><w:ind w:left="240"/></w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft JhengHei"/>
      <w:sz w:val="18"/><w:szCs w:val="18"/>
      <w:color w:val="${COLOR.muted}"/>
    </w:rPr>
  </w:style>
</w:styles>`;

  // ── 頁面可用寬度（A4, 邊距各 1080 dxa）──────────────────────────
  const PAGE_W = 9746;

  const BORDER_DEF = '<w:top w:val="single" w:sz="4" w:space="0" w:color="' + COLOR.border + '"/>'
    + '<w:left w:val="single" w:sz="4" w:space="0" w:color="' + COLOR.border + '"/>'
    + '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="' + COLOR.border + '"/>'
    + '<w:right w:val="single" w:sz="4" w:space="0" w:color="' + COLOR.border + '"/>'
    + '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="' + COLOR.border + '"/>'
    + '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="' + COLOR.border + '"/>';

  const TBL_PR = `<w:tblPr><w:tblW w:w="${PAGE_W}" w:type="dxa"/><w:tblBorders>${BORDER_DEF}</w:tblBorders></w:tblPr>`;

  // ── 基礎 XML 建構函式 ───────────────────────────────────────────────

  function pXml(text, style, bold, center, color) {
    const pParts = [];
    if (style)  pParts.push(`<w:pStyle w:val="${style}"/>`);
    if (center) pParts.push('<w:jc w:val="center"/>');
    const pPr = pParts.length ? `<w:pPr>${pParts.join('')}</w:pPr>` : '';
    const rProps = [];
    if (bold)  rProps.push('<w:b/>');
    if (color) rProps.push(`<w:color w:val="${color}"/>`);
    const rPr = rProps.length ? `<w:rPr>${rProps.join('')}</w:rPr>` : '';
    return `<w:p>${pPr}<w:r>${rPr}<w:t xml:space="preserve">${xe(text)}</w:t></w:r></w:p>`;
  }

  function tcXml(text, bold, shade, wDxa, opts) {
    opts = opts || {};
    const wp = wDxa ? `<w:tcW w:w="${wDxa}" w:type="dxa"/>` : '';
    const fillColor = opts.fill || (shade ? 'EDE8D0' : null);
    const sh = fillColor ? `<w:shd w:val="clear" w:color="auto" w:fill="${fillColor}"/>` : '';
    const vAlign = opts.vAlign ? `<w:vAlign w:val="${opts.vAlign}"/>` : '';
    const topBdr = opts.topBorder
      ? `<w:tcBdr><w:top w:val="single" w:sz="${opts.topBorder.sz || 4}" w:space="0" w:color="${opts.topBorder.color || COLOR.border}"/></w:tcBdr>`
      : '';
    const mar = '<w:tcMar><w:top w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/>'
      + '<w:left w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>';
    const rProps = [];
    if (bold) rProps.push('<w:b/>');
    if (opts.color) rProps.push(`<w:color w:val="${opts.color}"/>`);
    if (opts.size)  rProps.push(`<w:sz w:val="${opts.size}"/>`);
    const rPr = rProps.length ? `<w:rPr>${rProps.join('')}</w:rPr>` : '';
    const jc  = opts.center ? '<w:jc w:val="center"/>' : (opts.right ? '<w:jc w:val="right"/>' : '');
    return `<w:tc><w:tcPr>${wp}${sh}${topBdr}${vAlign}${mar}</w:tcPr>`
      + `<w:p><w:pPr>${jc}</w:pPr><w:r>${rPr}<w:t xml:space="preserve">${xe(text)}</w:t></w:r></w:p></w:tc>`;
  }

  function kvTblXml(rows) {
    const lw = Math.round(PAGE_W * 0.36), vw = PAGE_W - lw;
    return `<w:tbl>${TBL_PR}`
      + rows.map(([l, v]) => `<w:tr>${tcXml(l, true, true, lw)}${tcXml(v, false, false, vw)}</w:tr>`).join('')
      + '</w:tbl>';
  }

  function multiTblXml(headers, dataRows, footerRows) {
    const n = headers.length;
    const cw = Math.floor(PAGE_W / n);
    const cws = headers.map((_, i) => i === n - 1 ? PAGE_W - cw * (n - 1) : cw);
    const hRow = `<w:tr>${headers.map((h, i) => tcXml(h, true, true, cws[i], { fill: COLOR.ink, color: COLOR.headerTxt })).join('')}</w:tr>`;
    const bRows = dataRows.map((row, ri) => {
      const isShaded = ri % 2 === 1;
      return `<w:tr>${row.map((c, i) => tcXml(c, false, isShaded, cws[i])).join('')}</w:tr>`;
    }).join('');
    const fRows = footerRows
      ? footerRows.map(row => `<w:tr>${row.map((c, i) => tcXml(c, true, true, cws[i], { fill: COLOR.cardAlt })).join('')}</w:tr>`).join('')
      : '';
    return `<w:tbl>${TBL_PR}${hRow}${bRows}${fRows}</w:tbl>`;
  }

  // ── 新增：分頁符號 ───────────────────────────────────────────────
  function pageBreakXml() {
    return `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
  }

  // ── 新增：水平分隔線 ─────────────────────────────────────────────
  function hrXml(color, thickness) {
    color = color || COLOR.border;
    thickness = thickness || 4;
    return `<w:p><w:pPr><w:spacing w:before="80" w:after="80"/>` +
      `<w:pBdr><w:bottom w:val="single" w:sz="${thickness}" w:space="1" w:color="${color}"/></w:pBdr>` +
      `</w:pPr></w:p>`;
  }

  // ── 新增：報告標題頁 header ──────────────────────────────────────
  function reportTitleXml(title, subtitle, meta) {
    // meta = [{label, value}, ...]
    let xml = '';
    // 深色底標題段落
    xml += `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="240" w:after="0"/>` +
      `<w:shd w:val="clear" w:color="auto" w:fill="${COLOR.ink}"/>` +
      `</w:pPr><w:r><w:rPr><w:b/><w:sz w:val="44"/><w:szCs w:val="44"/>` +
      `<w:color w:val="${COLOR.headerTxt}"/>` +
      `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft JhengHei"/>` +
      `</w:rPr><w:t xml:space="preserve">${xe(title)}</w:t></w:r></w:p>`;

    if (subtitle) {
      xml += `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/>` +
        `<w:shd w:val="clear" w:color="auto" w:fill="${COLOR.ink}"/>` +
        `</w:pPr><w:r><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/>` +
        `<w:color w:val="${COLOR.subTxt}"/>` +
        `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft JhengHei"/>` +
        `</w:rPr><w:t xml:space="preserve">${xe(subtitle)}</w:t></w:r></w:p>`;
    }

    // 底部強調線
    xml += `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/>` +
      `<w:shd w:val="clear" w:color="auto" w:fill="${COLOR.accent}"/>` +
      `</w:pPr><w:r><w:rPr><w:sz w:val="6"/></w:rPr><w:t> </w:t></w:r></w:p>`;

    xml += `<w:p><w:pPr><w:spacing w:before="120" w:after="120"/></w:pPr></w:p>`;

    // 基本資訊橫排表（若有 meta）
    if (meta && meta.length > 0) {
      const colW = Math.floor(PAGE_W / meta.length);
      xml += `<w:tbl><w:tblPr><w:tblW w:w="${PAGE_W}" w:type="dxa"/>` +
        `<w:tblBorders>` +
        `<w:top w:val="single" w:sz="4" w:space="0" w:color="${COLOR.border}"/>` +
        `<w:bottom w:val="single" w:sz="4" w:space="0" w:color="${COLOR.border}"/>` +
        `<w:insideV w:val="single" w:sz="4" w:space="0" w:color="${COLOR.border}"/>` +
        `</w:tblBorders></w:tblPr><w:tr>`;
      meta.forEach((m, i) => {
        const w = i === meta.length - 1 ? PAGE_W - colW * (meta.length - 1) : colW;
        xml += `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>` +
          `<w:shd w:val="clear" w:color="auto" w:fill="${COLOR.cardAlt}"/>` +
          `<w:tcMar><w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/>` +
          `<w:left w:w="160" w:type="dxa"/><w:right w:w="160" w:type="dxa"/></w:tcMar>` +
          `</w:tcPr>` +
          `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="20"/></w:pPr>` +
          `<w:r><w:rPr><w:sz w:val="16"/><w:color w:val="${COLOR.muted}"/>` +
          `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft JhengHei"/>` +
          `</w:rPr><w:t xml:space="preserve">${xe(m.label)}</w:t></w:r></w:p>` +
          `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/></w:pPr>` +
          `<w:r><w:rPr><w:b/><w:sz w:val="22"/><w:color w:val="${COLOR.inkLight}"/>` +
          `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft JhengHei"/>` +
          `</w:rPr><w:t xml:space="preserve">${xe(m.value)}</w:t></w:r></w:p>` +
          `</w:tc>`;
      });
      xml += '</w:tr></w:tbl>';
    }

    return xml;
  }

  // ── 新增：執行摘要卡片組 ─────────────────────────────────────────
  /**
   * cards = [{ title, color(hex), items:[{label,value,unit?,note?,valueColor?}] }, ...]
   * 每張卡片獨佔一行，左側有色帶
   */
  function execSummaryCardsXml(cards) {
    let xml = '';
    cards.forEach(card => {
      const accentColor = card.color || COLOR.accent;
      // 卡片標題行（左側色帶用深色背景模擬）
      const titleW = Math.round(PAGE_W * 0.06);
      const bodyW  = PAGE_W - titleW;

      // 用表格模擬左側色帶 + 卡片內容
      xml += `<w:tbl><w:tblPr><w:tblW w:w="${PAGE_W}" w:type="dxa"/>` +
        `<w:tblBorders>` +
        `<w:top w:val="single" w:sz="4" w:space="0" w:color="${COLOR.border}"/>` +
        `<w:left w:val="single" w:sz="4" w:space="0" w:color="${COLOR.border}"/>` +
        `<w:bottom w:val="single" w:sz="4" w:space="0" w:color="${COLOR.border}"/>` +
        `<w:right w:val="single" w:sz="4" w:space="0" w:color="${COLOR.border}"/>` +
        `<w:insideH w:val="none"/>` +
        `<w:insideV w:val="single" w:sz="4" w:space="0" w:color="${COLOR.border}"/>` +
        `</w:tblBorders><w:tblCellSpacing w:w="0" w:type="dxa"/></w:tblPr>`;

      // 標題行
      xml += `<w:tr>`;
      // 色帶格（左側）
      xml += `<w:tc><w:tcPr><w:tcW w:w="${titleW}" w:type="dxa"/>` +
        `<w:shd w:val="clear" w:color="auto" w:fill="${accentColor}"/>` +
        `<w:tcMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>` +
        `<w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tcMar>` +
        `</w:tcPr><w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>` +
        `<w:r><w:t> </w:t></w:r></w:p></w:tc>`;
      // 標題格
      xml += `<w:tc><w:tcPr><w:tcW w:w="${bodyW}" w:type="dxa"/>` +
        `<w:shd w:val="clear" w:color="auto" w:fill="${COLOR.inkLight}"/>` +
        `<w:tcMar><w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/>` +
        `<w:left w:w="160" w:type="dxa"/><w:right w:w="160" w:type="dxa"/></w:tcMar>` +
        `</w:tcPr>` +
        `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>` +
        `<w:r><w:rPr><w:b/><w:sz w:val="22"/><w:color w:val="${COLOR.headerTxt}"/>` +
        `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft JhengHei"/>` +
        `</w:rPr><w:t xml:space="preserve">${xe(card.title)}</w:t></w:r></w:p></w:tc>`;
      xml += '</w:tr>';

      // 資料行
      const items = card.items || [];
      items.forEach((item, ri) => {
        const isAlt = ri % 2 === 1;
        const rowFill = isAlt ? COLOR.cardAlt : COLOR.cardBg;
        const lw = Math.round(bodyW * 0.44);
        const vw = Math.round(bodyW * 0.38);
        const uw = bodyW - lw - vw;
        const vColor = item.valueColor || (item.isPos ? COLOR.accentPos : item.isNeg ? COLOR.accentNeg : COLOR.inkLight);

        xml += `<w:tr>`;
        // 色帶延續
        xml += `<w:tc><w:tcPr><w:tcW w:w="${titleW}" w:type="dxa"/>` +
          `<w:shd w:val="clear" w:color="auto" w:fill="${accentColor}"/>` +
          `<w:tcMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>` +
          `<w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tcMar>` +
          `</w:tcPr><w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>` +
          `<w:r><w:t> </w:t></w:r></w:p></w:tc>`;

        // 標籤欄
        xml += `<w:tc><w:tcPr><w:tcW w:w="${lw}" w:type="dxa"/>` +
          `<w:shd w:val="clear" w:color="auto" w:fill="${rowFill}"/>` +
          `<w:tcMar><w:top w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/>` +
          `<w:left w:w="160" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tcMar>` +
          `</w:tcPr>` +
          `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>` +
          `<w:r><w:rPr><w:sz w:val="20"/><w:color w:val="${COLOR.muted}"/>` +
          `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft JhengHei"/>` +
          `</w:rPr><w:t xml:space="preserve">${xe(item.label)}</w:t></w:r></w:p></w:tc>`;

        // 數值欄
        xml += `<w:tc><w:tcPr><w:tcW w:w="${vw}" w:type="dxa"/>` +
          `<w:shd w:val="clear" w:color="auto" w:fill="${rowFill}"/>` +
          `<w:tcMar><w:top w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/>` +
          `<w:left w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tcMar>` +
          `</w:tcPr>` +
          `<w:p><w:pPr><w:jc w:val="right"/><w:spacing w:before="0" w:after="0"/></w:pPr>` +
          `<w:r><w:rPr><w:b/><w:sz w:val="22"/><w:color w:val="${vColor}"/>` +
          `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft JhengHei"/>` +
          `</w:rPr><w:t xml:space="preserve">${xe(item.value)}</w:t></w:r></w:p></w:tc>`;

        // 單位/備註欄
        const unitText = (item.unit || '') + (item.note ? ' ' + item.note : '');
        xml += `<w:tc><w:tcPr><w:tcW w:w="${uw}" w:type="dxa"/>` +
          `<w:shd w:val="clear" w:color="auto" w:fill="${rowFill}"/>` +
          `<w:tcMar><w:top w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/>` +
          `<w:left w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>` +
          `</w:tcPr>` +
          `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>` +
          `<w:r><w:rPr><w:sz w:val="18"/><w:color w:val="${COLOR.muted}"/>` +
          `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft JhengHei"/>` +
          `</w:rPr><w:t xml:space="preserve">${xe(unitText)}</w:t></w:r></w:p></w:tc>`;

        xml += '</w:tr>';
      });

      xml += '</w:tbl>';
      xml += `<w:p><w:pPr><w:spacing w:before="120" w:after="0"/></w:pPr></w:p>`;
    });
    return xml;
  }

  // ── 新增：風面卡片（控制工況頁）────────────────────────────────
  /**
   * faces = [{ name, cp, pPlus, pMinus, designValue, controlCase, type:'pos'|'neg'|'mod' }]
   */
  function windFacesTableXml(faces) {
    const headers = ['風面', 'Cp', 'p+ (N/m²)', 'p- (N/m²)', '設計值 (N/m²)', '控制工況'];
    const n = headers.length;
    const colWs = [
      Math.round(PAGE_W * 0.13),
      Math.round(PAGE_W * 0.10),
      Math.round(PAGE_W * 0.16),
      Math.round(PAGE_W * 0.16),
      Math.round(PAGE_W * 0.18),
      PAGE_W - Math.round(PAGE_W * 0.13) - Math.round(PAGE_W * 0.10)
        - Math.round(PAGE_W * 0.16) * 2 - Math.round(PAGE_W * 0.18),
    ];

    let xml = `<w:tbl><w:tblPr><w:tblW w:w="${PAGE_W}" w:type="dxa"/>` +
      `<w:tblBorders>${BORDER_DEF}</w:tblBorders></w:tblPr>`;

    // 表頭
    xml += `<w:tr>`;
    headers.forEach((h, i) => {
      xml += `<w:tc><w:tcPr><w:tcW w:w="${colWs[i]}" w:type="dxa"/>` +
        `<w:shd w:val="clear" w:color="auto" w:fill="${COLOR.ink}"/>` +
        `<w:tcMar><w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/>` +
        `<w:left w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>` +
        `</w:tcPr>` +
        `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/></w:pPr>` +
        `<w:r><w:rPr><w:b/><w:sz w:val="18"/><w:color w:val="${COLOR.headerTxt}"/>` +
        `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft JhengHei"/>` +
        `</w:rPr><w:t xml:space="preserve">${xe(h)}</w:t></w:r></w:p></w:tc>`;
    });
    xml += '</w:tr>';

    // 資料行
    faces.forEach((face, ri) => {
      const isAlt = ri % 2 === 1;
      const rowFill = isAlt ? COLOR.cardAlt : COLOR.cardBg;
      // 設計值顏色
      const typeMap = { pos: COLOR.accentPos, neg: COLOR.accentNeg, mod: COLOR.accentMod };
      const dvColor = typeMap[face.type] || COLOR.inkLight;
      // 控制工況顏色
      const ctrlColor = face.type === 'neg' ? COLOR.accentNeg : (face.type === 'mod' ? COLOR.accentMod : COLOR.accentPos);

      const cells = [
        { val: face.name,        bold: true,  color: COLOR.inkLight, center: true },
        { val: face.cp,          bold: false, color: COLOR.muted,    center: true },
        { val: face.pPlus,       bold: false, color: COLOR.accentPos,center: true },
        { val: face.pMinus,      bold: false, color: COLOR.accentNeg,center: true },
        { val: face.designValue, bold: true,  color: dvColor,        center: true },
        { val: face.controlCase, bold: false, color: ctrlColor,      center: false },
      ];

      xml += '<w:tr>';
      cells.forEach((cell, ci) => {
        xml += `<w:tc><w:tcPr><w:tcW w:w="${colWs[ci]}" w:type="dxa"/>` +
          `<w:shd w:val="clear" w:color="auto" w:fill="${rowFill}"/>` +
          `<w:tcMar><w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/>` +
          `<w:left w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>` +
          `</w:tcPr>` +
          `<w:p><w:pPr><w:jc w:val="${cell.center ? 'center' : 'left'}"/>` +
          `<w:spacing w:before="0" w:after="0"/></w:pPr>` +
          `<w:r><w:rPr>${cell.bold ? '<w:b/>' : ''}<w:sz w:val="20"/>` +
          `<w:color w:val="${cell.color}"/>` +
          `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft JhengHei"/>` +
          `</w:rPr><w:t xml:space="preserve">${xe(cell.val)}</w:t></w:r></w:p></w:tc>`;
      });
      xml += '</w:tr>';
    });

    xml += '</w:tbl>';
    return xml;
  }

  // ── 新增：附錄長表（逐高度，每10m加底色分隔）──────────────────
  /**
   * headers: string[]
   * rows: string[][]
   * markerCol: 哪個欄位是「高度」欄（index，預設 0）
   * markerInterval: 每幾行加深色分隔（預設按高度自動判斷 10 的倍數）
   */
  function appendixTableXml(headers, rows, opts) {
    opts = opts || {};
    const n = headers.length;
    const cw = Math.floor(PAGE_W / n);
    const cws = headers.map((_, i) => i === n - 1 ? PAGE_W - cw * (n - 1) : cw);

    // 確認哪些行是 10m 整數高度
    const markerCol = opts.markerCol !== undefined ? opts.markerCol : 0;

    let xml = `<w:tbl><w:tblPr><w:tblW w:w="${PAGE_W}" w:type="dxa"/>` +
      `<w:tblBorders>${BORDER_DEF}</w:tblBorders>` +
      `<w:tblHeader/></w:tblPr>`;

    // 表頭（設定 tblHeader 使跨頁重複）
    xml += `<w:tr><w:trPr><w:tblHeader/></w:trPr>`;
    headers.forEach((h, i) => {
      xml += `<w:tc><w:tcPr><w:tcW w:w="${cws[i]}" w:type="dxa"/>` +
        `<w:shd w:val="clear" w:color="auto" w:fill="${COLOR.ink}"/>` +
        `<w:tcMar><w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/>` +
        `<w:left w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>` +
        `</w:tcPr>` +
        `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/></w:pPr>` +
        `<w:r><w:rPr><w:b/><w:sz w:val="18"/><w:color w:val="${COLOR.headerTxt}"/>` +
        `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft JhengHei"/>` +
        `</w:rPr><w:t xml:space="preserve">${xe(h)}</w:t></w:r></w:p></w:tc>`;
    });
    xml += '</w:tr>';

    // 資料行
    rows.forEach((row, ri) => {
      // 判斷是否為 10m 整數高度（若第一欄為數值且為10的倍數）
      const heightVal = parseFloat(row[markerCol]);
      const is10m = !isNaN(heightVal) && Number.isInteger(heightVal) && heightVal % 10 === 0;
      const isAlt  = ri % 2 === 1;
      let rowFill  = isAlt ? COLOR.cardAlt : COLOR.cardBg;
      if (is10m) rowFill = COLOR.rowShade;

      xml += '<w:tr>';
      row.forEach((cell, ci) => {
        // 高度欄加粗
        const isBold = ci === markerCol && is10m;
        xml += `<w:tc><w:tcPr><w:tcW w:w="${cws[ci]}" w:type="dxa"/>` +
          `<w:shd w:val="clear" w:color="auto" w:fill="${rowFill}"/>` +
          `<w:tcMar><w:top w:w="60" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/>` +
          `<w:left w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>` +
          `</w:tcPr>` +
          `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/></w:pPr>` +
          `<w:r><w:rPr>${isBold ? '<w:b/>' : ''}<w:sz w:val="18"/>` +
          `<w:color w:val="${is10m ? COLOR.accent : COLOR.inkLight}"/>` +
          `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft JhengHei"/>` +
          `</w:rPr><w:t xml:space="preserve">${xe(cell)}</w:t></w:r></w:p></w:tc>`;
      });
      xml += '</w:tr>';
    });

    xml += '</w:tbl>';
    return xml;
  }

  // ── 新增：修正前後對照表 ─────────────────────────────────────────
  /**
   * headers: string[]
   * rows: {original, modified, isModified?}[] — 每一行含原始值、修正值、是否有變動
   * colDefs: [{key, label}] — 欄位定義
   */
  function comparisonTableXml(headers, rows, opts) {
    opts = opts || {};
    const n = headers.length;
    const cw = Math.floor(PAGE_W / n);
    const cws = headers.map((_, i) => i === n - 1 ? PAGE_W - cw * (n - 1) : cw);

    let xml = `<w:tbl><w:tblPr><w:tblW w:w="${PAGE_W}" w:type="dxa"/>` +
      `<w:tblBorders>${BORDER_DEF}</w:tblBorders></w:tblPr>`;

    // 表頭
    xml += '<w:tr>';
    headers.forEach((h, i) => {
      xml += `<w:tc><w:tcPr><w:tcW w:w="${cws[i]}" w:type="dxa"/>` +
        `<w:shd w:val="clear" w:color="auto" w:fill="${COLOR.ink}"/>` +
        `<w:tcMar><w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/>` +
        `<w:left w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>` +
        `</w:tcPr>` +
        `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/></w:pPr>` +
        `<w:r><w:rPr><w:b/><w:sz w:val="18"/><w:color w:val="${COLOR.headerTxt}"/>` +
        `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft JhengHei"/>` +
        `</w:rPr><w:t xml:space="preserve">${xe(h)}</w:t></w:r></w:p></w:tc>`;
    });
    xml += '</w:tr>';

    // 資料行
    rows.forEach((row, ri) => {
      // row 是 string[]，最後一欄若有修正則用紅色
      const isAlt = ri % 2 === 1;
      const hasFlag = opts.modifiedCol !== undefined && row[opts.modifiedCol] === true;
      const rowFill = isAlt ? COLOR.cardAlt : COLOR.cardBg;

      xml += '<w:tr>';
      row.forEach((cell, ci) => {
        // 若為修正欄且有修正，用修正色
        const isModCell = opts.modifiedCol !== undefined && ci === opts.modifiedCol - 1;
        const cellColor = (hasFlag && isModCell) ? COLOR.accentNeg : COLOR.inkLight;
        // 跳過 flag 欄（若 row 末尾有 boolean flag）
        if (cell === true || cell === false) return;

        xml += `<w:tc><w:tcPr><w:tcW w:w="${cws[ci]}" w:type="dxa"/>` +
          `<w:shd w:val="clear" w:color="auto" w:fill="${rowFill}"/>` +
          `<w:tcMar><w:top w:w="60" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/>` +
          `<w:left w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>` +
          `</w:tcPr>` +
          `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/></w:pPr>` +
          `<w:r><w:rPr><w:sz w:val="18"/>` +
          `<w:color w:val="${cellColor}"/>` +
          `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft JhengHei"/>` +
          `</w:rPr><w:t xml:space="preserve">${xe(cell)}</w:t></w:r></w:p></w:tc>`;
      });
      xml += '</w:tr>';
    });

    xml += '</w:tbl>';
    return xml;
  }

  // ── 新增：圖例色條說明 ───────────────────────────────────────────
  function legendXml(items) {
    // items = [{color, label}, ...]
    // 用小型表格排一行圖例
    const n = items.length;
    const cw = Math.floor(PAGE_W / n);
    const cws = items.map((_, i) => i === n - 1 ? PAGE_W - cw * (n - 1) : cw);

    let xml = `<w:tbl><w:tblPr><w:tblW w:w="${PAGE_W}" w:type="dxa"/>` +
      `<w:tblBorders><w:top w:val="none"/><w:bottom w:val="none"/>` +
      `<w:left w:val="none"/><w:right w:val="none"/>` +
      `<w:insideH w:val="none"/><w:insideV w:val="none"/>` +
      `</w:tblBorders></w:tblPr><w:tr>`;

    items.forEach((item, i) => {
      xml += `<w:tc><w:tcPr><w:tcW w:w="${cws[i]}" w:type="dxa"/>` +
        `<w:shd w:val="clear" w:color="auto" w:fill="${item.color}"/>` +
        `<w:tcMar><w:top w:w="60" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/>` +
        `<w:left w:w="160" w:type="dxa"/><w:right w:w="160" w:type="dxa"/></w:tcMar>` +
        `</w:tcPr>` +
        `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/></w:pPr>` +
        `<w:r><w:rPr><w:b/><w:sz w:val="18"/><w:color w:val="${COLOR.headerTxt}"/>` +
        `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Microsoft JhengHei"/>` +
        `</w:rPr><w:t xml:space="preserve">${xe(item.label)}</w:t></w:r></w:p></w:tc>`;
    });
    xml += '</w:tr></w:tbl>';
    return xml;
  }

  // ── DocxBuilder ─────────────────────────────────────────────────────
  function DocxBuilder() { this._p = []; }

  // ── 原有 API（完全保留，不更動）──────────────────────────────────
  DocxBuilder.prototype.addTitle      = function (t)      { this._p.push(pXml(t, 'Title', true, true)); };
  DocxBuilder.prototype.addPara       = function (t, b, c){ this._p.push(pXml(t, null, !!b, !!c)); };
  DocxBuilder.prototype.addH1         = function (t)      { this._p.push(pXml(t, 'Heading1')); };
  DocxBuilder.prototype.addH2         = function (t)      { this._p.push(pXml(t, 'Heading2')); };
  DocxBuilder.prototype.addSpacer     = function ()       { this._p.push('<w:p/>'); };
  DocxBuilder.prototype.addKvTable    = function (rows)   { this._p.push(kvTblXml(rows)); };
  DocxBuilder.prototype.addMultiTable = function (h, rows, footerRows){ this._p.push(multiTblXml(h, rows, footerRows)); };

  // ── 新增 API ─────────────────────────────────────────────────────
  /**
   * 插入分頁符號
   */
  DocxBuilder.prototype.addPageBreak = function () {
    this._p.push(pageBreakXml());
  };

  /**
   * 插入水平分隔線
   * @param {string} [color] - 十六進位色碼（不含#）
   * @param {number} [thickness] - 線寬（pt * 8，預設 4）
   */
  DocxBuilder.prototype.addHr = function (color, thickness) {
    this._p.push(hrXml(color, thickness));
  };

  /**
   * 插入報告標題頁（深色 header + 基本資訊橫排）
   * @param {string} title - 主標題
   * @param {string} [subtitle] - 副標題
   * @param {Array<{label:string, value:string}>} [meta] - 橫排基本資訊（例：地區、地況等）
   */
  DocxBuilder.prototype.addReportTitle = function (title, subtitle, meta) {
    this._p.push(reportTitleXml(title, subtitle, meta));
  };

  /**
   * 插入執行摘要卡片組（先看結論模式）
   * @param {Array<{
   *   title: string,
   *   color: string,
   *   items: Array<{
   *     label: string,
   *     value: string|number,
   *     unit?: string,
   *     note?: string,
   *     valueColor?: string,
   *     isPos?: boolean,
   *     isNeg?: boolean
   *   }>
   * }>} cards
   *
   * 色彩語義：
   *   item.isPos = true  → 藍色（正壓）
   *   item.isNeg = true  → 橘紅色（負壓）
   *   item.valueColor    → 自訂色（優先）
   */
  DocxBuilder.prototype.addExecSummaryCards = function (cards) {
    this._p.push(execSummaryCardsXml(cards));
  };

  /**
   * 插入風面控制工況表
   * @param {Array<{
   *   name: string,
   *   cp: string|number,
   *   pPlus: string|number,
   *   pMinus: string|number,
   *   designValue: string|number,
   *   controlCase: string,
   *   type: 'pos'|'neg'|'mod'
   * }>} faces
   */
  DocxBuilder.prototype.addWindFacesTable = function (faces) {
    this._p.push(windFacesTableXml(faces));
  };

  /**
   * 插入附錄長表（逐高度，每10m加底色）
   * @param {string[]} headers
   * @param {string[][]} rows
   * @param {object} [opts]
   * @param {number} [opts.markerCol=0] - 高度欄 index
   */
  DocxBuilder.prototype.addAppendixTable = function (headers, rows, opts) {
    this._p.push(appendixTableXml(headers, rows, opts));
  };

  /**
   * 插入修正前後對照表
   * @param {string[]} headers
   * @param {Array} rows - 每行為 string[]，末尾可加 boolean 標記是否有修正
   * @param {object} [opts]
   * @param {number} [opts.modifiedCol] - 修正欄 index（該欄用紅色顯示）
   */
  DocxBuilder.prototype.addComparisonTable = function (headers, rows, opts) {
    this._p.push(comparisonTableXml(headers, rows, opts));
  };

  /**
   * 插入圖例色條
   * @param {Array<{color:string, label:string}>} items
   */
  DocxBuilder.prototype.addLegend = function (items) {
    this._p.push(legendXml(items));
  };

  /**
   * 插入節標題（左側色帶樣式）
   * @param {string} text
   */
  DocxBuilder.prototype.addSectionLabel = function (text) {
    this._p.push(pXml(text, 'SectionLabel', true, false));
  };

  /**
   * 插入附錄備註段落（小字灰色）
   * @param {string} text
   */
  DocxBuilder.prototype.addNote = function (text) {
    this._p.push(pXml(text, 'AppendixNote', false, false));
  };

  /**
   * 輸出顏色常數供外部使用（例如設定 card.color）
   */
  DocxBuilder.COLORS = COLOR;

  // ── toBlob ───────────────────────────────────────────────────────
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
