/**
 * 零依赖 DOCX 生成器（柏宝绘 · 图文导出）
 *
 * 只做一件事：把「标题 / 楼层名 / 正文段落 / 图片」写成一份 Word 能打开的 .docx。
 *
 * 两个刻意的取舍：
 * 1. **不引第三方库**。插件是直接放进 SillyTavern 的第三方扩展目录、由浏览器 `import()`
 *    加载的，任何 npm 依赖都得先打包；而 .docx 本质就是一个 ZIP + 几份 OOXML 文本，
 *    自己写一遍比引一个几十 KB 的库更可控（也方便用户直接改源码）。
 * 2. **ZIP 用 store（不压缩）**。图片本身（PNG/JPEG）已经压过一遍，再 deflate 几乎不省；
 *    正文 XML 占比很小。省掉 DEFLATE 实现，也就省掉浏览器 CompressionStream 的兼容分支。
 *
 * 这个文件是纯函数模块：不碰 window / document / fetch，方便用 node 直接跑单测。
 */

/* ------------------------------------------------------------------ *
 * 1. CRC32 + ZIP（store）
 * ------------------------------------------------------------------ */

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let value = i;
    for (let bit = 0; bit < 8; bit++) {
      value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    }
    table[i] = value >>> 0;
  }
  return table;
})();

/** 标准 CRC-32（ZIP 每个条目都要算）。 */
export function crc32(bytes) {
  let crc = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) {
    crc = CRC_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function dosTimestamp(date) {
  const year = Math.max(1980, date.getFullYear());
  const time =
    ((date.getHours() & 0x1f) << 11) |
    ((date.getMinutes() & 0x3f) << 5) |
    (Math.floor(date.getSeconds() / 2) & 0x1f);
  const day =
    (((year - 1980) & 0x7f) << 9) |
    (((date.getMonth() + 1) & 0x0f) << 5) |
    (date.getDate() & 0x1f);
  return { time, day };
}

const encoder = new TextEncoder();

/**
 * 把若干「文件」打成 store 模式的 ZIP。
 * @param {Array<{name: string, data: Uint8Array}>} entries
 * @param {Date} [date]
 * @returns {Uint8Array}
 */
export function zipStore(entries, date = new Date()) {
  const { time, day } = dosTimestamp(date);
  const locals = [];
  const centrals = [];
  let offset = 0;

  for (const entry of entries) {
    const name = encoder.encode(entry.name);
    const data = entry.data instanceof Uint8Array ? entry.data : new Uint8Array(entry.data);
    const checksum = crc32(data);

    const local = new Uint8Array(30 + name.length);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true); // local file header
    lv.setUint16(4, 20, true); // version needed
    lv.setUint16(6, 0, true); // flags
    lv.setUint16(8, 0, true); // method: store
    lv.setUint16(10, time, true);
    lv.setUint16(12, day, true);
    lv.setUint32(14, checksum, true);
    lv.setUint32(18, data.length, true);
    lv.setUint32(22, data.length, true);
    lv.setUint16(26, name.length, true);
    lv.setUint16(28, 0, true); // extra length
    local.set(name, 30);
    locals.push(local, data);

    const central = new Uint8Array(46 + name.length);
    const cv = new DataView(central.buffer);
    cv.setUint32(0, 0x02014b50, true); // central directory header
    cv.setUint16(4, 20, true); // version made by
    cv.setUint16(6, 20, true); // version needed
    cv.setUint16(8, 0, true);
    cv.setUint16(10, 0, true);
    cv.setUint16(12, time, true);
    cv.setUint16(14, day, true);
    cv.setUint32(16, checksum, true);
    cv.setUint32(20, data.length, true);
    cv.setUint32(24, data.length, true);
    cv.setUint16(28, name.length, true);
    cv.setUint16(30, 0, true); // extra
    cv.setUint16(32, 0, true); // comment
    cv.setUint16(34, 0, true); // disk number
    cv.setUint16(36, 0, true); // internal attrs
    cv.setUint32(38, 0, true); // external attrs
    cv.setUint32(42, offset, true);
    central.set(name, 46);
    centrals.push(central);

    offset += local.length + data.length;
  }

  const centralSize = centrals.reduce((sum, part) => sum + part.length, 0);
  const end = new Uint8Array(22);
  const ev = new DataView(end.buffer);
  ev.setUint32(0, 0x06054b50, true); // end of central directory
  ev.setUint16(4, 0, true);
  ev.setUint16(6, 0, true);
  ev.setUint16(8, entries.length, true);
  ev.setUint16(10, entries.length, true);
  ev.setUint32(12, centralSize, true);
  ev.setUint32(16, offset, true);
  ev.setUint16(20, 0, true);

  const output = new Uint8Array(offset + centralSize + end.length);
  let cursor = 0;
  for (const part of [...locals, ...centrals, end]) {
    output.set(part, cursor);
    cursor += part.length;
  }
  return output;
}

/* ------------------------------------------------------------------ *
 * 2. XML 转义 / 正文清洗 / 行内 markdown
 * ------------------------------------------------------------------ */

const XML_CONTROL_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g;

export function escapeXml(value) {
  return String(value ?? '')
    .replace(XML_CONTROL_CHARS, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * ST 的 message.mes 是「markdown + 少量 HTML」的混合体。
 * 这里只做无痛清洗：保留文字、换行；剥掉宿主/第三方塞进来的 HTML 壳。
 * 未识别的尖括号内容原样保留（XML 转义后不会破坏文档）。
 */
const HTML_TAGS =
  /<\/?(?:br|p|div|span|font|strong|b|em|i|u|s|del|ins|strike|small|big|center|table|thead|tbody|tr|td|th|ul|ol|li|hr|h[1-6]|img|a|blockquote|code|pre|sup|sub|details|summary|style|script|iframe|audio|video|source|section|article|header|footer|nav|aside|figure|figcaption|label|button|input|textarea|select|option|time|mark|marquee)\b[^>]*>/gi;

export function sanitizeMessageText(raw) {
  let text = String(raw ?? '');
  // 先把 <br> 和块级标签变成换行，再剥掉其余标签
  text = text.replace(/<br\s*\/?>/gi, '\n');
  text = text.replace(/<\/?(?:div|p|li|tr|h[1-6])\b[^>]*>/gi, '\n');
  text = text.replace(HTML_TAGS, '');
  // 残留的生图标签（完整 tag 已在调用方切走，这里只清半截残骸）
  text = text.replace(/<\/?bbi_image\b[^>]*>/gi, '');
  // HTML 实体（必须在剥标签之后解码，避免解码出新的“标签”）
  text = text
    .replace(/&nbsp;/gi, ' ')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#0?39;/gi, "'")
    .replace(/&apos;/gi, "'")
    .replace(/&amp;/gi, '&');
  return text.replace(/\r\n?/g, '\n').replace(/[ \t]+$/gm, '').trim();
}

/** markdown 链接 → 可读文本（保留地址，避免丢信息）。 */
function flattenLinks(text) {
  return text.replace(/\[([^\]]+)\]\((https?:\/\/[^)\s]+)\)/g, '$1（$2）');
}

/**
 * 行内 markdown → run 数组。支持 **粗** / __粗__ / *斜* / _斜_ / ***粗斜*** /
 * ~~删除线~~ / `代码`。不做嵌套（够用且不会误吃整段）。
 */
const INLINE_TOKEN =
  /(\*\*\*[\s\S]+?\*\*\*|\*\*[\s\S]+?\*\*|\*[\s\S]+?\*|__[\s\S]+?__|_[\s\S]+?_|~~[\s\S]+?~~|`[^`\n]+`)/g;

export function splitInline(text) {
  const runs = [];
  const source = String(text ?? '');
  let cursor = 0;

  const pushPlain = value => {
    if (value) runs.push({ text: value });
  };

  for (const match of source.matchAll(INLINE_TOKEN)) {
    const token = match[0];
    const start = match.index ?? 0;
    if (start > cursor) pushPlain(source.slice(cursor, start));
    // 标记长度必须看前缀，不能看总长度（`*x*` 只有 3 个字符，按长度判断会算出空内容）
    const marker = token.startsWith('***')
      ? '***'
      : token.startsWith('**') || token.startsWith('__') || token.startsWith('~~')
        ? token.slice(0, 2)
        : token.slice(0, 1);
    const inner = token.slice(marker.length, token.length - marker.length);
    if (marker === '***') runs.push({ text: inner, bold: true, italic: true });
    else if (marker === '**' || marker === '__') runs.push({ text: inner, bold: true });
    else if (marker === '~~') runs.push({ text: inner, strike: true });
    else if (marker === '`') runs.push({ text: inner, mono: true });
    else runs.push({ text: inner, italic: true });
    cursor = start + token.length;
  }
  if (cursor < source.length) pushPlain(source.slice(cursor));
  return runs.length ? runs : [{ text: source }];
}

/* ------------------------------------------------------------------ *
 * 3. OOXML 片段
 * ------------------------------------------------------------------ */

function runXml(text, style = {}) {
  if (!text) return '';
  const props = [];
  if (style.bold) props.push('<w:b/>');
  if (style.italic) props.push('<w:i/>');
  if (style.strike) props.push('<w:strike/>');
  if (style.mono) {
    props.push('<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:eastAsia="宋体" w:cs="Consolas"/>');
  }
  if (style.color) props.push(`<w:color w:val="${style.color}"/>`);
  if (style.size) props.push(`<w:sz w:val="${style.size}"/><w:szCs w:val="${style.size}"/>`);
  const rPr = props.length ? `<w:rPr>${props.join('')}</w:rPr>` : '';
  return `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
}

function paragraphPrefix(config = {}) {
  const bits = [];
  // CT_PPr 是强序：pBdr → spacing → ind → jc，顺序错了 Word 会直接判文件损坏
  const borders = [];
  if (config.borderTop) {
    borders.push('<w:top w:val="single" w:sz="6" w:space="6" w:color="D9D9D9"/>');
  }
  if (config.borderBottom) {
    borders.push('<w:bottom w:val="single" w:sz="6" w:space="2" w:color="D9D9D9"/>');
  }
  if (borders.length) bits.push(`<w:pBdr>${borders.join('')}</w:pBdr>`);
  const spacing = [];
  if (config.before != null) spacing.push(`w:before="${config.before}"`);
  if (config.after != null) spacing.push(`w:after="${config.after}"`);
  if (config.line != null) spacing.push(`w:line="${config.line}" w:lineRule="auto"`);
  if (spacing.length) bits.push(`<w:spacing ${spacing.join(' ')}/>`);
  if (config.indent) bits.push(`<w:ind w:left="${config.indent}"/>`);
  if (config.align) bits.push(`<w:jc w:val="${config.align}"/>`);
  return bits.length ? `<w:pPr>${bits.join('')}</w:pPr>` : '';
}

function paragraphXml(runs, config) {
  return `<w:p>${paragraphPrefix(config)}${runs}</w:p>`;
}

/** 段内换行必须是 run 的子元素：`<w:p><w:r><w:br/></w:r></w:p>`，不能裸放在段落里。 */
const LINE_BREAK = '<w:r><w:br/></w:r>';

function textToParagraphs(text, baseStyle = {}) {
  const paragraphs = [];
  const normalized = flattenLinks(String(text ?? '').replace(/\r\n?/g, '\n')).trim();
  if (!normalized) return paragraphs;

  const blocks = normalized.split(/\n{2,}/);
  for (const block of blocks) {
    const trimmed = block.trim();
    if (!trimmed) continue;

    // 标题行（# 开头）：既支持独占一段，也支持「标题行 + 紧随其后的正文」
    const firstLine = trimmed.split('\n')[0];
    const heading = firstLine.match(/^(#{1,6})[ \t]+(.*)$/);
    if (heading) {
      const level = heading[1].length;
      const sizes = { 1: 36, 2: 32, 3: 28, 4: 26, 5: 24, 6: 22 };
      const runs = splitInline(heading[2])
        .map(run => runXml(run.text, { ...baseStyle, bold: true, size: sizes[level] }))
        .join('');
      paragraphs.push(paragraphXml(runs, { before: 240, after: 120, line: 320 }));

      const rest = trimmed.slice(firstLine.length).replace(/^\n+/, '');
      if (rest) paragraphs.push(...textToParagraphs(rest, baseStyle));
      continue;
    }

    // 引用块（> 开头）
    if (/^>/.test(trimmed)) {
      const quoted = trimmed
        .split('\n')
        .map(line => line.replace(/^>\s?/, ''))
        .join('\n');
      const runs = quoted
        .split('\n')
        .map(line => splitInline(line).map(run => runXml(run.text, { ...baseStyle, color: '666666' })).join(''))
        .join(LINE_BREAK);
      paragraphs.push(paragraphXml(runs, { indent: 420, before: 60, after: 60 }));
      continue;
    }

    // 无序列表（- / * / • / ·）：逐行处理，续行不加圆点
    const bullet = trimmed.match(/^[-*•·]\s+/);
    if (bullet) {
      const runs = trimmed
        .split('\n')
        .map(line => {
          const isBullet = /^[-*•·]\s+/.test(line);
          const body = isBullet ? line.replace(/^[-*•·]\s+/, '') : line;
          const prefix = isBullet ? runXml('• ', baseStyle) : '';
          return prefix + splitInline(body).map(run => runXml(run.text, baseStyle)).join('');
        })
        .join(LINE_BREAK);
      paragraphs.push(paragraphXml(runs, { indent: 420, before: 0, after: 60 }));
      continue;
    }

    // 普通段落：段内单换行 → 一个只含 <w:br/> 的 run
    const lines = trimmed.split('\n');
    const runs = lines
      .map(line => splitInline(line).map(run => runXml(run.text, baseStyle)).join(''))
      .join(LINE_BREAK);
    paragraphs.push(paragraphXml(runs, { before: 0, after: 140, line: 340 }));
  }
  return paragraphs;
}

/* ------------------------------------------------------------------ *
 * 4. 图片尺寸与 drawing
 * ------------------------------------------------------------------ */

const EMU_PER_CM = 360000;

/** 像素（按 96 DPI）→ EMU，并按最大宽高等比缩放。导出便于单测。 */
export function imageExtent(width, height, maxWidthCm = 15, maxHeightCm = 22) {
  const safeWidth = Number(width) > 0 ? Number(width) : 832;
  const safeHeight = Number(height) > 0 ? Number(height) : 1216;
  const pxToCm = 2.54 / 96;
  let widthCm = safeWidth * pxToCm;
  let heightCm = safeHeight * pxToCm;
  const scale = Math.min(1, maxWidthCm / widthCm, maxHeightCm / heightCm);
  widthCm *= scale;
  heightCm *= scale;
  return {
    cx: Math.max(1, Math.round(widthCm * EMU_PER_CM)),
    cy: Math.max(1, Math.round(heightCm * EMU_PER_CM)),
    widthCm,
    heightCm,
  };
}

function drawingXml({ id, relId, name, extent }) {
  return (
    '<w:r><w:drawing>' +
    '<wp:inline distT="0" distB="0" distL="0" distR="0">' +
    `<wp:extent cx="${extent.cx}" cy="${extent.cy}"/>` +
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
    `<wp:docPr id="${id}" name="图片 ${id}"/>` +
    '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
    '<a:graphic>' +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic>' +
    `<pic:nvPicPr><pic:cNvPr id="${id}" name="${escapeXml(name)}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/>' +
    `<a:ext cx="${extent.cx}" cy="${extent.cy}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
    '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
  );
}

/* ------------------------------------------------------------------ *
 * 5. 包结构与 buildDocx
 * ------------------------------------------------------------------ */

function extensionForMime(mime) {
  if (mime === 'image/jpeg') return 'jpeg';
  if (mime === 'image/gif') return 'gif';
  return 'png';
}

function xmlDocument(body) {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" ' +
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<w:body>${body}<w:sectPr>` +
    '<w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
    '</w:sectPr></w:body></w:document>'
  );
}

const STYLES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:docDefaults><w:rPrDefault><w:rPr>' +
  '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="宋体" w:cs="Calibri"/>' +
  '<w:sz w:val="22"/><w:szCs w:val="22"/>' +
  '<w:lang w:val="en-US" w:eastAsia="zh-CN"/>' +
  '</w:rPr></w:rPrDefault>' +
  '<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="340" w:lineRule="auto"/></w:pPr></w:pPrDefault>' +
  '</w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
  '</w:styles>';

function contentTypesXml(hasPng, hasJpeg, hasGif) {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    (hasPng ? '<Default Extension="png" ContentType="image/png"/>' : '') +
    (hasJpeg ? '<Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="jpg" ContentType="image/jpeg"/>' : '') +
    (hasGif ? '<Default Extension="gif" ContentType="image/gif"/>' : '') +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
    '</Types>'
  );
}

const ROOT_RELS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
  '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
  '</Relationships>';

function documentRelsXml(imageRels) {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    imageRels
      .map(
        rel =>
          `<Relationship Id="${rel.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${rel.target}"/>`,
      )
      .join('') +
    '</Relationships>'
  );
}

function coreXml({ title, createdAt }) {
  const stamp = createdAt.toISOString().replace(/\.\d{3}Z$/, 'Z');
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
    'xmlns:dc="http://purl.org/dc/elements/1.1/" ' +
    'xmlns:dcterms="http://purl.org/dc/terms/" ' +
    'xmlns:dcmitype="http://purl.org/dc/dcmitype/" ' +
    'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    `<dc:title>${escapeXml(title)}</dc:title>` +
    '<dc:creator>柏宝绘 · 图文导出</dc:creator>' +
    '<cp:lastModifiedBy>柏宝绘 · 图文导出</cp:lastModifiedBy>' +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${stamp}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${stamp}</dcterms:modified>` +
    '</cp:coreProperties>'
  );
}

const APP_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' +
  'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
  '<Application>柏宝绘 · 图文导出</Application>' +
  '<DocSecurity>0</DocSecurity>' +
  '<ScaleCrop>false</ScaleCrop>' +
  '</Properties>';

function titleXml(title, subtitle) {
  const paragraphs = [];
  if (title) {
    paragraphs.push(
      paragraphXml(runXml(title, { bold: true, size: 40 }), { align: 'center', before: 0, after: 120, line: 320 }),
    );
  }
  if (subtitle) {
    paragraphs.push(
      paragraphXml(runXml(subtitle, { color: '777777', size: 18 }), { align: 'center', before: 0, after: 240 }),
    );
  }
  return paragraphs.join('');
}

function floorHeaderXml(name, meta) {
  const dateRun = meta ? runXml(`　${meta}`, { color: '999999', size: 18 }) : '';
  const title = runXml(name || '（未命名楼层）', { bold: true, size: 26, color: '404040' });
  return paragraphXml(title + dateRun, {
    before: 320,
    after: 100,
    borderTop: true,
  });
}

function captionXml(caption) {
  const lines = String(caption ?? '')
    .split('\n')
    .map(line => line.trim())
    .filter(Boolean);
  return lines
    .map(line =>
      paragraphXml(runXml(line, { color: '888888', size: 18, italic: true }), {
        align: 'center',
        before: 0,
        after: 40,
      }),
    )
    .join('');
}

/**
 * 生成 .docx。
 *
 * blocks:
 * - { type: 'floor', name, meta? }        楼层标题
 * - { type: 'text', text }                正文（内部再切段）
 * - { type: 'image', data, mime, width, height, caption? }  图片
 * - { type: 'note', text }                灰色小字提示（图片缺失等）
 * - { type: 'divider' }                   分隔线
 */
export function buildDocx({ title, subtitle, blocks, createdAt = new Date(), maxWidthCm = 15, maxHeightCm = 22 }) {
  const bodyParts = [titleXml(title, subtitle)];
  const media = [];
  const imageRels = [];
  let docPrId = 0;

  for (const block of blocks || []) {
    if (!block) continue;
    switch (block.type) {
      case 'floor':
        bodyParts.push(floorHeaderXml(block.name, block.meta));
        break;
      case 'divider':
        bodyParts.push(paragraphXml('', { borderBottom: true, before: 240, after: 240 }));
        break;
      case 'note':
        if (block.text) {
          bodyParts.push(
            paragraphXml(runXml(block.text, { color: '999999', size: 18 }), {
              before: 60,
              after: 60,
              align: 'center',
            }),
          );
        }
        break;
      case 'text': {
        const paragraphs = textToParagraphs(block.text);
        bodyParts.push(...paragraphs);
        break;
      }
      case 'image': {
        if (!block.data || !block.data.length) break;
        const ext = extensionForMime(block.mime);
        const index = media.length + 1;
        const fileName = `image${index}.${ext}`;
        media.push({ name: `word/media/${fileName}`, data: block.data });
        const relId = `rIdImg${index}`;
        imageRels.push({ id: relId, target: `media/${fileName}` });
        docPrId += 1;
        const extent = imageExtent(block.width, block.height, maxWidthCm, maxHeightCm);
        bodyParts.push(
          paragraphXml(drawingXml({ id: docPrId, relId, name: fileName, extent }), {
            align: 'center',
            before: 120,
            after: block.caption ? 20 : 160,
          }),
        );
        if (block.caption) bodyParts.push(captionXml(block.caption));
        break;
      }
      default:
        break;
    }
  }

  const hasPng = media.some(item => item.name.endsWith('.png'));
  const hasJpeg = media.some(item => item.name.endsWith('.jpeg'));
  const hasGif = media.some(item => item.name.endsWith('.gif'));

  const files = [
    { name: '[Content_Types].xml', data: encoder.encode(contentTypesXml(hasPng, hasJpeg, hasGif)) },
    { name: '_rels/.rels', data: encoder.encode(ROOT_RELS_XML) },
    { name: 'docProps/core.xml', data: encoder.encode(coreXml({ title: title || '图文导出', createdAt })) },
    { name: 'docProps/app.xml', data: encoder.encode(APP_XML) },
    { name: 'word/document.xml', data: encoder.encode(xmlDocument(bodyParts.join(''))) },
    { name: 'word/styles.xml', data: encoder.encode(STYLES_XML) },
    { name: 'word/_rels/document.xml.rels', data: encoder.encode(documentRelsXml(imageRels)) },
    ...media,
  ];

  return zipStore(files, createdAt);
}
