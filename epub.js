/**
 * 零依赖 EPUB 3 生成器。
 *
 * 输入与 docx.js 的 buildDocx 相同：{ title, subtitle, blocks, createdAt }。
 * 输出一个 store 模式的 EPUB（ZIP），图片内嵌，不依赖任何第三方库。
 */
import { escapeXml, zipStore } from './docx.js';

const MIME = {
  'image/jpeg': 'jpg',
  'image/png': 'png',
  'image/gif': 'gif',
  'image/webp': 'webp',
  'image/svg+xml': 'svg',
};

function imageExtension(mime) {
  return MIME[String(mime || '').toLowerCase()] || 'jpg';
}

/** 章节号用中文：1 → 一，10 → 十，101 → 一百零一，236 → 二百三十六。 */
function chineseNumber(value) {
  const digits = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九'];
  const units = ['', '十', '百', '千'];
  const n = Math.max(0, Math.floor(Number(value) || 0));
  if (n >= 10000) return String(n);
  let text = '';
  for (let place = 3; place >= 0; place--) {
    const digit = Math.floor(n / 10 ** place) % 10;
    if (digit === 0) {
      if (text && !text.endsWith('零') && n % 10 ** place !== 0) text += '零';
      continue;
    }
    text += digits[digit] + units[place];
  }
  text = text.replace(/零+$/, '');
  return text.replace(/^一十/, '十') || '零';
}

/** 章节名：第N章：【用户层标题】；没有用户层标题时只保留第N章。 */
function chapterHeading(number, userTitle = '') {
  const base = `第${chineseNumber(number)}章`;
  const title = String(userTitle || '').replace(/^【/, '').replace(/】$/, '').trim();
  return title ? `${base}：【${title}】` : base;
}

function makeId() {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return `urn:uuid:${crypto.randomUUID()}`;
  }
  return `urn:uuid:${Date.now().toString(16)}-${Math.random().toString(16).slice(2)}`;
}

function isoDate(value) {
  return new Date(value).toISOString().replace(/\.\d{3}Z$/, 'Z');
}

/** 行内 markdown 的最小实现：粗体、斜体、删除线、行内代码。 */
function inlineMarkup(raw) {
  let text = escapeXml(raw);
  text = text.replace(/`([^`]+)`/g, '<code>$1</code>');
  text = text.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
  text = text.replace(/\*([^*]+)\*/g, '<em>$1</em>');
  text = text.replace(/~~([^~]+)~~/g, '<s>$1</s>');
  return text;
}

/** 把一段正文转成 XHTML 段落；支持标题、列表、引用、分隔线。 */
function renderText(text) {
  const lines = String(text ?? '').replace(/\r\n?/g, '\n').split('\n');
  const out = [];
  let list = null;
  let quote = null;

  const flush = () => {
    if (list) {
      out.push(`<ul>${list.join('')}</ul>`);
      list = null;
    }
    if (quote) {
      out.push(`<blockquote>${quote.join('')}</blockquote>`);
      quote = null;
    }
  };

  for (const rawLine of lines) {
    const line = rawLine.trimEnd();
    const trimmed = line.trim();
    if (!trimmed) {
      flush();
      continue;
    }
    const heading = /^(#{1,6})[ \t]+(.+)$/.exec(trimmed);
    if (heading) {
      flush();
      const level = Math.min(6, heading[1].length);
      out.push(`<h${level}>${inlineMarkup(heading[2])}</h${level}>`);
      continue;
    }
    if (/^(?:-{3,}|\*{3,}|_{3,})$/.test(trimmed)) {
      flush();
      out.push('<hr/>');
      continue;
    }
    if (/^[-*][ \t]+/.test(trimmed)) {
      if (quote) {
        out.push(`<blockquote>${quote.join('')}</blockquote>`);
        quote = null;
      }
      if (!list) list = [];
      list.push(`<li>${inlineMarkup(trimmed.replace(/^[-*][ \t]+/, ''))}</li>`);
      continue;
    }
    if (/^&gt;[ \t]/.test(trimmed) || /^>[ \t]/.test(trimmed)) {
      if (list) {
        out.push(`<ul>${list.join('')}</ul>`);
        list = null;
      }
      if (!quote) quote = [];
      quote.push(`<p>${inlineMarkup(trimmed.replace(/^>[ \t]/, ''))}</p>`);
      continue;
    }
    flush();
    out.push(`<p>${inlineMarkup(trimmed)}</p>`);
  }
  flush();
  return out.join('\n');
}

function renderBlocks(blocks, imageNames) {
  const parts = [];
  for (const block of blocks) {
    if (block.type === 'floor') {
      if (block.hidden) continue;
      const meta = block.meta ? `<div class="floor-meta">${escapeXml(block.meta)}</div>` : '';
      parts.push(`<h2 class="floor">${escapeXml(block.name || '')}</h2>${meta}`);
      continue;
    }
    if (block.type === 'text') {
      parts.push(renderText(block.text));
      continue;
    }
    if (block.type === 'note') {
      parts.push(`<p class="note">${escapeXml(block.text)}</p>`);
      continue;
    }
    if (block.type === 'image') {
      const name = imageNames.get(block);
      if (!name) continue;
      const caption = block.caption ? `<figcaption>${escapeXml(block.caption)}</figcaption>` : '';
      parts.push(`<figure><img src="../images/${name}" alt="配图"/>${caption}</figure>`);
    }
  }
  return parts.join('\n');
}

function chapterXhtml(title, body, lang = 'zh-CN') {
  return `<?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" lang="${lang}" xml:lang="${lang}">
<head>
  <meta charset="utf-8"/>
  <title>${escapeXml(title)}</title>
  <link rel="stylesheet" type="text/css" href="../style.css"/>
</head>
<body>
${body}
</body>
</html>`;
}

export function buildEpub({ title, subtitle = '', blocks = [], createdAt = new Date() }) {
  const now = isoDate(createdAt);
  const bookTitle = String(title || '柏宝绘 · 图文导出');
  const images = [];
  const imageNames = new Map();
  for (const block of blocks) {
    if (block.type !== 'image') continue;
    const ext = imageExtension(block.mime);
    const name = `image-${images.length + 1}.${ext}`;
    images.push({ name, data: block.data, mime: block.mime || 'image/jpeg' });
    imageNames.set(block, name);
  }

  const chapters = [];
  let current = null;
  let floorNumber = 0;
  let preface = [];
  for (const block of blocks) {
    if (block.type === 'floor') {
      floorNumber += 1;
      if (current && current.blocks.length) chapters.push(current);
      const number = floorNumber;
      current = { title: chapterHeading(number, block.userTitle), blocks: [...preface, block] };
      preface = [];
      continue;
    }
    if (!current) {
      preface.push(block);
      continue;
    }
    current.blocks.push(block);
  }
  if (current && current.blocks.length) chapters.push(current);
  else if (preface.length) chapters.push({ title: '正文', blocks: preface });

  const titlePage = chapterXhtml(
    bookTitle,
    `<h1>${escapeXml(bookTitle)}</h1>${subtitle ? `<p class="subtitle">${escapeXml(subtitle)}</p>` : ''}<p class="meta">由 柏宝绘 · 图文导出 生成</p>`,
  );
  const chapterFiles = chapters.map((chapter, index) => ({
    name: `chapter-${index + 1}.xhtml`,
    title: chapter.title,
    xhtml: chapterXhtml(
      chapter.title,
      `<h1 class="chapter-title">${escapeXml(chapter.title)}</h1>${renderBlocks(chapter.blocks, imageNames)}`,
    ),
  }));

  const navItems = [
    '<li><a href="text/title.xhtml">封面</a></li>',
    ...chapterFiles.map(chapter => `<li><a href="text/${chapter.name}">${escapeXml(chapter.title)}</a></li>`),
  ];

  const nav = `<?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" lang="zh-CN" xml:lang="zh-CN">
<head><meta charset="utf-8"/><title>目录</title><link rel="stylesheet" type="text/css" href="style.css"/></head>
<body>
<nav epub:type="toc" id="toc"><h1>目录</h1><ol>${navItems.join('\n')}</ol></nav>
</body>
</html>`;

  const identifier = makeId();
  const tocNcx = `<?xml version="1.0" encoding="UTF-8"?>
<ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1">
<head><meta name="dtb:uid" content="${identifier}"/></head>
<docTitle><text>${escapeXml(bookTitle)}</text></docTitle>
<navMap>
<navPoint id="navPoint-1" playOrder="1"><navLabel><text>封面</text></navLabel><content src="text/title.xhtml"/></navPoint>
${chapterFiles.map((chapter, index) => `<navPoint id="navPoint-${index + 2}" playOrder="${index + 2}"><navLabel><text>${escapeXml(chapter.title)}</text></navLabel><content src="text/${chapter.name}"/></navPoint>`).join('\n')}
</navMap>
</ncx>`;

  const imageManifest = images.map((image, index) => `<item id="img-${index + 1}" href="images/${image.name}" media-type="${escapeXml(image.mime)}"/>`).join('\n');
  const chapterManifest = chapterFiles.map((chapter, index) => `<item id="chapter-${index + 1}" href="text/${chapter.name}" media-type="application/xhtml+xml"/>`).join('\n');
  const chapterSpine = chapterFiles.map((chapter, index) => `<itemref idref="chapter-${index + 1}"/>`).join('\n');

  const opf = `<?xml version="1.0" encoding="utf-8"?>
<package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="book-id" prefix="dcterms: http://purl.org/dc/terms/">
  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
    <dc:identifier id="book-id">${identifier}</dc:identifier>
    <dc:title>${escapeXml(bookTitle)}</dc:title>
    <dc:language>zh-CN</dc:language>
    <dc:creator>柏宝绘 · 图文导出</dc:creator>
    <dc:description>${escapeXml(subtitle)}</dc:description>
    <meta property="dcterms:modified">${now}</meta>
  </metadata>
  <manifest>
    <item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>
    <item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/>
    <item id="css" href="style.css" media-type="text/css"/>
    <item id="titlepage" href="text/title.xhtml" media-type="application/xhtml+xml"/>
    ${chapterManifest}
    ${imageManifest}
  </manifest>
  <spine toc="ncx">
    <itemref idref="titlepage"/>
    ${chapterSpine}
  </spine>
</package>`;

  const container = `<?xml version="1.0" encoding="UTF-8"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <rootfiles><rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/></rootfiles>
</container>`;

  const css = `body { font-family: serif; line-height: 1.75; margin: 1.2em; color: #222; }
h1 { font-size: 1.8em; line-height: 1.3; margin: 0.8em 0 0.4em; }
h1.chapter-title { font-size: 1.6em; text-align: center; margin: 0.5em 0 1em; letter-spacing: 0.08em; }
h2 { font-size: 1.25em; margin: 1.6em 0 0.6em; padding-bottom: 0.25em; border-bottom: 1px solid #ccc; }
h3, h4, h5, h6 { line-height: 1.4; }
p { margin: 0.55em 0; text-indent: 1.2em; }
p.note, p.meta, p.subtitle, .floor-meta { text-indent: 0; color: #777; font-size: 0.9em; }
blockquote { margin: 0.8em 0; padding-left: 0.8em; border-left: 3px solid #999; color: #444; }
ul { margin: 0.6em 0 0.6em 1.4em; }
li { margin: 0.2em 0; }
code { font-family: monospace; background: #f3f3f3; padding: 0 0.2em; }
figure { margin: 1.2em 0; text-align: center; page-break-inside: avoid; }
img { max-width: 100%; height: auto; }
figcaption { margin-top: 0.35em; color: #666; font-size: 0.85em; text-align: left; }
hr { border: none; border-top: 1px solid #bbb; margin: 1.4em 0; }`;

  const entries = [
    { name: 'mimetype', data: new Uint8Array([...new TextEncoder().encode('application/epub+zip')]) },
    { name: 'META-INF/container.xml', data: new TextEncoder().encode(container) },
    { name: 'OEBPS/content.opf', data: new TextEncoder().encode(opf) },
    { name: 'OEBPS/nav.xhtml', data: new TextEncoder().encode(nav) },
    { name: 'OEBPS/toc.ncx', data: new TextEncoder().encode(tocNcx) },
    { name: 'OEBPS/style.css', data: new TextEncoder().encode(css) },
    { name: 'OEBPS/text/title.xhtml', data: new TextEncoder().encode(titlePage) },
    ...chapterFiles.map(chapter => ({ name: `OEBPS/text/${chapter.name}`, data: new TextEncoder().encode(chapter.xhtml) })),
    ...images.map(image => ({ name: `OEBPS/images/${image.name}`, data: image.data })),
  ];
  return zipStore(entries, createdAt);
}
