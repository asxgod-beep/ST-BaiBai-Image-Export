/**
 * 柏宝绘 · 图文导出（SillyTavern 第三方扩展）
 *
 * 做一件事：把**当前聊天里所有 AI 楼层**的正文 + 柏宝绘配图，导出成一份图文混排的 Word(.docx)。
 *
 * 数据来源（只读，不写任何聊天记录）：
 * - 正文：`chat[i].mes`（当前 swipe 的那一份），默认**只取正文段**：
 *   先整块剥掉思维链 / 吐槽 / 状态类标签，再以 `【 时间 】` 行作为正文起点；
 * - 配图：正文里的 `<bbi_image>…</bbi_image>` 标签 + `chat[i].extra.bbiImage`
 *   （`{ [swipeId]: { [promptHash]: Entry[] } }`，与柏宝绘卡片同源）
 *
 * 注意：正文提取必须在解析 `<bbi_image>` **之前**完成。思维链里经常把 tag 举例写出来，
 * 先剥掉再切图，`slotSeq` 才和柏宝绘真正生成的槽位对得上。
 *
 * 几个关键口径必须与柏宝绘逐字一致，否则图片会「在聊天里看得到、导出后没有」：
 * - `promptHash`：柏宝绘侧同款 cyrb53 派生，吃 tag **原文**；
 * - 槽位匹配：有 `<slot>` 用 slotId，没有则按 tag 序号 + `slotSeq`；
 * - 找图顺序：先找当前提示词的成品，再退到同槽位的旧提示词成品（卡片显示的就是这张）。
 *
 * 导出产物完全在浏览器本地生成，不经过 ST 服务端；图片先由浏览器从 ST 静态目录取回，
 * 再以 base64 之外的原生二进制写进 docx（docx 内部本来就是 ZIP，不需要再编码）。
 */

import { buildDocx, sanitizeMessageText } from './docx.js';
import { buildEpub } from './epub.js';

const SETTINGS_KEY = 'baibai_image_export';
const PANEL_ID = 'bbi-export-panel';
const MENU_ITEM_ID = 'bbi-export-menu-item';
const RUN_BUTTON_ID = 'bbi-export-run';
const STATUS_ID = 'bbi-export-status';
const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const EPUB_MIME = 'application/epub+zip';

/** 与柏宝绘默认设置保持同一套「清晰阅读」取向：提示词默认不导出，图片默认压缩。 */
const DEFAULT_SETTINGS = {
  showFloorName: true,
  showTime: false,
  includePrompt: false,
  onlyWithImages: false,
  compressImages: true,
  bodyOnly: true,
  bodyTag: '',
  excludeTags: '',
  format: 'docx',
};

/** 压缩后的最长边（px）。15cm 宽约等于 270DPI，够印刷级阅读。 */
const MAX_IMAGE_EDGE = 1600;
const JPEG_QUALITY = 0.86;
/** 同时读取几张图。ST 是单进程，压太狠会卡住用户自己的聊天。 */
const FETCH_CONCURRENCY = 4;

let busy = false;
let initialized = false;

/* ------------------------------------------------------------------ *
 * 1. 宿主上下文与设置
 * ------------------------------------------------------------------ */

function getContext() {
  try {
    return window.SillyTavern?.getContext?.() ?? null;
  } catch {
    return null;
  }
}

function readSettings(ctx = getContext()) {
  const settings = { ...DEFAULT_SETTINGS };
  const stored = ctx?.extensionSettings?.[SETTINGS_KEY];
  if (stored && typeof stored === 'object') {
    for (const key of Object.keys(DEFAULT_SETTINGS)) {
      const fallback = DEFAULT_SETTINGS[key];
      if (typeof fallback === 'boolean' && typeof stored[key] === 'boolean') settings[key] = stored[key];
      else if (typeof fallback === 'string' && typeof stored[key] === 'string') settings[key] = stored[key];
    }
  }
  return settings;
}

function writeSetting(key, value) {
  const ctx = getContext();
  if (!ctx?.extensionSettings) return;
  const settings = readSettings(ctx);
  settings[key] = value;
  ctx.extensionSettings[SETTINGS_KEY] = settings;
  ctx.saveSettingsDebounced?.();
}

/* ------------------------------------------------------------------ *
 * 2. 与柏宝绘同源的口径：tag、hash、槽位、图片指针
 * ------------------------------------------------------------------ */

/** 柏宝绘 `st/imageTagRegex.ts` 的运行时正则同款（每次新建，避免 /g 的 lastIndex 状态）。 */
function findImageTags(mes) {
  const regex = /<bbi_image>[\s\S]+?<\/bbi_image>/gi;
  const tags = [];
  let match;
  while ((match = regex.exec(mes)) !== null) {
    tags.push({ raw: match[0], index: match.index, end: match.index + match[0].length });
  }
  return tags;
}

/** 柏宝绘 `floor/storage.ts` 的 promptHash：必须逐字一致，否则 extra 里的桶对不上。 */
function promptHash(text) {
  let h1 = 0xdeadbeef;
  let h2 = 0x41c6ce57;
  for (let i = 0; i < text.length; i++) {
    const ch = text.charCodeAt(i);
    h1 = Math.imul(h1 ^ ch, 2654435761);
    h2 = Math.imul(h2 ^ ch, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(16).padStart(14, '0');
}

function parseSlotId(rawTag) {
  const inner = rawTag.replace(/^<bbi_image[^>]*>/i, '').replace(/<\/bbi_image>$/i, '');
  return inner.match(/<slot>([a-zA-Z0-9_-]{1,100})<\/slot>/i)?.[1] || '';
}

function readStore(message, swipeId) {
  const candidates = [
    message?.extra?.bbiImage,
    message?.swipe_info?.[swipeId]?.extra?.bbiImage,
  ];
  for (const candidate of candidates) {
    if (candidate && typeof candidate === 'object') return candidate;
  }
  return null;
}

function matchesSlot(entry, seq, slotId) {
  return slotId ? entry.slotId === slotId : !entry.slotId && (entry.slotSeq ?? 0) === seq;
}

/**
 * 取这个 tag 对应的那张图：
 * 1) 当前提示词的桶里找同槽位最新一条（卡片 ready）；
 * 2) 找不到就跨桶找同槽位最新一条（卡片 stale，用户看到的也是它）。
 * 与柏宝绘卡片一致：旧提示词的图仍然导出，不让「改过提示词」变成「图丢了」。
 */
function pickImageEntry(store, swipeId, rawTag, seq) {
  const bucket = store?.[String(swipeId)];
  if (!bucket || typeof bucket !== 'object') return null;
  const hash = promptHash(rawTag);
  const slotId = parseSlotId(rawTag);
  const usable = entry => entry && entry.status !== 'error' && typeof entry.path === 'string' && entry.path;

  const current = Array.isArray(bucket[hash]) ? bucket[hash] : [];
  for (let i = current.length - 1; i >= 0; i--) {
    if (usable(current[i]) && matchesSlot(current[i], seq, slotId)) return current[i];
  }

  let latest = null;
  for (const [key, entries] of Object.entries(bucket)) {
    if (key === hash || !Array.isArray(entries)) continue;
    for (let i = entries.length - 1; i >= 0; i--) {
      const entry = entries[i];
      if (!usable(entry) || !matchesSlot(entry, seq, slotId)) continue;
      if (!latest || (entry.createdAt ?? 0) > (latest.createdAt ?? 0)) latest = entry;
      break;
    }
  }
  return latest;
}

/** 与柏宝绘 `isAiStoryMessage` 同口径：不是用户楼，且不是带 extra.type 的真系统楼。 */
function isAiStoryMessage(message) {
  if (!message || typeof message.mes !== 'string' || !message.mes.trim()) return false;
  if (message.is_user) return false;
  if (message.is_system && typeof message.extra?.type === 'string') return false;
  return true;
}

/** 图片下方那行小字：把 tag 拆成「正文提示词 / 自然语言 / 角色 / 负面」。 */
function promptText(rawTag) {
  const inner = rawTag.replace(/^<bbi_image[^>]*>/i, '').replace(/<\/bbi_image>$/i, '');
  const oneLine = value => String(value ?? '').trim().replace(/[\r\n]+/g, ' ');
  const nl = oneLine(inner.match(/<nl>([\s\S]*?)<\/nl>/i)?.[1] || '');
  const negative = oneLine(inner.match(/<negative>([\s\S]*?)<\/negative>/i)?.[1] || '');
  const stripped = inner
    .replace(/<(slot|pending|render|nl|negative|characters|size)>[\s\S]*?<\/\1>/gi, '')
    .replace(/<\/?tag>/gi, '');
  const tag = oneLine(stripped);

  const lines = [];
  if (tag) lines.push(tag);
  if (nl) lines.push(nl);
  try {
    const parsed = JSON.parse(inner.match(/<characters>([\s\S]*?)<\/characters>/i)?.[1] || '[]');
    if (Array.isArray(parsed)) {
      for (const character of parsed.slice(0, 32)) {
        if (!character || typeof character !== 'object') continue;
        const name = oneLine(character.name);
        const characterTag = oneLine(character.tag);
        if (!name || !characterTag) continue;
        lines.push(`角色: ${name}`, characterTag);
        if (character.nl) lines.push(oneLine(character.nl));
      }
    }
  } catch {
    // 角色块坏了不影响整段提示词的展示
  }
  if (negative) lines.push(`Negative: ${negative}`);
  return lines.join('\n');
}

/* ------------------------------------------------------------------ *
 * 3. 正文提取：剥掉思维链 / 吐槽 / 状态块，只留正文
 * ------------------------------------------------------------------ */

/**
 * 常见「非正文」标签。用户的角色卡千奇百怪（比如 `<konatan_planning~>`、`<tucao>`、
 * `<current_event>`、`<progress>`），所以这里既有精确名单，也有关键词包含匹配
 * （`planning` / `cot` / `tucao` / `progress` …），任何前缀或后缀都能命中。
 */
const META_TAG_KEYWORDS = [
  'planning', 'plan', 'thinking', 'think', 'thought', 'cot', 'reasoning', 'analysis', 'analyse', 'analyze',
  'tucao', 'comment', 'commentary', 'meta', 'memo', 'memory', 'summary', 'note', 'notes', 'outline',
  'status', 'state', 'progress', 'event', 'instruction', 'rules', 'rule', 'suggestion', 'options', 'choices',
  'scratchpad', 'brainstorm', 'director', 'debug', 'log', 'system',
];

const META_TAG_EXACT = new Set([
  ...META_TAG_KEYWORDS,
  'current_event', 'currentevent', 'current-event', 'chain_of_thought', 'scratch_pad',
  'advice', 'advices', 'tool_call', 'tool_response', 'toolcall', 'toolresponse',
  'dream_scene', 'dream_big_discuss', 'dream', 'q', 'c', 'tuao', 'scene', 'disclaimer',
  'content_target', 'creative_status', 'review_exemption', 'security_bypass', 'ethics_removal',
  'violence_restriction_removal', 'content_will_autonomy', 'date', 'location', 'user_input',
  'history', 'admin', 'character', 'narrative_config', 'writing_style', 'writing_guidance', 'ban',
  'details',
  '思考', '规划', '计划', '吐槽', '备注', '状态', '进度', '事件', '记忆', '总结', '大纲', '设定', '指令', '系统', '草稿',
]);

/** 这些是「思维链容器」：即使只出现闭合标签（模型漏写开标签），也能从开头整段切掉。 */
const PLANNING_CONTAINER_NAMES = new Set([
  'konatan_planning', 'planning', 'plan', 'think', 'thought', 'advice', 'dream_big_discuss', 'dream_scene',
  'scratchpad', 'brainstorm',
]);

/** 这些名字更像「正文容器」，遇到时解开外壳而不是整块丢掉。 */
const BODY_TAG_NAMES = new Set([
  'body', 'content', 'text', 'story', 'narrative', 'main', '正文', '内容', '剧情', '故事', '本文', '小说',
]);

/** 属于排版/展示的 HTML 标签，任何时候都不能当成「非正文块」删掉。 */
const HTML_TAG_NAMES = new Set([
  'br', 'p', 'div', 'span', 'font', 'strong', 'b', 'em', 'i', 'u', 's', 'del', 'ins', 'strike', 'small', 'big',
  'center', 'table', 'thead', 'tbody', 'tr', 'td', 'th', 'ul', 'ol', 'li', 'hr', 'h1', 'h2', 'h3', 'h4', 'h5',
  'h6', 'img', 'a', 'blockquote', 'code', 'pre', 'sup', 'sub', 'summary', 'style', 'script', 'iframe',
  'audio', 'video', 'source', 'section', 'article', 'header', 'footer', 'nav', 'aside', 'figure', 'figcaption',
  'label', 'button', 'input', 'textarea', 'select', 'option', 'time', 'mark', 'marquee',
]);

const escapeRegExp = value => String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

const TAG_NAME_PATTERN = '[A-Za-z0-9_\\u4E00-\\u9FFF~:.\\-]{1,80}';

function normalizeTagName(name) {
  return String(name || '').trim().toLowerCase().replace(/[~-]+$/, '');
}

function parseTagList(value) {
  return String(value || '')
    .split(/[,，;；\s]+/)
    .map(normalizeTagName)
    .filter(Boolean);
}

function isMetaTagName(name, extraNames = new Set(), keepNames = new Set()) {
  const normalized = normalizeTagName(name);
  if (!normalized) return false;
  if (normalized === 'bbi_image' || HTML_TAG_NAMES.has(normalized)) return false;
  if (keepNames.has(normalized)) return false;
  if (extraNames.has(normalized)) return true;
  if (BODY_TAG_NAMES.has(normalized)) return false;
  if (META_TAG_EXACT.has(normalized)) return true;
  return META_TAG_KEYWORDS.some(keyword => normalized.includes(keyword));
}

const BODY_TAG_PATTERN = [...BODY_TAG_NAMES].map(escapeRegExp).join('|');

/** `<正文>…</正文>` / `<story>…</story>` 这类容器：保留内层，不留外壳。 */
function unwrapBodyTags(text) {
  const re = new RegExp(`<(${BODY_TAG_PATTERN})(?:[ \\t][^>]*)?>([\\s\\S]*?)</\\1[^>]*>`, 'gi');
  let result = text;
  for (let depth = 0; depth < 3; depth++) {
    const next = result.replace(re, '$2');
    if (next === result) break;
    result = next;
  }
  return result;
}

/** 时间戳切点落在容器内部时，结尾会剩一个 `</正文>`，这里清掉。 */
function stripBodyWrapperTags(text) {
  return text.replace(new RegExp(`</(?:${BODY_TAG_PATTERN})[^>]*>`, 'gi'), '');
}

/** 整块删除命中的 `<tag>…</tag>`，按范围倒序替换，避免索引错位。 */
function removeMetaBlocks(text, extraNames = [], keepNames = new Set()) {
  const extra = new Set(extraNames);
  const ranges = [];
  const openRe = new RegExp(`<(${TAG_NAME_PATTERN})(?:[ \\t][^>]*)?>`, 'g');
  let match;

  while ((match = openRe.exec(text)) !== null) {
    const name = match[1];
    if (!isMetaTagName(name, extra, keepNames)) continue;
    const closeRe = new RegExp(`</${escapeRegExp(name)}[^>]*>`, 'i');
    const rest = text.slice(openRe.lastIndex);
    const close = closeRe.exec(rest);
    if (!close) continue;
    const end = openRe.lastIndex + close.index + close[0].length;
    ranges.push([match.index, end]);
    openRe.lastIndex = end;
  }

  let result = text;
  for (const [start, end] of ranges.reverse()) {
    result = `${result.slice(0, start)}\n${result.slice(end)}`;
  }
  return result;
}

/**
 * 处理「只有半个标签」的情况：模型常把思维链包在 `<konatan_planning~>` 里，
 * 但流式输出只留下结尾的 `</konatan_planning~>`，开标签丢了。这种孤儿闭合标签
 * 之前的内容整段都是思维链，直接从开头切到闭合标签结束。
 * 反过来，正文末尾只出现 `<tucao>` 而没有闭合，则从该行切到结尾。
 */
function removeOrphanMetaBlocks(text, extraNames = [], keepNames = new Set()) {
  const extra = new Set(extraNames);
  let result = text;

  // 1a) 思维链容器：闭合标签可能粘在文字末尾（`开始写！</konatan_planning~>`），允许行内匹配
  const inlineCloseRe = new RegExp(`</(${TAG_NAME_PATTERN})[^>]*>`, 'gi');
  let bestEnd = -1;
  let match;
  while ((match = inlineCloseRe.exec(result)) !== null) {
    const name = match[1];
    if (!PLANNING_CONTAINER_NAMES.has(normalizeTagName(name))) continue;
    const openRe = new RegExp(`<${escapeRegExp(name)}(?:[ \\t][^>]*)?>`, 'i');
    if (openRe.test(result)) continue;
    const end = match.index + match[0].length;
    if (end <= result.length * 0.7) bestEnd = Math.max(bestEnd, end);
  }
  if (bestEnd > 0) result = result.slice(bestEnd).replace(/^[ \t\r\n]+/, '');

  // 1b) 其他元信息标签：只认独占一行的孤儿闭合标签
  const lineCloseRe = new RegExp(`^[ \\t]*<\\/(${TAG_NAME_PATTERN})>[ \\t]*$`, 'gmi');
  bestEnd = -1;
  while ((match = lineCloseRe.exec(result)) !== null) {
    const name = match[1];
    if (!isMetaTagName(name, extra, keepNames)) continue;
    const openRe = new RegExp(`<${escapeRegExp(name)}(?:[ \\t][^>]*)?>`, 'i');
    if (openRe.test(result)) continue;
    const end = match.index + match[0].length;
    if (end <= result.length * 0.7) bestEnd = Math.max(bestEnd, end);
  }
  if (bestEnd > 0) result = result.slice(bestEnd).replace(/^[ \t\r\n]+/, '');

  // 2) 孤儿开标签：行首出现 `<name>` 但后面没有闭合，且位置靠后（>55%），切到结尾
  const tailLength = result.length;
  const openRe = new RegExp(`^[ \\t]*<(${TAG_NAME_PATTERN})(?:[ \\t][^>]*)?>[ \\t]*$`, 'gmi');
  let cutAt = -1;
  while ((match = openRe.exec(result)) !== null) {
    const name = match[1];
    if (!isMetaTagName(name, extra, keepNames)) continue;
    const closeReOne = new RegExp(`</${escapeRegExp(name)}[^>]*>`, 'i');
    if (closeReOne.test(result.slice(match.index + match[0].length))) continue;
    if (match.index >= tailLength * 0.55) {
      cutAt = cutAt === -1 ? match.index : Math.min(cutAt, match.index);
    }
  }
  if (cutAt >= 0) result = result.slice(0, cutAt).replace(/[ \t\r\n]+$/, '');

  // 3) 落在正文末尾的孤儿闭合标签（`</tucao>`、`</progress>`）：正文已经写完了，
  //    只把标签本身摘掉，不能因为它位置靠后就把前面的正文截掉。
  const standaloneCloseRe = new RegExp(`^[ \\t]*</(${TAG_NAME_PATTERN})>[ \\t]*$`, 'gmi');
  result = result.replace(standaloneCloseRe, (whole, name) => {
    if (!isMetaTagName(name, extra, keepNames)) return whole;
    const paired = new RegExp(`<${escapeRegExp(name)}(?:[ \\t][^>]*)?>`, 'i').test(result);
    return paired ? whole : '';
  });
  const tailCloseRe = new RegExp(`</(${TAG_NAME_PATTERN})>[ \\t]*$`, 'i');
  const tailMatch = tailCloseRe.exec(result);
  if (tailMatch && isMetaTagName(tailMatch[1], extra, keepNames)) {
    const paired = new RegExp(`<${escapeRegExp(tailMatch[1])}(?:[ \\t][^>]*)?>`, 'i').test(result);
    if (!paired) result = result.slice(0, tailMatch.index).replace(/[ \t]+$/, '');
  }

  return result;
}

/** 模型偶尔会把提示词原样复述成一行（如「根据前文，输出下文的大纲」，后面接一段大纲）。
 *  出现在正文尾部时，连它后面的大纲一起切掉；出现在中间时只删这一行。 */
function stripPromptEchoLines(text) {
  const echo = /^(?:根据前文|输出下文|继续(?:写|输出)|请(?:继续|写|输出)|写(?:一段|下文)|字数(?:要求|控制)|以下是(?:正文|续写))/.source;
  const echoRe = new RegExp(echo, 'i');
  const lines = text.split('\n');
  let offset = 0;
  let cutAt = -1;
  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i].trim();
    if (echoRe.test(line) && offset > text.length * 0.8) {
      const rest = lines.slice(i + 1).filter(item => item.trim());
      const looksLikeOutline =
        rest.length >= 3 &&
        rest.every(item => item.length <= 120 && !/[“”「」『』"]/.test(item)) &&
        !rest.some(item => /【[^】\n]{1,80}】|^#{1,6}[ \t]/.test(item));
      if (looksLikeOutline) cutAt = i;
    }
    offset += lines[i].length + 1;
  }
  if (cutAt >= 0) return lines.slice(0, cutAt).join('\n').replace(/[ \t\r\n]+$/, '');
  return lines.filter(line => !echoRe.test(line.trim())).join('\n');
}

/** 没有时间戳时，删掉开头的未知标签块；遇到 `<正文>` 这类容器则解开保留内容。 */
function stripLeadingBlocks(text, keepNames = new Set()) {
  let result = text.replace(/^[ \t\r\n]+/, '');

  for (;;) {
    const open = new RegExp(`^<(${TAG_NAME_PATTERN})(?:[ \\t][^>]*)?>`, 'i').exec(result);
    if (!open) break;
    const name = open[1];
    const normalized = normalizeTagName(name);
    if (normalized === 'bbi_image' || HTML_TAG_NAMES.has(normalized) || keepNames.has(normalized)) break;

    const close = new RegExp(`</${escapeRegExp(name)}[^>]*>`, 'i').exec(result.slice(open[0].length));
    if (!close) break;
    const inner = result.slice(open[0].length, open[0].length + close.index);

    result = BODY_TAG_NAMES.has(normalized)
      ? inner.replace(/^[ \t\r\n]+/, '')
      : result.slice(open[0].length + close.index + close[0].length).replace(/^[ \t\r\n]+/, '');
  }

  return result;
}

/** 删掉结尾的未知标签块（`<current_event>` / `<progress>` 之类），从后往前剥。 */
function stripTrailingBlocks(text, keepNames = new Set()) {
  let result = text.replace(/[ \t\r\n]+$/, '');

  for (;;) {
    const close = new RegExp(`</(${TAG_NAME_PATTERN})[^>]*>[ \\t]*$`, 'i').exec(result);
    if (!close) break;
    const name = close[1];
    const normalized = normalizeTagName(name);
    if (normalized === 'bbi_image' || HTML_TAG_NAMES.has(normalized) || BODY_TAG_NAMES.has(normalized)) break;
    if (keepNames.has(normalized)) break;

    const openRe = new RegExp(`<${escapeRegExp(name)}(?:[ \\t][^>]*)?>`, 'gi');
    let open = null;
    for (const candidate of result.slice(0, close.index).matchAll(openRe)) open = candidate;
    if (!open) break;

    result = result.slice(0, open.index).replace(/[ \t\r\n]+$/, '');
  }

  return result;
}

/**
 * 正文提取（默认开启）：
 * 1. 可选「正文标签」优先取内层；
 * 2. 整块剥掉思维链 / 吐槽 / 状态类标签；
 * 3. 以开头的 `【 时间 】` 行作为正文起点（兼容 `### 【 … 】`）；
 * 4. 没有时间戳就丢掉开头残留的未知标签块；
 * 5. 再剥掉结尾的状态块，压缩多余空行。
 *
 * 时间戳只在开头 800 字符里找，避免正文中间的 `【 场景 】` 把前半段切掉。
 */
function extractBody(raw, options = {}) {
  let text = String(raw ?? '').replace(/\r\n?/g, '\n');
  if (options.bodyOnly === false) return text.trim();

  const keepNames = new Set(parseTagList(options.keepTags));
  const bodyTag = String(options.bodyTag || '').trim().replace(/^<\/?/, '').replace(/>$/, '');
  if (bodyTag) {
    const re = new RegExp(`<${escapeRegExp(bodyTag)}(?:[ \\t][^>]*)?>([\\s\\S]*?)</${escapeRegExp(bodyTag)}[^>]*>`, 'i');
    const wrapped = re.exec(text);
    if (wrapped) text = wrapped[1];
  }

  text = removeOrphanMetaBlocks(text, parseTagList(options.excludeTags), keepNames);
  text = removeMetaBlocks(text, parseTagList(options.excludeTags), keepNames);
  text = unwrapBodyTags(text);

  const head = text.slice(0, 800);
  const stamp = /^[ \t]*(?:#{1,6}[ \t]*)?【[^】\n]{1,80}】[ \t]*$/m.exec(head);
  text = stamp ? text.slice(stamp.index) : stripLeadingBlocks(text, keepNames);

  text = stripBodyWrapperTags(stripTrailingBlocks(text, keepNames));
  text = stripPromptEchoLines(text);
  return text
    .replace(/^(?:[-*_]{3,}[ \t]*\n)+/, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/* ------------------------------------------------------------------ *
 * 4. 标签清单与导出前确认弹窗
 * ------------------------------------------------------------------ */

const DIALOG_ID = 'bbi-export-dialog';
let dialogOpen = false;

function escapeHtml(value) {
  const entities = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
  return String(value ?? '').replace(/[&<>"']/g, character => entities[character]);
}

/**
 * 扫描当前聊天里所有 AI 楼层，列出「最外层」的成对标签。
 * 思维链里嵌套的 `<slot>`、`<nl>` 之类不会单独列出来，避免清单被噪声淹没。
 */
function scanTagInventory(input) {
  const chat = Array.isArray(input) ? input : Array.isArray(input?.chat) ? input.chat : [];
  const tags = new Map();
  let imageBlocks = 0;
  let floors = 0;

  for (const message of chat) {
    if (!isAiStoryMessage(message)) continue;
    floors += 1;
    const mes = String(message.mes || '');
    const firstContent = Math.max(0, mes.search(/\S/));
    const trimmedEnd = mes.replace(/\s+$/, '').length;
    const openRe = new RegExp(`<(${TAG_NAME_PATTERN})(?:[ \\t][^>]*)?>`, 'g');
    let match;

    while ((match = openRe.exec(mes)) !== null) {
      const name = match[1];
      const normalized = normalizeTagName(name);
      const closeRe = new RegExp(`</${escapeRegExp(name)}[^>]*>`, 'i');
      const close = closeRe.exec(mes.slice(openRe.lastIndex));
      if (!close) continue;

      const end = openRe.lastIndex + close.index + close[0].length;
      const inner = mes.slice(openRe.lastIndex, openRe.lastIndex + close.index);

      if (normalized === 'bbi_image') {
        imageBlocks += 1;
        openRe.lastIndex = end;
        continue;
      }
      if (HTML_TAG_NAMES.has(normalized)) continue;
      // 思维链/状态块整块跳过（里面的示例标签不参与清单）；正文容器则继续往里扫描
      if (isMetaTagName(name, new Set())) openRe.lastIndex = end;

      const item = tags.get(normalized) || {
        key: normalized,
        name,
        count: 0,
        floors: new Set(),
        chars: 0,
        sample: '',
        containsStamp: false,
        containsImage: false,
        edge: false,
      };
      item.count += 1;
      item.floors.add(floors);
      item.chars += inner.length;
      item.containsStamp ||= /^[ \t]*(?:#{1,6}[ \t]*)?【[^】\n]{1,80}】[ \t]*$/m.test(inner);
      item.containsImage ||= /<bbi_image/i.test(inner);
      item.edge ||= match.index <= firstContent || end >= trimmedEnd;
      if (!item.sample) {
        const firstLine = inner.trim().split('\n').find(line => line.trim()) || '';
        item.sample = firstLine.replace(/\s+/g, ' ').trim().slice(0, 48);
      }
      tags.set(normalized, item);
    }
  }

  return {
    floors,
    imageBlocks,
    tags: [...tags.values()].sort((a, b) => b.count - a.count),
  };
}

function suggestExclude(item, extraNames = new Set()) {
  if (BODY_TAG_NAMES.has(item.key)) return false;
  if (item.containsStamp) return false;
  if (item.containsImage && !isMetaTagName(item.name, extraNames)) return false;
  if (isMetaTagName(item.name, extraNames)) return true;
  return item.edge;
}

function tagBadges(item) {
  const badges = [];
  if (item.containsStamp) badges.push('含时间行');
  if (item.containsImage) badges.push('含配图');
  if (BODY_TAG_NAMES.has(item.key)) badges.push('正文容器');
  if (isMetaTagName(item.name, new Set())) badges.push('疑似非正文');
  if (item.edge) badges.push('首尾块');
  if (!badges.length) badges.push('其他');
  return badges;
}

/**
 * 导出前弹窗：让用户先看清当前聊天里到底有哪些标签，
 * 再决定「正文从哪个标签里取」以及「哪些标签整块不导出」。
 * 返回 null 表示取消；返回对象会与设置合并成本次导出的选项。
 */
function openExportDialog(ctx, settings) {
  if (dialogOpen) return Promise.resolve(null);
  dialogOpen = true;

  return new Promise(resolve => {
    const inventory = scanTagInventory(ctx.chat);
    const extraExclude = new Set(parseTagList(settings.excludeTags));
    const defaultBodyKey = normalizeTagName(settings.bodyTag);
    const hasDefaultBody = Boolean(defaultBodyKey) && inventory.tags.some(item => item.key === defaultBodyKey);

    const bodyOptions = [
      `<label class="bbi-export-dialog__body-option">
        <input type="radio" name="bbi-body-tag" value="" ${hasDefaultBody ? '' : 'checked'}>
        <span><b>自动识别</b><em>从时间行开始，首尾标签块自动裁掉（推荐）</em></span>
      </label>`,
      ...inventory.tags.map(item => `
        <label class="bbi-export-dialog__body-option">
          <input type="radio" name="bbi-body-tag" value="${escapeHtml(item.name)}" ${item.key === defaultBodyKey ? 'checked' : ''}>
          <span><code>&lt;${escapeHtml(item.name)}&gt;</code><em>${item.count} 次${item.containsStamp ? ' · 含时间行' : ''}${item.containsImage ? ' · 含配图' : ''}</em></span>
        </label>`),
    ].join('');

    const tagRows = inventory.tags.length
      ? inventory.tags.map(item => `
        <label class="bbi-export-dialog__tag">
          <input type="checkbox" data-exclude="${escapeHtml(item.name)}" ${item.key === defaultBodyKey ? 'disabled' : suggestExclude(item, extraExclude) ? 'checked' : ''}>
          <span class="bbi-export-dialog__tag-main">
            <span class="bbi-export-dialog__tag-title"><code>&lt;${escapeHtml(item.name)}&gt;</code><em>${item.count} 次 · ${item.chars} 字</em></span>
            <span class="bbi-export-dialog__badges">${tagBadges(item).map(badge => `<span>${escapeHtml(badge)}</span>`).join('')}</span>
            ${item.sample ? `<span class="bbi-export-dialog__sample">${escapeHtml(item.sample)}</span>` : ''}
          </span>
        </label>`).join('')
      : '<p class="bbi-export-dialog__empty">当前聊天的 AI 楼层里没有检测到成对标签，将直接按时间行截取正文。</p>';

    const overlay = document.createElement('div');
    overlay.id = DIALOG_ID;
    overlay.className = 'bbi-export-dialog';
    overlay.innerHTML = `
      <div class="bbi-export-dialog__backdrop" data-action="backdrop"></div>
      <div class="bbi-export-dialog__card" role="dialog" aria-modal="true" aria-label="导出前确认正文标签">
        <header class="bbi-export-dialog__head">
          <b>导出前确认：只取哪部分正文</b>
          <p>当前聊天 <strong>${inventory.floors}</strong> 个 AI 楼层，检测到 <strong>${inventory.tags.length}</strong> 种标签；配图标签 <code>&lt;bbi_image&gt;</code> ${inventory.imageBlocks} 个，始终保留。</p>
        </header>

        <label class="bbi-export-dialog__switch">
          <input type="checkbox" id="bbi-dialog-body-only" ${settings.bodyOnly === false ? '' : 'checked'}>
          <span>只导出正文（去掉思维链、吐槽、状态块）</span>
        </label>

        <fieldset class="bbi-export-dialog__fields">
          <section class="bbi-export-dialog__section">
            <h4>1. 导出格式</h4>
            <div class="bbi-export-dialog__format">
              <label class="bbi-export-dialog__choice">
                <input type="radio" name="bbi-format" value="docx" ${settings.format === 'epub' ? '' : 'checked'}>
                <span><b>Word</b><em>.docx · 图文混排，适合打印、存档、发给人看</em></span>
              </label>
              <label class="bbi-export-dialog__choice">
                <input type="radio" name="bbi-format" value="epub" ${settings.format === 'epub' ? 'checked' : ''}>
                <span><b>EPUB</b><em>.epub · 手机阅读器/电纸书，流式翻页</em></span>
              </label>
            </div>
          </section>

          <section class="bbi-export-dialog__section">
            <h4>2. 导出楼层范围</h4>
            <div class="bbi-export-dialog__range">
              <label>从 <input type="number" id="bbi-floor-start" min="1" max="${Math.max(1, inventory.floors)}" value="1" inputmode="numeric"> 层</label>
              <span>到</span>
              <label><input type="number" id="bbi-floor-end" min="1" max="${Math.max(1, inventory.floors)}" value="${Math.max(1, inventory.floors)}" inputmode="numeric"> 层</label>
              <button type="button" class="menu_button" data-action="range-all">全部</button>
            </div>
            <p class="bbi-export-dialog__section-note">当前共 <b>${inventory.floors}</b> 个 AI 楼层（不含你的发言）。只导某一层时，起止填同一个数字。</p>
            <p class="bbi-export-dialog__range-error" hidden></p>
          </section>

          <section class="bbi-export-dialog__section">
            <h4>3. 正文从哪个标签里取</h4>
            <p class="bbi-export-dialog__section-note">一般保持「自动识别」；只有角色卡把正文固定包在某个标签里时，才需要选下面的标签。</p>
            <div class="bbi-export-dialog__body-list">${bodyOptions}</div>
          </section>

          <section class="bbi-export-dialog__section">
            <h4>4. 整块不导出（勾选的标签会被删除）</h4>
            <div class="bbi-export-dialog__bulk">
              <button type="button" class="menu_button" data-action="check-all">全选</button>
              <button type="button" class="menu_button" data-action="uncheck-all">全不选</button>
            </div>
            <div class="bbi-export-dialog__tags">${tagRows}</div>
          </section>
        </fieldset>

        <footer class="bbi-export-dialog__foot">
          <span>未勾选的标签会按普通文字保留；取消则不会导出。</span>
          <div class="bbi-export-dialog__buttons">
            <button type="button" class="menu_button" data-action="cancel">取消</button>
            <button type="button" class="menu_button" data-action="confirm">开始导出</button>
          </div>
        </footer>
      </div>`;

    document.body.appendChild(overlay);

    const bodyOnly = overlay.querySelector('#bbi-dialog-body-only');
    const floorStart = overlay.querySelector('#bbi-floor-start');
    const floorEnd = overlay.querySelector('#bbi-floor-end');
    const rangeError = overlay.querySelector('.bbi-export-dialog__range-error');
    const totalFloors = Math.max(1, inventory.floors);

    const readRange = () => {
      const start = Number.parseInt(floorStart.value, 10);
      const end = Number.parseInt(floorEnd.value, 10);
      const fail = message => {
        rangeError.textContent = message;
        rangeError.hidden = false;
        return null;
      };
      if (!Number.isInteger(start) || !Number.isInteger(end)) {
        return fail('请填写整数楼层号。');
      }
      if (start < 1 || end < 1 || start > totalFloors || end > totalFloors) {
        return fail(`楼层号必须在 1–${totalFloors} 之间。`);
      }
      if (start > end) {
        return fail('起始层不能大于结束层；只导一层时请把起止填成同一个数字。');
      }
      rangeError.hidden = true;
      rangeError.textContent = '';
      return { floorStart: start, floorEnd: end };
    };

    overlay.querySelector('[data-action="range-all"]')?.addEventListener('click', () => {
      floorStart.value = '1';
      floorEnd.value = String(totalFloors);
      readRange();
    });
    for (const input of [floorStart, floorEnd]) {
      input.addEventListener('input', () => {
        if (!rangeError.hidden) readRange();
      });
    }

    const syncBodyTag = () => {
      const chosen = overlay.querySelector('input[name="bbi-body-tag"]:checked')?.value || '';
      for (const box of overlay.querySelectorAll('input[data-exclude]')) {
        const isBody = box.dataset.exclude === chosen;
        box.disabled = !bodyOnly.checked || isBody;
        if (isBody) box.checked = false;
      }
    };
    const syncBodyOnly = () => {
      overlay.classList.toggle('is-body-off', !bodyOnly.checked);
      for (const input of overlay.querySelectorAll('input[name="bbi-body-tag"]')) {
        input.disabled = !bodyOnly.checked;
      }
      syncBodyTag();
    };
    syncBodyOnly();
    bodyOnly.addEventListener('change', syncBodyOnly);
    overlay.addEventListener('change', event => {
      if (event.target.name === 'bbi-body-tag') syncBodyTag();
    });
    syncBodyTag();

    overlay.querySelector('[data-action="check-all"]')?.addEventListener('click', () => {
      for (const box of overlay.querySelectorAll('input[data-exclude]:not(:disabled)')) box.checked = true;
    });
    overlay.querySelector('[data-action="uncheck-all"]')?.addEventListener('click', () => {
      for (const box of overlay.querySelectorAll('input[data-exclude]:not(:disabled)')) box.checked = false;
    });

    let settled = false;
    const finish = value => {
      if (settled) return;
      settled = true;
      dialogOpen = false;
      document.removeEventListener('keydown', onKey);
      overlay.remove();
      resolve(value);
    };
    const onKey = event => {
      if (event.key === 'Escape') finish(null);
    };
    document.addEventListener('keydown', onKey);

    overlay.querySelector('[data-action="backdrop"]')?.addEventListener('click', () => finish(null));
    overlay.querySelector('[data-action="cancel"]')?.addEventListener('click', () => finish(null));
    overlay.querySelector('[data-action="confirm"]')?.addEventListener('click', () => {
      const format = overlay.querySelector('input[name="bbi-format"]:checked')?.value === 'epub' ? 'epub' : 'docx';
      const range = readRange();
      if (!range) {
        rangeError.scrollIntoView({ block: 'nearest' });
        return;
      }
      if (!bodyOnly.checked) {
        finish({ bodyOnly: false, format, ...range });
        return;
      }
      const chosenBody = overlay.querySelector('input[name="bbi-body-tag"]:checked')?.value || '';
      const boxes = [...overlay.querySelectorAll('input[data-exclude]')];
      const excluded = boxes.filter(box => box.checked && !box.disabled).map(box => box.dataset.exclude);
      const kept = boxes
        .map(box => box.dataset.exclude)
        .filter(key => !excluded.includes(key) && key !== chosenBody);
      finish({
        bodyOnly: true,
        bodyTag: chosenBody,
        excludeTags: excluded.join(', '),
        keepTags: kept.join(', '),
        format,
        ...range,
      });
    });
  });
}

/* ------------------------------------------------------------------ *
 * 5. 收集楼层
 * ------------------------------------------------------------------ */

function collectFloors(ctx, options = {}) {
  const floors = [];
  const chat = Array.isArray(ctx.chat) ? ctx.chat : [];

  for (let messageId = 0; messageId < chat.length; messageId++) {
    const message = chat[messageId];
    if (!isAiStoryMessage(message)) continue;

    const swipeId = Number.isInteger(message.swipe_id) ? message.swipe_id : 0;
    const store = readStore(message, swipeId);
    const bodyText = extractBody(message.mes, options);
    const tags = findImageTags(bodyText);
    const blocks = [];
    let cursor = 0;
    let imageCount = 0;

    for (let seq = 0; seq < tags.length; seq++) {
      const tag = tags[seq];
      const before = sanitizeMessageText(bodyText.slice(cursor, tag.index));
      if (before) blocks.push({ type: 'text', text: before });
      const entry = pickImageEntry(store, swipeId, tag.raw, seq);
      if (entry) {
        blocks.push({ type: 'image', path: entry.path, rawTag: tag.raw });
        imageCount += 1;
      }
      cursor = tag.end;
    }

    const tail = sanitizeMessageText(bodyText.slice(cursor));
    if (tail) blocks.push({ type: 'text', text: tail });
    if (!blocks.length) continue;

    floors.push({
      messageId,
      name: String(message.name || '').trim(),
      date: String(message.send_date || '').trim(),
      blocks,
      imageCount,
    });
  }

  return floors;
}

/* ------------------------------------------------------------------ *
 * 4. 读取图片（ST 静态目录 → 二进制；可选压缩）
 * ------------------------------------------------------------------ */

function sniffMime(bytes) {
  if (bytes.length >= 8 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) return 'image/png';
  if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return 'image/jpeg';
  if (bytes.length >= 6 && bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) return 'image/gif';
  if (
    bytes.length >= 12 &&
    bytes[0] === 0x52 && bytes[1] === 0x49 && bytes[2] === 0x46 && bytes[3] === 0x46 &&
    bytes[8] === 0x57 && bytes[9] === 0x45 && bytes[10] === 0x42 && bytes[11] === 0x50
  ) return 'image/webp';
  return 'image/png';
}

const WORD_MIMES = new Set(['image/png', 'image/jpeg', 'image/gif']);

function loadBitmap(bytes, mime) {
  const blob = new Blob([bytes], { type: mime || 'application/octet-stream' });
  if (typeof createImageBitmap === 'function') {
    return createImageBitmap(blob);
  }
  return new Promise((resolve, reject) => {
    const image = new Image();
    const url = URL.createObjectURL(blob);
    image.onload = () => {
      URL.revokeObjectURL(url);
      resolve(image);
    };
    image.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error('图片解码失败'));
    };
    image.src = url;
  });
}

function scaleToMaxEdge(width, height, maxEdge) {
  if (!maxEdge || maxEdge <= 0) return { width, height };
  const scale = Math.min(1, maxEdge / Math.max(width, height));
  return {
    width: Math.max(1, Math.round(width * scale)),
    height: Math.max(1, Math.round(height * scale)),
  };
}

function canvasToBlob(canvas, mime, quality) {
  return new Promise((resolve, reject) => {
    canvas.toBlob(
      blob => {
        if (blob) resolve(blob);
        else reject(new Error('图片编码失败'));
      },
      mime,
      quality,
    );
  });
}

/**
 * 读一张图。默认压成 JPEG（保留白底），文件体积约为原图 1/5–1/10；
 * 关闭压缩时保留原格式，只有 Word 不认的格式（webp/bmp…）才转 PNG。
 */
async function loadImageAsset(path, compress) {
  const response = await fetch(path, { credentials: 'same-origin' });
  if (!response.ok) throw new Error(`HTTP ${response.status}`);
  const bytes = new Uint8Array(await response.arrayBuffer());
  if (!bytes.length) throw new Error('图片为空');

  let mime = (response.headers.get('content-type') || '').split(';')[0].trim().toLowerCase();
  if (!mime || mime === 'application/octet-stream') mime = sniffMime(bytes);

  const bitmap = await loadBitmap(bytes, mime);
  try {
    const width = Number(bitmap.width) || 0;
    const height = Number(bitmap.height) || 0;

    if (!compress && WORD_MIMES.has(mime)) {
      return { data: bytes, mime, width, height };
    }

    const target = compress ? scaleToMaxEdge(width, height, MAX_IMAGE_EDGE) : { width, height };
    const canvas = document.createElement('canvas');
    canvas.width = target.width;
    canvas.height = target.height;
    const context = canvas.getContext('2d');
    if (!context) throw new Error('canvas 2D 不可用');
    // JPEG 没有透明通道：先铺白底，否则透明区域在 Word 里会变黑
    context.fillStyle = '#ffffff';
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.drawImage(bitmap, 0, 0, canvas.width, canvas.height);

    const outputMime = compress ? 'image/jpeg' : 'image/png';
    const blob = await canvasToBlob(canvas, outputMime, compress ? JPEG_QUALITY : undefined);
    return {
      data: new Uint8Array(await blob.arrayBuffer()),
      mime: outputMime,
      width: canvas.width,
      height: canvas.height,
    };
  } finally {
    bitmap.close?.();
  }
}

async function mapLimit(items, limit, worker) {
  let next = 0;
  const runners = Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (true) {
      const index = next++;
      if (index >= items.length) return;
      await worker(items[index], index);
    }
  });
  await Promise.all(runners);
}

/* ------------------------------------------------------------------ *
 * 5. 导出主流程
 * ------------------------------------------------------------------ */

function formatDateTime(date) {
  const pad = value => String(value).padStart(2, '0');
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}`;
}

function formatStamp(date) {
  const pad = value => String(value).padStart(2, '0');
  return `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}-${pad(date.getHours())}${pad(date.getMinutes())}`;
}

function safeFileName(name) {
  return (
    String(name || '')
      .replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_')
      .replace(/\s+/g, ' ')
      .trim() || '聊天'
  );
}

function downloadFile(bytes, filename, mime) {
  const blob = new Blob([bytes], { type: mime });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  link.rel = 'noopener';
  document.body.appendChild(link);
  link.click();
  link.remove();
  // 立刻 revoke 会让部分浏览器的下载中断，留一分钟再回收
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

/**
 * 按 1-based 的 AI 楼层序号截取；起止越界会被夹到合法范围，起止颠倒则交换。
 * 这里的“层”是当前聊天里 AI 回复的顺序，不是 ST 的 messageId。
 */
function selectFloorRange(floors, start, end) {
  const list = Array.isArray(floors) ? floors : [];
  const total = list.length;
  if (!total) return [];
  let from = Number.parseInt(start, 10);
  let to = Number.parseInt(end, 10);
  if (!Number.isInteger(from)) from = 1;
  if (!Number.isInteger(to)) to = total;
  from = Math.min(Math.max(from, 1), total);
  to = Math.min(Math.max(to, 1), total);
  if (from > to) [from, to] = [to, from];
  return list.slice(from - 1, to);
}

function toDocBlocks(floors, assets, options) {
  const blocks = [];
  const stats = { floors: floors.length, images: 0, failed: 0 };

  for (const floor of floors) {
    if (options.showFloorName) {
      blocks.push({ type: 'floor', name: floor.name, meta: options.showTime ? floor.date : '' });
    }
    for (const item of floor.blocks) {
      if (item.type === 'text') {
        blocks.push({ type: 'text', text: item.text });
        continue;
      }
      const asset = assets.get(item.path);
      if (!asset) {
        blocks.push({ type: 'note', text: '〔本张配图读取失败，已跳过〕' });
        stats.failed += 1;
        continue;
      }
      blocks.push({
        type: 'image',
        data: asset.data,
        mime: asset.mime,
        width: asset.width,
        height: asset.height,
        caption: options.includePrompt ? promptText(item.rawTag) : '',
      });
      stats.images += 1;
    }
  }

  return { blocks, stats };
}

async function exportAll() {
  if (busy) return;
  const ctx = getContext();
  if (!ctx || !Array.isArray(ctx.chat) || !ctx.chat.length) {
    toast('warning', '当前没有可导出的聊天。');
    return;
  }

  const settings = readSettings(ctx);
  const choice = await openExportDialog(ctx, settings);
  if (!choice) return;

  const options = { ...settings, ...choice };
  busy = true;
  setBusyUi(true, '正在收集楼层…');

  try {
    const allFloors = collectFloors(ctx, options);
    const floors = selectFloorRange(allFloors, options.floorStart, options.floorEnd)
      .filter(floor => !options.onlyWithImages || floor.imageCount > 0);
    if (!floors.length) {
      const rangeHint = allFloors.length ? `第 ${options.floorStart || 1}–${options.floorEnd || allFloors.length} 层里` : '';
      toast('warning', options.onlyWithImages ? `${rangeHint}没有找到带配图的 AI 楼层。` : `${rangeHint}没有找到可导出的 AI 楼层。`);
      return;
    }

    const paths = new Set();
    for (const floor of floors) {
      for (const item of floor.blocks) {
        if (item.type === 'image') paths.add(item.path);
      }
    }
    const uniquePaths = [...paths];
    const assets = new Map();
    let loaded = 0;

    if (uniquePaths.length) {
      setBusyUi(true, `正在读取配图 0/${uniquePaths.length}`);
      await mapLimit(uniquePaths, FETCH_CONCURRENCY, async path => {
        try {
          assets.set(path, await loadImageAsset(path, options.compressImages));
        } catch (error) {
          console.warn('[BBI 导出] 配图读取失败：', path, error);
          assets.set(path, null);
        } finally {
          loaded += 1;
          setBusyUi(true, `正在读取配图 ${loaded}/${uniquePaths.length}`);
        }
      });
    }

    const format = options.format === 'epub' ? 'epub' : 'docx';
    setBusyUi(true, format === 'epub' ? '正在生成 EPUB…' : '正在生成 Word…');
    const { blocks, stats } = toDocBlocks(floors, assets, options);

    const now = new Date();
    const characterName = String(ctx.name2 || '').trim() || floors[0]?.name || '聊天';
    const title = `${characterName} · 图文导出`;
    const subtitle = `${stats.floors} 个 AI 楼层 · ${stats.images} 张配图 · ${formatDateTime(now)}`;

    const bytes = format === 'epub'
      ? buildEpub({ title, subtitle, blocks, createdAt: now })
      : buildDocx({ title, subtitle, blocks, createdAt: now });
    const extension = format === 'epub' ? '.epub' : '.docx';
    const filename = `柏宝绘_${safeFileName(characterName)}_图文_${formatStamp(now)}${extension}`;
    downloadFile(bytes, filename, format === 'epub' ? EPUB_MIME : DOCX_MIME);

    const formatLabel = format === 'epub' ? 'EPUB' : 'Word';
    if (stats.failed) {
      toast('warning', `已导出 ${stats.floors} 个 AI 楼层、${stats.images} 张配图（${formatLabel}）；${stats.failed} 张图片读取失败已跳过。`);
    } else {
      toast('success', `已导出 ${stats.floors} 个 AI 楼层、${stats.images} 张配图（${formatLabel}）。`);
    }
    if (typeof window.toastr?.info === 'function') {
      window.toastr.info(`文件：${filename}`, '柏宝绘 · 图文导出', { timeOut: 6000 });
    }
  } catch (error) {
    console.error('[BBI 导出] 导出失败', error);
    toast('error', `导出失败：${error?.message || error}`);
  } finally {
    busy = false;
    setBusyUi(false, '');
  }
}

/* ------------------------------------------------------------------ *
 * 6. UI：扩展设置面板 + 魔杖菜单
 * ------------------------------------------------------------------ */

function setBusyUi(running, text) {
  const button = document.getElementById(RUN_BUTTON_ID);
  if (button) {
    button.disabled = running;
    button.textContent = running ? text || '导出中…' : '导出 Word / EPUB';
  }
  const status = document.getElementById(STATUS_ID);
  if (status) status.textContent = text || '';

  const menuItem = document.getElementById(MENU_ITEM_ID);
  if (menuItem) {
    menuItem.style.opacity = running ? '0.55' : '';
    menuItem.style.pointerEvents = running ? 'none' : '';
  }
}

function injectSettingsPanel() {
  if (document.getElementById(PANEL_ID)) return true;
  const host = document.getElementById('extensions_settings2') || document.getElementById('extensions_settings');
  if (!host) return false;

  const panel = document.createElement('div');
  panel.id = PANEL_ID;
  panel.className = 'bbi-export-panel';
  panel.innerHTML = `
    <div class="bbi-export-panel__head">
      <b>柏宝绘 · 图文导出</b>
      <span>把当前聊天里所有 AI 楼层的正文与配图导出成一份 Word</span>
    </div>
    <div class="bbi-export-panel__options">
      <label><input type="checkbox" data-setting="bodyOnly"> 只导出正文（自动去掉思维链、吐槽、状态块）</label>
      <label class="bbi-export-panel__field">
        <span>正文标签（可选）</span>
        <input type="text" data-setting="bodyTag" placeholder="留空自动识别，例如：正文">
      </label>
      <label class="bbi-export-panel__field">
        <span>额外排除的标签</span>
        <input type="text" data-setting="excludeTags" placeholder="逗号分隔，例如：konatan_planning, tucao">
      </label>
      <label><input type="checkbox" data-setting="showFloorName"> 显示楼层角色名</label>
      <label><input type="checkbox" data-setting="showTime"> 显示楼层时间</label>
      <label><input type="checkbox" data-setting="includePrompt"> 在图片下方附上出图提示词</label>
      <label><input type="checkbox" data-setting="onlyWithImages"> 只导出带配图的楼层</label>
      <label><input type="checkbox" data-setting="compressImages"> 压缩图片以减小文档体积（推荐）</label>
    </div>
    <div class="bbi-export-panel__actions">
      <button id="${RUN_BUTTON_ID}" class="menu_button" type="button">导出 Word / EPUB</button>
      <span id="${STATUS_ID}" class="bbi-export-status"></span>
    </div>
  `;

  const settings = readSettings();
  for (const input of panel.querySelectorAll('input[data-setting]')) {
    const key = input.getAttribute('data-setting');
    if (input.type === 'checkbox') {
      input.checked = settings[key] === true;
      input.addEventListener('change', () => writeSetting(key, input.checked));
    } else {
      input.value = typeof settings[key] === 'string' ? settings[key] : '';
      input.addEventListener('change', () => writeSetting(key, input.value.trim()));
    }
  }

  panel.querySelector(`#${RUN_BUTTON_ID}`)?.addEventListener('click', () => {
    void exportAll();
  });

  host.appendChild(panel);
  return true;
}

function injectMenuButton() {
  if (document.getElementById(MENU_ITEM_ID)) return true;
  const menu = document.getElementById('extensionsMenu');
  if (!menu) return false;

  const container = document.createElement('div');
  container.className = 'extension_container interactable';
  container.tabIndex = 0;

  const link = document.createElement('a');
  link.id = MENU_ITEM_ID;
  link.className = 'list-group-item';
  link.href = '#';
  link.title = '把当前聊天的 AI 楼层（正文 + 柏宝绘配图）导出为 Word 或 EPUB';
  link.innerHTML = '<i class="fa-solid fa-file-word"></i><span>导出图文 Word</span>';
  link.addEventListener('click', event => {
    event.preventDefault();
    menu.style.display = 'none';
    void exportAll();
  });

  container.appendChild(link);
  menu.appendChild(container);
  return true;
}

function toast(level, message) {
  const toastr = window.toastr;
  const method = { success: 'success', warning: 'warning', error: 'error', info: 'info' }[level] || 'info';
  if (toastr && typeof toastr[method] === 'function') {
    toastr[method](message, '柏宝绘 · 图文导出', { timeOut: level === 'error' ? 8000 : 5000 });
    return;
  }
  console.log(`[BBI 导出] ${message}`);
  showFallbackToast(message, level);
}

/** ST 没加载 toastr 时的兜底：一个自动消失的小条，不阻塞操作。 */
function showFallbackToast(message, level) {
  const host = document.createElement('div');
  host.textContent = message;
  host.style.cssText = [
    'position:fixed',
    'right:16px',
    'bottom:16px',
    'z-index:2147483647',
    'max-width:360px',
    'padding:10px 14px',
    'border-radius:8px',
    'font-size:13px',
    'line-height:1.5',
    'color:#fff',
    'box-shadow:0 6px 20px rgba(0,0,0,.25)',
    `background:${level === 'error' ? '#8b2f2f' : level === 'warning' ? '#7a5a1f' : '#2f5d3a'}`,
  ].join(';');
  document.body.appendChild(host);
  setTimeout(() => host.remove(), 6000);
}

/* ------------------------------------------------------------------ *
 * 7. 启动
 * ------------------------------------------------------------------ */

function registerSlashCommand() {
  try {
    const { SlashCommandParser, SlashCommand } = window;
    if (!SlashCommandParser || !SlashCommand || typeof SlashCommand.fromProps !== 'function') return false;
    SlashCommandParser.addCommandObject(
      SlashCommand.fromProps({
        name: 'bbi-export-docx',
        helpString: '把当前聊天的 AI 楼层（正文 + 柏宝绘配图）导出为 Word 或 EPUB 文档。',
        callback: async () => {
          await exportAll();
          return '';
        },
      }),
    );
    return true;
  } catch (error) {
    console.warn('[BBI 导出] 斜杠命令注册失败（不影响按钮）', error);
    return false;
  }
}

let commandsRegistered = false;

function init() {
  if (initialized) return;
  const ctx = getContext();
  if (!ctx) return;

  if (!commandsRegistered) commandsRegistered = registerSlashCommand();

  const panelReady = injectSettingsPanel();
  const menuReady = injectMenuButton();
  if (panelReady && menuReady) initialized = true;
}

(function start() {
  // node 单测会 import 本文件测纯函数；没有 document 时直接跳过 UI 启动。
  if (typeof document === 'undefined') return;

  init();

  // ST 的扩展设置面板与魔杖菜单都是异步就绪的，用户也可能稍后才展开扩展面板：
  // 用 MutationObserver 监听插入，比固定次数轮询更稳，也不会在流式输出时反复空转。
  let scheduled = false;
  const scheduleInit = () => {
    if (scheduled || initialized) return;
    scheduled = true;
    setTimeout(() => {
      scheduled = false;
      init();
      if (initialized) observer.disconnect();
    }, 800);
  };

  const observer = new MutationObserver(scheduleInit);
  const observeRoot = document.body || document.documentElement;
  if (observeRoot) observer.observe(observeRoot, { childList: true, subtree: true });

  if (document.readyState === 'loading') {
    document.addEventListener(
      'DOMContentLoaded',
      () => {
        init();
        scheduleInit();
      },
      { once: true },
    );
  }
})();

/**
 * 纯函数导出：给 tests/ 下的 node 单测用。
 * SillyTavern 只会加载本文件，不会读这些 export，不影响插件行为。
 */
export {
  collectFloors,
  extractBody,
  findImageTags,
  isAiStoryMessage,
  openExportDialog,
  parseSlotId,
  pickImageEntry,
  promptHash,
  promptText,
  readSettings,
  scanTagInventory,
  selectFloorRange,
  suggestExclude,
};
