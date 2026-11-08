/* count-by-msword-obsidian 1.2.3. No runtime dependencies other than Obsidian/Node.
 * Markdown conversion revision is part of every cache key. DOCX bytes are shared
 * by count and export; Word alone supplies the final word count.
 */
"use strict";

const { Plugin, MarkdownView, PluginSettingTab, Setting, Notice, Modal, Platform } = require("obsidian");
const fs = require("fs");
const fsp = fs.promises;
const path = require("path");
const os = require("os");
const crypto = require("crypto");
const { spawn } = require("child_process");
const { deflateRaw } = require("zlib");
const { promisify } = require("util");
const deflateZip = promisify(deflateRaw);

const CONVERTER_VERSION = "1.2.1-1";
const STATISTICS_VERSION = 1;
const STATISTICS = Object.freeze([
  ["words", "字数"], ["characters", "字符数（不计空格）"],
  ["charactersWithSpaces", "字符数（计空格）"], ["paragraphs", "段落数"],
  ["lines", "行数"], ["pages", "页数"], ["farEastCharacters", "亚洲语言字符数"]
]);
function validStatistics(value) {
  return value && typeof value === "object" && STATISTICS.every(([key]) =>
    (key !== "words" && value[key] === null) || (Number.isSafeInteger(value[key]) && value[key] >= 0));
}
function statisticsErrors(value) {
  const errors = {};
  if (value && typeof value === "object") for (const [key] of STATISTICS) {
    if (typeof value[key] === "string") errors[key] = value[key];
  }
  return errors;
}
const DEFAULTS = Object.freeze({
  autoCount: true, updateOnReturn: true, showTime: false, exportDirectory: "",
  chineseFont: "宋体", englishFont: "Times New Roman", fontSize: 12,
  lineSpacing: 1.5, firstLineIndent: 2, includeYaml: false,
  punctuationMode: "full", keepBlankLines: false,
  timeoutSeconds: 30, detailedErrors: false
});
const CONTENT_SETTINGS = ["chineseFont", "englishFont", "fontSize", "lineSpacing", "firstLineIndent", "includeYaml", "punctuationMode", "keepBlankLines"];
const RULES = [
  "统计和导出使用同一份 DOCX 转换规则；字数只取 Microsoft Word 的 ComputeStatistics(0)。",
  "悬停状态栏可查看 Word 返回的全部七项统计，并点击导出 DOCX。行数和页数按当前导出字体、字号、行距及 A4 页面排版计算；个别语言统计不可用时会注明原因。",
  "默认排除开头的 YAML；开启“包含 YAML”后将其作为原样文本加入正文。",
  "标点可选择中文全角或英文半角，默认全角；中文引号及全角标点使用中文字体。只转换常用标点，不转换字母和数字；代码、公式、YAML、网址、邮箱和数字格式保留原样。",
  "默认移除 Markdown 段落之间的空白行，各段仍独立成段；可开启“保留空行”。代码块、公式及 YAML 内的空行始终保留。",
  "链接保留显示文字；WikiLink 无别名时保留目标文字。嵌入笔记不展开，保留别名或目标名称。图片不嵌入，保留 alt 文本；无 alt 的 Markdown 图片不产生正文。",
  "正文标签保留 #标签。代码块和行内代码保留代码文字。Callout 保留类型、标题和正文。Obsidian %% 注释及 HTML 注释排除。",
  "脚注引用显示为 [编号]；脚注定义按原顺序完整放入文末正文（非 Word 脚注），因此属于正文统计范围。未定义的引用保留原文。",
  "LaTeX 去掉外围 $ 或 $$，公式源码作为正文；表格按行保留各单元格文字。任务列表使用 ☐ / ☑。",
  "支持标题、粗斜体、删除线、列表、引用、分隔线、换行、Unicode。HTML 采用简化文本转换，忽略 script/style；其他复杂语法尽量原样保留，导出时提示。",
  "Word 不统计图片、文本框及真正脚注：本插件不生成这些对象。Word 中请比较“字数”而非“字符数”，并保持导出后的文档内容不变。",
  "仅支持 Windows 桌面版 Word。任务串行执行；超时仅清理已确认属于本插件的 Word 进程。若 COM 激活在报告进程身份前卡住，插件会终止辅助进程并保留未确认归属的进程。"
];

function digest(value) { return crypto.createHash("sha256").update(value, "utf8").digest("hex"); }
function conversionKey(settings) {
  return digest(JSON.stringify([CONVERTER_VERSION, ...CONTENT_SETTINGS.map(k => settings[k])]));
}
function cacheId(snapshot) { return JSON.stringify([snapshot.filePath, snapshot.contentHash, snapshot.conversionKey]); }
function validatedSettings(raw) {
  const s = { ...DEFAULTS };
  if (!raw || typeof raw !== "object") return s;
  for (const k of Object.keys(s)) {
    if (typeof raw[k] === typeof s[k]) s[k] = raw[k];
  }
  for (const [k, min, max] of [["fontSize", 6, 72], ["lineSpacing", 1, 4], ["firstLineIndent", 0, 8], ["timeoutSeconds", 5, 180]]) {
    s[k] = Number.isFinite(s[k]) ? Math.min(max, Math.max(min, s[k])) : DEFAULTS[k];
  }
  s.chineseFont = s.chineseFont.trim().slice(0, 100) || DEFAULTS.chineseFont;
  s.englishFont = s.englishFont.trim().slice(0, 100) || DEFAULTS.englishFont;
  if (!["full", "half"].includes(s.punctuationMode)) s.punctuationMode = DEFAULTS.punctuationMode;
  return s;
}
function xml(text) {
  let clean = "";
  for (const char of String(text)) {
    const n = char.codePointAt(0);
    if (n === 9 || n === 10 || n === 13 || (n >= 32 && n <= 0xd7ff) ||
        (n >= 0xe000 && n <= 0xfffd) || (n >= 0x10000 && n <= 0x10ffff)) clean += char;
    else if (n >= 0xd800 && n <= 0xdfff) clean += "\ufffd";
  }
  return clean.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}
function decodeEntities(text) {
  const names = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", nbsp: "\u00a0", copy: "©", reg: "®", hellip: "…", mdash: "—", ndash: "–", laquo: "«", raquo: "»" };
  return text.replace(/&(#x[\da-f]+|#\d+|[a-z]+);/gi, (all, entity) => {
    if (entity[0] !== "#") return names[entity] || all;
    const n = /^#x/i.test(entity) ? parseInt(entity.slice(2), 16) : Number(entity.slice(1));
    return n > 0 && n <= 0x10ffff && !(n >= 0xd800 && n <= 0xdfff) ? String.fromCodePoint(n) : "\ufffd";
  });
}
function pushText(runs, text, format) {
  if (!text) return;
  const previous = runs[runs.length - 1];
  if (previous && JSON.stringify(previous.format) === JSON.stringify(format)) previous.text += text;
  else runs.push({ text, format: { ...format } });
}
const FULL_PUNCTUATION = Object.freeze({
  ",": "，", ".": "。", ":": "：", ";": "；", "!": "！", "?": "？",
  "(": "（", ")": "）", "[": "［", "]": "］", "{": "｛", "}": "｝", "<": "＜", ">": "＞"
});
const HALF_PUNCTUATION = Object.freeze({
  "，": ",", "。": ".", "．": ".", "：": ":", "；": ";", "！": "!", "？": "?",
  "（": "(", "）": ")", "［": "[", "］": "]", "【": "[", "】": "]", "〔": "[", "〕": "]",
  "｛": "{", "｝": "}", "＜": "<", "＞": ">", "《": "<", "》": ">", "〈": "<", "〉": ">",
  "“": '"', "”": '"', "＂": '"', "‘": "'", "’": "'", "＇": "'", "「": '"', "」": '"', "『": "'", "』": "'"
});
function typographyRuns(runs, settings, quotes) {
  const result = [], text = runs.map(run => run.text).join("");
  // Protect visible URLs, email addresses, local paths and numeric notation,
  // even when a Markdown emphasis boundary splits a token into several runs.
  const protectedRanges = [];
  const technical = /(?:https?:\/\/|mailto:|www\.)[^\s<>"“”]+|[\p{L}\p{N}._%+-]+@[\p{L}\p{N}.-]+\.[\p{L}]{2,}|[a-z]:[\\/][^\s"“”]+|[+-]?\d+(?:[.,:/-]\d+)+(?:[%％])?/giu;
  for (const match of text.matchAll(technical)) protectedRanges.push([match.index, match.index + match[0].length]);
  let offset = 0, rangeIndex = 0;
  for (const run of runs) {
    if (run.format.code || run.format.literal) { pushText(result, run.text, run.format); offset += run.text.length; continue; }
    for (const original of run.text) {
      while (rangeIndex < protectedRanges.length && protectedRanges[rangeIndex][1] <= offset) rangeIndex++;
      const protectedText = rangeIndex < protectedRanges.length && protectedRanges[rangeIndex][0] <= offset;
      const before = text[offset - 1] || "", after = text[offset + original.length] || "";
      let char = original;
      if (!protectedText) {
        if (settings.punctuationMode === "half") char = HALF_PUNCTUATION[char] || char;
        else {
          if (char === "＂") char = '"';
          if (char === "＇") char = "'";
          if (char === '"' || char === "'") {
            // Apostrophes inside English words are not quotation delimiters.
            const apostrophe = char === "'" && /[\p{L}\p{N}]/u.test(before) && /[\p{L}\p{N}]/u.test(after);
            if (!apostrophe) {
              const key = char === '"' ? "double" : "single";
              const opening = !after || /[\s，。！？；：、.,!?;:)\]}）】]/u.test(after) ? false :
                !before || /[\s([\{（【：:，,。.!?！？]/u.test(before) ? true : !quotes[key];
              quotes[key] = opening;
              char = key === "double" ? (opening ? "“" : "”") : (opening ? "‘" : "’");
            }
          } else {
            // Preserve decimal points and dot sequences, such as .5 and ... .
            const numeric = char === "." && (/\d/.test(after) || before === "." || after === ".");
            if (!numeric) char = FULL_PUNCTUATION[char] || char;
            if (char === "“") quotes.double = true;
            if (char === "”") quotes.double = false;
            if (char === "‘") quotes.single = true;
            if (char === "’") quotes.single = false;
          }
        }
      }
      const chinesePunctuation = /[“”‘’…—–·\u3000-\u303f\uff01-\uff0f\uff1a-\uff20\uff3b-\uff40\uff5b-\uff65]/u.test(char);
      pushText(result, char, chinesePunctuation ? { ...run.format, chinesePunctuation: true } : run.format);
      offset += original.length;
    }
  }
  return result;
}
function closing(text, token, from) {
  let pos = from;
  while ((pos = text.indexOf(token, pos)) !== -1) {
    let escapes = 0;
    for (let j = pos - 1; j >= 0 && text[j] === "\\"; j--) escapes++;
    if (escapes % 2 === 0) return pos;
    pos += token.length;
  }
  return -1;
}
function bracketEnd(text, start) {
  let depth = 1;
  for (let i = start + 1; i < text.length; i++) {
    if (text[i] === "\\") { i++; continue; }
    if (text[i] === "[") depth++;
    if (text[i] === "]" && --depth === 0) return i;
  }
  return -1;
}
function linkTargetEnd(text, start) {
  let depth = 1;
  for (let i = start + 1; i < text.length; i++) {
    if (text[i] === "\\") { i++; continue; }
    if (text[i] === "(") depth++;
    if (text[i] === ")" && --depth === 0) return i;
  }
  return -1;
}
function emphasisEnd(text, start, token) {
  const char = token[0], nested = [];
  for (let i = start + token.length; i < text.length; i++) {
    if (text[i] === "\\") { i++; continue; }
    if (text[i] === "`") {
      const ticks = /^`+/.exec(text.slice(i))[0], end = text.indexOf(ticks, i + ticks.length);
      if (end >= 0) { i = end + ticks.length - 1; continue; }
    }
    if (text[i] !== char) continue;
    let length = 1;
    while (text[i + length] === char) length++;
    const before = text[i - 1] || " ", after = text[i + length] || " ";
    const canClose = !/\s/.test(before) && !(char === "_" && /[\p{L}\p{N}]/u.test(after));
    const canOpen = !/\s/.test(after) && !(char === "_" && /[\p{L}\p{N}]/u.test(before));
    let used = 0;
    if (canClose) {
      while (nested.length && length - used >= nested[nested.length - 1]) used += nested.pop();
      if (!nested.length && length - used >= token.length) return i + used;
    }
    if (canOpen && length > used) nested.push(Math.min(3, length - used));
    i += length - 1;
  }
  return -1;
}
function stripComments(markdown) {
  let result = "", comment = null, fence = null;
  for (const line of markdown.split("\n")) {
    const marker = /^ {0,3}(?:>\s?)*(`{3,}|~{3,})(.*)$/.exec(line);
    if (!comment && marker) {
      if (!fence) fence = { char: marker[1][0], length: marker[1].length };
      else if (marker[1][0] === fence.char && marker[1].length >= fence.length && !marker[2].trim()) fence = null;
      result += line + "\n"; continue;
    }
    if (fence || (!comment && /^( {4}|\t)/.test(line))) { result += line + "\n"; continue; }
    for (let i = 0; i < line.length;) {
      if (comment) {
        const end = line.indexOf(comment, i);
        if (end < 0) break;
        i = end + comment.length; comment = null; continue;
      }
      if (line[i] === "\\" && i + 1 < line.length) { result += line.slice(i, i + 2); i += 2; continue; }
      if (line[i] === "`") {
        const ticks = /^`+/.exec(line.slice(i))[0], end = line.indexOf(ticks, i + ticks.length);
        if (end >= 0) { result += line.slice(i, end + ticks.length); i = end + ticks.length; continue; }
      }
      if (line.startsWith("%%", i)) { comment = "%%"; i += 2; continue; }
      if (line.startsWith("<!--", i)) { comment = "-->"; i += 4; continue; }
      result += line[i++];
    }
    result += "\n";
  }
  return result.slice(0, -1);
}
function inline(text, context, format = {}, depth = 0) {
  const runs = [];
  if (depth > 30) { pushText(runs, text, format); return runs; }
  let i = 0, plain = "";
  const flush = () => { pushText(runs, decodeEntities(plain), format); plain = ""; };
  while (i < text.length) {
    const rest = text.slice(i);
    if (text[i] === "\\" && i + 1 < text.length && /[\\`*{}\[\]()#+\-.!_>~|$%=]/.test(text[i + 1])) {
      plain += text[i + 1]; i += 2; continue;
    }
    if (text[i] === "`") {
      const ticks = /^`+/.exec(rest)[0];
      let end = text.indexOf(ticks, i + ticks.length);
      while (end >= 0 && (text[end - 1] === "`" || text[end + ticks.length] === "`")) end = text.indexOf(ticks, end + ticks.length);
      if (end >= 0) {
        flush(); let code = text.slice(i + ticks.length, end).replace(/\n/g, " ");
        if (/^ .* $/.test(code) && /\S/.test(code)) code = code.slice(1, -1);
        pushText(runs, code, { ...format, code: true }); i = end + ticks.length; continue;
      }
    }
    if (rest.startsWith("%%")) {
      const end = text.indexOf("%%", i + 2);
      if (end >= 0) { i = end + 2; continue; }
    }
    if (rest.startsWith("<!--")) {
      const end = text.indexOf("-->", i + 4);
      if (end >= 0) { i = end + 3; continue; }
    }
    const wikiOffset = rest.startsWith("![[") ? 3 : rest.startsWith("[[") ? 2 : 0;
    if (wikiOffset) {
      const end = closing(text, "]]", i + wikiOffset);
      if (end >= 0) {
        flush(); const target = text.slice(i + wikiOffset, end);
        const bar = target.indexOf("|");
        const label = bar >= 0 ? target.slice(bar + 1) : target;
        // A wiki target may contain underscores or stars as literal filename text.
        pushText(runs, label.replace(/\\([|\[\]])/g, "$1"), format);
        i = end + 2; continue;
      }
    }
    const foot = /^\[\^([^\]]+)\]/.exec(rest);
    if (foot && context.footnotes.has(foot[1])) {
      flush(); pushText(runs, "[" + context.footnotes.get(foot[1]).number + "]", format);
      i += foot[0].length; continue;
    }
    const image = rest.startsWith("![");
    if (image || text[i] === "[") {
      const start = i + (image ? 1 : 0), end = bracketEnd(text, start);
      if (end >= 0) {
        let after = end + 1, valid = false;
        if (text[after] === "(") {
          const targetEnd = linkTargetEnd(text, after);
          if (targetEnd >= 0) { valid = true; after = targetEnd + 1; }
        } else if (text[after] === "[") {
          const refEnd = closing(text, "]", after + 1);
          if (refEnd >= 0) {
            const ref = (text.slice(after + 1, refEnd) || text.slice(start + 1, end)).trim().toLowerCase();
            if (context.references.has(ref)) { valid = true; after = refEnd + 1; }
          }
        } else if (context.references.has(text.slice(start + 1, end).trim().toLowerCase())) valid = true;
        if (valid) {
          flush(); runs.push(...inline(text.slice(start + 1, end), context, format, depth + 1));
          i = after; continue;
        }
      }
    }
    const inlineFootnote = /^\^\[/.test(rest);
    if (inlineFootnote) {
      const end = bracketEnd(text, i + 1);
      if (end >= 0) {
        flush(); pushText(runs, " (", format);
        runs.push(...inline(text.slice(i + 2, end), context, format, depth + 1));
        pushText(runs, ")", format); i = end + 1; continue;
      }
    }
    if (text[i] === "<") {
      const auto = /^<((?:https?:\/\/|mailto:)[^<>\s]+|[^<>\s@]+@[^<>\s@]+)>/.exec(rest);
      if (auto) { plain += auto[1].replace(/^mailto:/, ""); i += auto[0].length; continue; }
      const html = /^<\/?([a-z][a-z0-9-]*)\b[^>]*>/i.exec(rest);
      if (html) {
        context.warnings.add("HTML 已按可见文本简化转换。");
        const tag = html[1].toLowerCase();
        if ((tag === "script" || tag === "style") && !rest.startsWith("</")) {
          const re = new RegExp("</" + tag + "\\s*>", "i"), match = re.exec(text.slice(i + html[0].length));
          i = match ? i + html[0].length + match.index + match[0].length : text.length; continue;
        }
        if (tag === "br" || /^(p|div|li|tr|h[1-6])$/.test(tag)) plain += "\n";
        if (tag === "img") {
          const alt = /\balt\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))/i.exec(html[0]);
          if (alt) plain += alt[1] ?? alt[2] ?? alt[3] ?? "";
        }
        i += html[0].length; continue;
      }
    }
    let matched = false;
    for (const [token, properties] of [["***", { bold: true, italic: true }], ["___", { bold: true, italic: true }], ["**", { bold: true }], ["__", { bold: true }], ["~~", { strike: true }], ["==", { highlight: true }], ["*", { italic: true }], ["_", { italic: true }]]) {
      if (!rest.startsWith(token) || /\s/.test(text[i + token.length] || " ")) continue;
      if (token[0] === "_" && /[\p{L}\p{N}]/u.test(text[i - 1] || "") && /[\p{L}\p{N}]/u.test(text[i + token.length] || "")) continue;
      let end = token[0] === "*" || token[0] === "_" ? emphasisEnd(text, i, token) : closing(text, token, i + token.length);
      while (end >= 0 && (/\s/.test(text[end - 1]) || (token[0] === "_" && /[\p{L}\p{N}]/u.test(text[end + token.length] || "")))) end = closing(text, token, end + token.length);
      if (end < 0) continue;
      flush(); runs.push(...inline(text.slice(i + token.length, end), context, { ...format, ...properties }, depth + 1));
      i = end + token.length; matched = true; break;
    }
    if (matched) continue;
    if (text[i] === "$") {
      const token = rest.startsWith("$$") ? "$$" : "$", end = closing(text, token, i + token.length);
      if (end >= 0 && end > i + token.length && !(token === "$" && (/\s/.test(text[i + 1]) || /\s/.test(text[end - 1]) || /\d/.test(text[end + 1] || "")))) {
        flush(); pushText(runs, text.slice(i + token.length, end), { ...format, literal: true }); i = end + token.length; continue;
      }
    }
    if (text[i] === "\n") {
      plain = plain.replace(/(?: {2,}|\\)$/, "");
    }
    plain += text[i++];
  }
  flush(); return runs;
}

function parseMarkdown(markdown, settings) {
  const context = { footnotes: new Map(), references: new Set(), warnings: new Set() };
  let lines = String(markdown).replace(/^\ufeff/, "").replace(/\r\n?/g, "\n").split("\n");
  if (lines[lines.length - 1] === "") lines.pop();
  let yaml = [];
  if (lines[0] === "---") {
    const end = lines.findIndex((line, i) => i > 0 && /^(---|\.\.\.)\s*$/.test(line));
    if (end >= 0) { if (settings.includeYaml) yaml = lines.slice(1, end); lines = lines.slice(end + 1); }
  }
  lines = stripComments(lines.join("\n")).split("\n");
  // Definitions are collected outside fenced/indented code. Keep all footnote text,
  // including unused definitions, so none of the author's notes silently vanish.
  const body = [];
  let fence = null;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i], fm = /^ {0,3}(`{3,}|~{3,})/.exec(line);
    if (fm) {
      if (!fence) fence = { char: fm[1][0], length: fm[1].length };
      else if (fm[1][0] === fence.char && fm[1].length >= fence.length && /^ {0,3}(?:`+|~+)\s*$/.test(line)) fence = null;
      body.push(line); continue;
    }
    if (fence || /^( {4}|\t)/.test(line)) { body.push(line); continue; }
    const foot = /^ {0,3}\[\^([^\]]+)\]:[ \t]*(.*)$/.exec(line);
    if (foot) {
      const parts = [foot[2]];
      while (i + 1 < lines.length && (/^( {4}|\t)/.test(lines[i + 1]) || (lines[i + 1] === "" && /^( {4}|\t)/.test(lines[i + 2] || "")))) {
        parts.push(lines[++i].replace(/^( {4}|\t)/, ""));
      }
      if (context.footnotes.has(foot[1])) {
        context.footnotes.get(foot[1]).text += "\n" + parts.join("\n");
      } else context.footnotes.set(foot[1], { number: context.footnotes.size + 1, text: parts.join("\n") });
      continue;
    }
    const ref = /^ {0,3}\[([^\]]+)\]:[ \t]*(?:<[^<>\n]*>|\S+?)(?:[ \t]+(?:"[^"]*"|'[^']*'|\([^)]*\)))?[ \t]*$/.exec(line);
    if (ref) { context.references.add(ref[1].trim().toLowerCase()); continue; }
    body.push(line);
  }
  const blocks = [];
  if (yaml.length) blocks.push({ style: "Code", text: yaml.join("\n"), literal: true });
  let para = [], paraMeta = null, code = null, listStack = [], nextListId = 1, tableActive = false;
  const flush = () => {
    if (para.length) blocks.push({ ...paraMeta, text: para.join("\n") });
    para = []; paraMeta = null;
  };
  for (let i = 0; i < body.length; i++) {
    const raw = body[i];
    if (code) {
      const codeLine = code.quoted ? raw.replace(/^ {0,3}(?:>\s?)+/, "") : raw;
      const close = /^ {0,3}(`+|~+)\s*$/.exec(codeLine);
      if (close && close[1][0] === code.char && close[1].length >= code.length) {
        blocks.push({ style: "Code", text: code.lines.join("\n"), literal: true }); code = null;
      } else code.lines.push(codeLine);
      continue;
    }
    const quote = /^ {0,3}(?:>\s?)+/.exec(raw);
    let line = quote ? raw.slice(quote[0].length) : raw;
    const quoted = !!quote;
    const fm = /^ {0,3}(`{3,}|~{3,})(.*)$/.exec(line);
    if (fm) { flush(); code = { char: fm[1][0], length: fm[1].length, lines: [], quoted }; listStack = []; tableActive = false; continue; }
    if (line.trim() === "$$") {
      flush(); const math = [];
      let end = i + 1;
      while (end < body.length && body[end].trim() !== "$$") math.push(body[end++]);
      if (end < body.length) { blocks.push({ style: "Normal", literal: true, text: math.join("\n") }); i = end; continue; }
      // An unmatched delimiter is retained as text.
    }
    if (!line.trim()) {
      flush(); tableActive = false;
      if (settings.keepBlankLines) blocks.push({ text: "", style: "Normal" });
      continue;
    }
    const list = /^(\s*)([-+*]|\d{1,9}[.)])\s+(.*)$/.exec(line);
    if (!list && /^( {4}|\t)/.test(line) && !listStack.length) {
      flush(); const parts = [line.replace(/^( {4}|\t)/, "")];
      while (i + 1 < body.length && /^( {4}|\t)/.test(body[i + 1])) parts.push(body[++i].replace(/^( {4}|\t)/, ""));
      blocks.push({ style: "Code", text: parts.join("\n"), literal: true }); continue;
    }
    if (/^ {0,3}(?:(?:\*\s*){3,}|(?:-\s*){3,}|(?:_\s*){3,})$/.test(line)) {
      flush(); blocks.push({ rule: true, text: "" }); listStack = []; continue;
    }
    const heading = /^ {0,3}(#{1,6})(?:[ \t]+(.*)|[ \t]*)$/.exec(line);
    if (heading) {
      flush(); blocks.push({ style: "Heading" + heading[1].length, text: (heading[2] || "").replace(/[ \t]+#+[ \t]*$/, "") }); listStack = []; tableActive = false; continue;
    }
    if (i + 1 < body.length && /^ {0,3}(?:=+|-+)\s*$/.test(body[i + 1]) && !list && !quoted) {
      flush(); blocks.push({ style: /^\s*=/.test(body[i + 1]) ? "Heading1" : "Heading2", text: line }); i++; listStack = []; continue;
    }
    if (list) {
      flush(); const indent = list[1].replace(/\t/g, "    ").length, ordered = /^\d/.test(list[2]);
      while (listStack.length && listStack[listStack.length - 1].indent > indent) listStack.pop();
      let current = listStack[listStack.length - 1];
      if (!current || current.indent < indent || current.ordered !== ordered) {
        if (current && current.indent === indent) listStack.pop();
        current = { indent, ordered, id: nextListId++, start: ordered ? parseInt(list[2], 10) : 1 };
        listStack.push(current);
      }
      let content = list[3];
      content = content.replace(/^\[([ xX])\]\s*/, (_, c) => c === " " ? "☐ " : "☑ ");
      paraMeta = { style: quoted ? "Quote" : "Normal", list: { ...current, level: Math.min(8, listStack.length - 1) } };
      para.push(content); continue;
    }
    const callout = /^\[!([^\]]+)\][+-]?\s*(.*)$/.exec(line);
    if (quoted && callout) {
      flush(); blocks.push({ style: "Quote", text: callout[2] || callout[1] }); continue;
    }
    // Pipe tables: retain row/cell order, skip only the formatting separator row.
    if (/^\s*\|?\s*:?-{3,}:?\s*\|(?:\s*:?-{3,}:?\s*\|?)+\s*$/.test(line)) continue;
    if (line.includes("|") && (tableActive || line.trim().startsWith("|") || (i + 1 < body.length && /^\s*\|?\s*:?-{3,}:?\s*\|/.test(body[i + 1])))) {
      tableActive = true;
      flush(); context.warnings.add("表格已按行转换为正文，单元格以制表符分隔。");
      const cells = line.trim().replace(/^\|/, "").replace(/(?<!\\)\|$/, "").split(/(?<!\\)\|/).map(c => c.trim());
      blocks.push({ style: quoted ? "Quote" : "Normal", text: cells.join("\t") }); continue;
    }
    const isContinuation = listStack.length && /^\s+/.test(line);
    if (paraMeta && (paraMeta.style !== (quoted ? "Quote" : "Normal") || (paraMeta.list && !isContinuation))) flush();
    if (!isContinuation) listStack = [];
    if (!paraMeta) paraMeta = { style: quoted ? "Quote" : "Normal" };
    para.push(isContinuation ? line.trimStart() : line);
  }
  flush();
  if (code) blocks.push({ style: "Code", text: code.lines.join("\n"), literal: true });
  if (context.footnotes.size) {
    blocks.push({ style: "Heading2", text: "脚注" });
    for (const foot of context.footnotes.values()) blocks.push({ style: "Normal", text: "[" + foot.number + "] " + foot.text });
  }
  const quotes = { double: false, single: false };
  for (const block of blocks) block.runs = block.literal ? [{ text: block.text, format: { code: block.style === "Code" } }] : typographyRuns(inline(block.text, context), settings, quotes);
  return { blocks, warnings: [...context.warnings] };
}

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
function runXml(run, settings) {
  const f = run.format;
  // All four font slots keep Chinese quotation marks in the chosen font.
  // Do not force an eastAsia hint: it changes Word's counting of ellipses.
  const fonts = f.chinesePunctuation ? '<w:rFonts w:ascii="' + xml(settings.chineseFont) + '" w:hAnsi="' + xml(settings.chineseFont) + '" w:eastAsia="' + xml(settings.chineseFont) + '" w:cs="' + xml(settings.chineseFont) + '"/>' : f.code ? '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>' : '';
  const props = [fonts, f.bold ? '<w:b/>' : '', f.italic ? '<w:i/>' : '', f.strike ? '<w:strike/>' : '', f.highlight ? '<w:highlight w:val="yellow"/>' : '', f.chinesePunctuation ? '<w:lang w:val="zh-CN" w:eastAsia="zh-CN"/>' : ''].join('');
  const text = run.text.split(/(\n|\t)/).map(t => t === "\n" ? '<w:br/>' : t === "\t" ? '<w:tab/>' : '<w:t xml:space="preserve">' + xml(t) + '</w:t>').join('');
  return '<w:r>' + (props ? '<w:rPr>' + props + '</w:rPr>' : '') + text + '</w:r>';
}
function stylesXml(s) {
  const size = Math.round(s.fontSize * 2), indent = Math.round(s.firstLineIndent * 100);
  let styles = XML_HEADER + '<w:styles xmlns:w="' + W + '"><w:docDefaults><w:rPrDefault><w:rPr>' +
    '<w:rFonts w:ascii="' + xml(s.englishFont) + '" w:hAnsi="' + xml(s.englishFont) + '" w:eastAsia="' + xml(s.chineseFont) + '" w:cs="' + xml(s.englishFont) + '"/>' +
    '<w:sz w:val="' + size + '"/><w:szCs w:val="' + size + '"/><w:lang w:val="en-US" w:eastAsia="zh-CN"/>' +
    '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="' + Math.round(s.lineSpacing * 240) + '" w:lineRule="auto"/>' +
    '<w:ind w:firstLineChars="' + indent + '"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>';
  for (let level = 1; level <= 6; level++) {
    styles += '<w:style w:type="paragraph" w:styleId="Heading' + level + '"><w:name w:val="heading ' + level + '"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>' +
      '<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:ind w:firstLine="0" w:firstLineChars="0"/><w:outlineLvl w:val="' + (level - 1) + '"/></w:pPr><w:rPr><w:b/><w:sz w:val="' + Math.max(size, 36 - level * 2) + '"/></w:rPr></w:style>';
  }
  styles += '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:pPr><w:pBdr><w:left w:val="single" w:sz="12" w:space="8" w:color="AAAAAA"/></w:pBdr><w:ind w:left="420" w:firstLine="0" w:firstLineChars="0"/></w:pPr></w:style>';
  styles += '<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="80" w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0" w:firstLineChars="0"/><w:shd w:fill="F3F3F3"/></w:pPr><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="' + size + '"/></w:rPr></w:style></w:styles>';
  return styles;
}
function numberingXml(lists) {
  let result = XML_HEADER + '<w:numbering xmlns:w="' + W + '">';
  for (let type = 0; type < 2; type++) {
    result += '<w:abstractNum w:abstractNumId="' + type + '"><w:multiLevelType w:val="multilevel"/>';
    for (let level = 0; level < 9; level++) {
      result += '<w:lvl w:ilvl="' + level + '"><w:start w:val="1"/><w:numFmt w:val="' + (type ? 'decimal' : 'bullet') + '"/><w:lvlText w:val="' + (type ? '%' + (level + 1) + '.' : '•') + '"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="' + ((level + 1) * 420) + '"/></w:tabs><w:ind w:left="' + ((level + 1) * 420) + '" w:hanging="210"/></w:pPr></w:lvl>';
    }
    result += '</w:abstractNum>';
  }
  for (const list of lists.values()) {
    result += '<w:num w:numId="' + list.id + '"><w:abstractNumId w:val="' + (list.ordered ? 1 : 0) + '"/>';
    if (list.ordered) result += '<w:lvlOverride w:ilvl="' + list.level + '"><w:startOverride w:val="' + list.start + '"/></w:lvlOverride>';
    result += '</w:num>';
  }
  return result + '</w:numbering>';
}
let crcTable;
function crc32(buffer) {
  if (!crcTable) crcTable = Array.from({ length: 256 }, (_, n) => {
    let c = n;
    for (let j = 0; j < 8; j++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    return c >>> 0;
  });
  let crc = 0xffffffff;
  for (const byte of buffer) crc = crcTable[(crc ^ byte) & 255] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}
async function zip(entries) {
  const local = [], central = []; let offset = 0;
  for (const [name, text] of entries) {
    const filename = Buffer.from(name, "utf8"), data = Buffer.from(text, "utf8"), crc = crc32(data);
    if (data.length > 0xffffffff) throw new Error("文档超出 ZIP32 大小限制。");
    // ZIP method 8 requires raw DEFLATE, without the zlib wrapper. Compress on
    // Node's worker pool so compression does not block the Obsidian UI thread.
    const compressed = await deflateZip(data, { level: 6 });
    if (compressed.length > 0xffffffff || offset + compressed.length + filename.length + 30 > 0xffffffff) throw new Error("文档超出 ZIP32 大小限制。");
    const head = Buffer.alloc(30);
    head.writeUInt32LE(0x04034b50, 0); head.writeUInt16LE(20, 4); head.writeUInt16LE(0x800, 6); head.writeUInt16LE(8, 8);
    head.writeUInt16LE(0x21, 12); head.writeUInt32LE(crc, 14); head.writeUInt32LE(compressed.length, 18); head.writeUInt32LE(data.length, 22); head.writeUInt16LE(filename.length, 26);
    local.push(head, filename, compressed);
    const dir = Buffer.alloc(46);
    dir.writeUInt32LE(0x02014b50, 0); dir.writeUInt16LE(20, 4); dir.writeUInt16LE(20, 6); dir.writeUInt16LE(0x800, 8); dir.writeUInt16LE(8, 10);
    dir.writeUInt16LE(0x21, 14); dir.writeUInt32LE(crc, 16); dir.writeUInt32LE(compressed.length, 20); dir.writeUInt32LE(data.length, 24); dir.writeUInt16LE(filename.length, 28); dir.writeUInt32LE(offset, 42);
    central.push(dir, filename); offset += head.length + filename.length + compressed.length;
  }
  const centralSize = central.reduce((n, b) => n + b.length, 0), end = Buffer.alloc(22);
  if (offset + centralSize > 0xffffffff) throw new Error("文档超出 ZIP32 大小限制。");
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(entries.length, 8); end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(centralSize, 12); end.writeUInt32LE(offset, 16);
  return Buffer.concat([...local, ...central, end]);
}
async function markdownToDocx(content, settings) {
  const parsed = parseMarkdown(content, settings), lists = new Map();
  let document = XML_HEADER + '<w:document xmlns:w="' + W + '"><w:body>';
  for (const block of parsed.blocks) {
    let ppr = '<w:pStyle w:val="' + (block.style || 'Normal') + '"/>';
    if (block.list) {
      lists.set(block.list.id, block.list);
      ppr += '<w:numPr><w:ilvl w:val="' + block.list.level + '"/><w:numId w:val="' + block.list.id + '"/></w:numPr><w:ind w:firstLine="0" w:firstLineChars="0"/>';
    }
    if (block.rule) ppr += '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="999999"/></w:pBdr>';
    document += '<w:p><w:pPr>' + ppr + '</w:pPr>' + block.runs.map(run => runXml(run, settings)).join('') + '</w:p>';
  }
  if (!parsed.blocks.length) document += '<w:p/>';
  document += '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>';
  const entries = [
    ['[Content_Types].xml', XML_HEADER + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'],
    ['_rels/.rels', XML_HEADER + '<Relationships xmlns="' + REL + '"><Relationship Id="rId1" Type="' + OFFICE_REL + 'officeDocument" Target="word/document.xml"/></Relationships>'],
    ['word/document.xml', document],
    ['word/_rels/document.xml.rels', XML_HEADER + '<Relationships xmlns="' + REL + '"><Relationship Id="rIdStyles" Type="' + OFFICE_REL + 'styles" Target="styles.xml"/><Relationship Id="rIdNumbering" Type="' + OFFICE_REL + 'numbering" Target="numbering.xml"/></Relationships>'],
    ['word/styles.xml', stylesXml(settings)], ['word/numbering.xml', numberingXml(lists)]
  ];
  return { buffer: await zip(entries), warnings: parsed.warnings };
}

// Only fixed program text is passed to -EncodedCommand. Paths and process
// identities travel via the child environment, never through command interpolation.
const WORD_SCRIPT = String.raw`
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
function Send($value) { [Console]::WriteLine(($value | ConvertTo-Json -Compress -Depth 4)); [Console]::Out.Flush() }
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class WceWindow {
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
$word = $null; $documents = $null; $doc = $null; $probe = $null; $window = $null; $owned = $false; $outcome = $null
$dontSave = [object]0
$existing = @{}
Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $existing[$_.Id] = $true }
try {
  $word = New-Object -ComObject Word.Application
  $documents = $word.Documents
  $missing = [Type]::Missing
  # Word exposes Hwnd on Window, not Application. A private invisible blank
  # document supplies a window before opening the actual file (which may hang).
  $probe = $documents.Add($missing, $false, $missing, $false)
  $window = $probe.ActiveWindow
  [uint32]$wordPid = 0
  [void][WceWindow]::GetWindowThreadProcessId([IntPtr]$window.Hwnd, [ref]$wordPid)
  if ($wordPid -eq 0 -or $existing.ContainsKey([int]$wordPid)) {
    throw 'WCE_SHARED: Word returned an existing or unidentified process. No user instance was changed.'
  }
  $process = [System.Diagnostics.Process]::GetProcessById([int]$wordPid)
  $owned = $true
  Send @{ type = 'owner'; pid = [int]$wordPid; ticks = $process.StartTime.ToUniversalTime().Ticks.ToString() }
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $word.AutomationSecurity = 3
  [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($window)
  $window = $null
  $probe.Close([ref]$dontSave)
  [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($probe)
  $probe = $null
  # Explicitly disable document visibility, repair prompts and encoding dialogs.
  # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, passwords/template,
  # Revert, write passwords, Format, Encoding, Visible, OpenAndRepair,
  # DocumentDirection, NoEncodingDialog.
  # IDispatch reflection avoids PowerShell/Office PIA optional-ref binding bugs.
  $openArgs = [object[]]@($env:WCE_DOCX_PATH, $false, $true, $false, $missing, $missing, $false, $missing, $missing, $missing, $missing, $false, $false, $missing, $true, $missing)
  $doc = $documents.GetType().InvokeMember('Open', [Reflection.BindingFlags]::InvokeMethod, $null, $documents, $openArgs)
  $statistics = @{ words = [int]$doc.ComputeStatistics(0, $false) }
  $statErrors = @{}
  $layoutError = $null
  try { $doc.Repaginate() } catch { $layoutError = $_.Exception.Message }
  $metrics = @(
    @{ name = 'characters'; id = 3 }, @{ name = 'charactersWithSpaces'; id = 5 },
    @{ name = 'paragraphs'; id = 4 }, @{ name = 'lines'; id = 1 },
    @{ name = 'pages'; id = 2 }, @{ name = 'farEastCharacters'; id = 6 }
  )
  foreach ($metric in $metrics) {
    try {
      if ($layoutError -and ($metric.name -eq 'lines' -or $metric.name -eq 'pages')) { throw $layoutError }
      $statistics[$metric.name] = [int]$doc.ComputeStatistics([int]$metric.id, $false)
    } catch {
      $statistics[$metric.name] = $null
      $statErrors[$metric.name] = $_.Exception.Message
    }
  }
  $outcome = @{ type = 'result'; statistics = $statistics; statisticsErrors = $statErrors }
} catch {
  $hr = $_.Exception.HResult
  $missing = ($hr -eq -2147221164) -or ($_.Exception.ToString() -match '80040154|REGDB_E_CLASSNOTREG')
  $outcome = @{ type = 'error'; missing = [bool]$missing; message = $_.Exception.Message; detail = $_.Exception.ToString() }
} finally {
  if ($null -ne $window) {
    try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($window) } catch {}
    $window = $null
  }
  if ($null -ne $probe) {
    try { $probe.Close([ref]$dontSave) } catch {}
    try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($probe) } catch {}
    $probe = $null
  }
  if ($null -ne $doc) {
    try { $doc.Close([ref]$dontSave) } catch {}
    try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc) } catch {}
    $doc = $null
  }
  if ($null -ne $documents) {
    try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($documents) } catch {}
    $documents = $null
  }
  if ($null -ne $word) {
    if ($owned) { try { $word.Quit([ref]$dontSave) } catch {} }
    try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) } catch {}
    $word = $null
  }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
if ($null -ne $outcome) { Send $outcome }
`;
const CLEANUP_SCRIPT = String.raw`
$ErrorActionPreference = 'Stop'
try {
  $p = [System.Diagnostics.Process]::GetProcessById([int]$env:WCE_OWNED_PID)
  if ($p.ProcessName -eq 'WINWORD' -and $p.StartTime.ToUniversalTime().Ticks.ToString() -eq $env:WCE_OWNED_TICKS) {
    $p.Kill()
    [void]$p.WaitForExit(4000)
  }
} catch [System.ArgumentException] { } catch { [Console]::Error.WriteLine($_.Exception.Message); exit 1 }
`;
function powershellPath() {
  return path.join(process.env.SystemRoot || "C:\\Windows", "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
}
function spawnScript(script, extraEnv) {
  return spawn(powershellPath(), ["-NoLogo", "-NoProfile", "-NonInteractive", "-Sta", "-ExecutionPolicy", "Bypass", "-EncodedCommand", Buffer.from(script, "utf16le").toString("base64")], {
    windowsHide: true, stdio: ["ignore", "pipe", "pipe"], env: { ...process.env, ...extraEnv }
  });
}
function cleanOwnedWord(owner) {
  if (!owner) return Promise.resolve();
  return new Promise((resolve, reject) => {
    let child;
    try { child = spawnScript(CLEANUP_SCRIPT, { WCE_OWNED_PID: String(owner.pid), WCE_OWNED_TICKS: owner.ticks }); }
    catch (error) { reject(error); return; }
    let errors = "", settled = false;
    const finish = error => { if (settled) return; settled = true; clearTimeout(timer); error ? reject(error) : resolve(); };
    const timer = setTimeout(() => { child.kill(); finish(new Error("清理 Word 辅助进程超时。")); }, 7000);
    child.stdout.resume(); child.stderr.on("data", data => { if (errors.length < 4096) errors += data.toString("utf8"); });
    child.on("error", finish);
    child.on("close", code => finish(code === 0 ? null : new Error(errors || "无法清理插件创建的 Word 进程。")));
  });
}
function wordStatistics(docxPath, timeoutSeconds, signal) {
  if (process.platform !== "win32") return Promise.reject(new Error("Word 字数统计仅支持 Windows 桌面版。"));
  return new Promise((resolve, reject) => {
    let child;
    try { child = spawnScript(WORD_SCRIPT, { WCE_DOCX_PATH: docxPath }); }
    catch (error) { reject(error); return; }
    let owner = null, result = null, pending = "", stderr = "", failure = null, finished = false, closingChild = false;
    const timeoutError = new Error("Word 统计超时（" + timeoutSeconds + " 秒）。请检查 Word 的首次启动、授权或恢复对话框。");
    const consume = line => {
      if (!line.trim()) return;
      try {
        const value = JSON.parse(line.replace(/^\ufeff/, ""));
        if (value.type === "owner" && Number.isSafeInteger(value.pid) && value.pid > 0 && /^\d{15,20}$/.test(value.ticks)) owner = value;
        if (value.type === "result" && validStatistics(value.statistics)) result = value;
        if (value.type === "error") {
          failure = new Error(value.message || "Word 统计失败。"); failure.missingWord = value.missing === true; failure.detail = value.detail;
        }
      } catch (_) { /* Ignore PowerShell informational output; never evaluate it. */ }
    };
    const stop = error => {
      if (finished || closingChild) return;
      failure = error; closingChild = true;
      child.kill();
    };
    const timer = setTimeout(() => stop(timeoutError), timeoutSeconds * 1000);
    const abort = () => { const error = new Error("统计已取消。"); error.cancelled = true; stop(error); };
    if (signal) { signal.addEventListener("abort", abort, { once: true }); if (signal.aborted) abort(); }
    child.stdout.setEncoding("utf8"); child.stderr.setEncoding("utf8");
    child.stdout.on("data", data => {
      pending += data;
      let end;
      while ((end = pending.indexOf("\n")) >= 0) { consume(pending.slice(0, end)); pending = pending.slice(end + 1); }
      if (pending.length > 65536) pending = pending.slice(-8192);
    });
    child.stderr.on("data", data => { if (stderr.length < 16384) stderr += data; });
    child.on("error", error => { failure = error; });
    child.on("close", async code => {
      if (finished) return;
      finished = true; clearTimeout(timer); if (signal) signal.removeEventListener("abort", abort);
      consume(pending);
      // Also verify cleanup after normal completion: Quit can fail silently in COM.
      try { await cleanOwnedWord(owner); }
      catch (error) { failure = new Error((failure ? failure.message + "\n" : "") + error.message); }
      if (failure) reject(failure);
      else if (code !== 0 || !result) reject(new Error(stderr.trim() || "Word 辅助进程未返回有效字数（退出码 " + code + "）。"));
      else resolve({ statistics: result.statistics, statisticsErrors: statisticsErrors(result.statisticsErrors) });
    });
  });
}

class PathModal extends Modal {
  constructor(app, initialPath, track) { super(app); this.initialPath = initialPath; this.track = track; }
  openResult() { return new Promise(resolve => { this.resolve = resolve; this.track.add(this); this.open(); }); }
  onOpen() {
    this.titleEl.setText("导出为 Word (.docx)");
    this.contentEl.createEl("p", { text: "选择导出位置：输入完整 Windows 本地路径。原 Markdown 笔记保持原样。" });
    this.input = this.contentEl.createEl("input", { type: "text", cls: "count-by-msword-obsidian-path" });
    this.input.value = this.initialPath;
    this.errorEl = this.contentEl.createDiv({ cls: "count-by-msword-obsidian-error" });
    const actions = this.contentEl.createDiv({ cls: "count-by-msword-obsidian-actions" });
    actions.createEl("button", { text: "取消" }).addEventListener("click", () => this.close());
    const save = actions.createEl("button", { text: "导出", cls: "mod-cta" });
    const submit = () => {
      let value = this.input.value.trim();
      if (value.startsWith('"') && value.endsWith('"')) value = value.slice(1, -1);
      if (!/^[a-z]:[\\/]/i.test(value)) { this.errorEl.setText("请输入带盘符的 Windows 本地绝对路径，例如 D:\\文稿\\文章.docx。"); return; }
      if (path.extname(value).toLowerCase() !== ".docx") { this.errorEl.setText("文件扩展名必须为 .docx。"); return; }
      const segments = value.slice(3).split(/[\\/]/);
      if (segments.some(part => !part || /[<>:"|?*\x00-\x1f]/.test(part) || /[. ]$/.test(part) || /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(part))) {
        this.errorEl.setText("路径包含 Windows 不允许的文件名或字符。"); return;
      }
      this.result = path.normalize(value); this.close();
    };
    save.addEventListener("click", submit);
    this.input.addEventListener("keydown", event => { if (event.key === "Enter") { event.preventDefault(); submit(); } });
    this.input.focus(); this.input.select();
  }
  onClose() { this.track.delete(this); this.contentEl.empty(); if (this.resolve) { this.resolve(this.result || null); this.resolve = null; } }
}
class ConfirmModal extends Modal {
  constructor(app, target, track) { super(app); this.target = target; this.track = track; }
  openResult() { return new Promise(resolve => { this.resolve = resolve; this.track.add(this); this.open(); }); }
  onOpen() {
    this.titleEl.setText("覆盖现有 Word 文档？");
    this.contentEl.createEl("p", { text: this.target });
    this.contentEl.createEl("p", { text: "已有文件将被当前笔记的 DOCX 替换。" });
    const actions = this.contentEl.createDiv({ cls: "count-by-msword-obsidian-actions" });
    actions.createEl("button", { text: "取消" }).addEventListener("click", () => this.close());
    actions.createEl("button", { text: "覆盖", cls: "mod-warning" }).addEventListener("click", () => { this.confirmed = true; this.close(); });
  }
  onClose() { this.track.delete(this); this.contentEl.empty(); if (this.resolve) { this.resolve(!!this.confirmed); this.resolve = null; } }
}
function safeFilename(name) {
  let safe = name.replace(/[<>:"/\\|?*\x00-\x1f]/g, "_").replace(/[. ]+$/, "");
  if (/^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(safe)) safe = "_" + safe;
  return (safe || "未命名") + ".docx";
}
function sameFileStat(a, b) { return a.dev === b.dev && a.ino === b.ino && a.size === b.size && a.mtimeMs === b.mtimeMs && a.ctimeMs === b.ctimeMs; }
async function removeTempDirectory(directory) {
  // directory is always the exact result of mkdtemp, never an export/user path.
  for (let attempt = 0; attempt < 5; attempt++) {
    try { await fsp.rm(directory, { recursive: true, force: true }); return; }
    catch (error) { if (attempt === 4) throw error; await new Promise(resolve => setTimeout(resolve, 150 * (attempt + 1))); }
  }
}

class CountByMSWordObsidian extends Plugin {
  async onload() {
    this.disposed = false; this.cache = new Map(); this.errors = new Map(); this.seen = new Set();
    this.edits = new WeakMap(); this.editSequence = 0; this.modals = new Set();
    this.pending = null; this.running = null; this.pumpPromise = null; this.activationTimer = null;
    this.activeIdentity = null; this.saveTail = Promise.resolve(); this.cacheEpoch = 0;
    let data;
    try { data = await this.loadData(); } catch (error) { console.error("count-by-msword-obsidian: cannot load settings", error); }
    if (this.disposed) return;
    this.settings = validatedSettings(data && data.settings);
    const key = conversionKey(this.settings);
    if (data && Array.isArray(data.cache)) {
      for (const entry of data.cache.slice(-100)) {
        if (entry && typeof entry.filePath === "string" && /^[a-f0-9]{64}$/.test(entry.contentHash) && entry.conversionKey === key && entry.statisticsVersion === STATISTICS_VERSION && validStatistics(entry.statistics) && entry.wordCount === entry.statistics.words && Number.isFinite(entry.calculatedAt)) {
          this.cache.set(cacheId(entry), { filePath: entry.filePath, contentHash: entry.contentHash, conversionKey: entry.conversionKey, wordCount: entry.wordCount, statisticsVersion: STATISTICS_VERSION, statistics: entry.statistics, statisticsErrors: statisticsErrors(entry.statisticsErrors), calculatedAt: entry.calculatedAt });
        }
      }
    }
    this.status = this.addStatusBarItem();
    this.status.addClass("count-by-msword-obsidian-status"); this.status.setAttribute("role", "button"); this.status.setAttribute("tabindex", "0");
    this.status.setAttribute("aria-description", "点击刷新，悬停查看统计并导出 DOCX");
    this.registerDomEvent(this.status, "click", () => this.refresh());
    this.registerDomEvent(this.status, "keydown", event => { if (event.key === "Enter" || event.key === " ") { event.preventDefault(); this.refresh(); } });
    this.createStatisticsHover();
    this.addCommand({ id: "refresh-word-word-count", name: "刷新 Word 字数 (Refresh Word word count)", callback: () => this.refresh() });
    this.addCommand({ id: "export-current-note-to-word", name: "导出当前笔记为 Word (Export current note to Word)", callback: () => { void this.exportNote(); } });
    this.addCommand({ id: "clear-word-count-cache", name: "清除 Word 字数缓存 (Clear Word count cache)", callback: () => this.clearCache() });
    this.registerEvent(this.app.workspace.on("file-open", () => this.scheduleActivation()));
    this.registerEvent(this.app.workspace.on("active-leaf-change", () => this.scheduleActivation()));
    this.registerEvent(this.app.workspace.on("editor-change", (_editor, info) => {
      if (info) this.edits.set(info, ++this.editSequence);
      this.renderStatus(); // Editing never starts Word.
    }));
    this.registerEvent(this.app.vault.on("modify", file => { if (this.currentView()?.file?.path === file.path) this.renderStatus(); }));
    this.registerEvent(this.app.vault.on("rename", (file, oldPath) => {
      this.removePathCache(oldPath); this.errors.delete(oldPath); this.seen.delete(oldPath);
      this.persist(); this.scheduleActivation();
    }));
    this.registerEvent(this.app.vault.on("delete", file => {
      this.removePathCache(file.path); this.errors.delete(file.path); this.seen.delete(file.path); this.persist(); this.renderStatus();
    }));
    this.registerEvent(this.app.workspace.on("file-menu", (menu, file) => {
      if (file && file.extension === "md") menu.addItem(item => item.setTitle("导出为 Word (.docx)").setIcon("file-output").onClick(() => { void this.exportNote(file); }));
    }));
    this.addSettingTab(new CountByMSWordSettings(this.app, this));
    this.app.workspace.onLayoutReady(() => { if (!this.disposed) this.scheduleActivation(); });
    this.renderStatus();
  }
  currentView() { return this.app.workspace.getActiveViewOfType(MarkdownView); }
  snapshotFromView(view) {
    if (!view || !view.file || view.file.extension !== "md") return null;
    const content = view.editor ? view.editor.getValue() : view.getViewData();
    return this.makeSnapshot(view.file.path, content);
  }
  makeSnapshot(filePath, content) { return { filePath, content, contentHash: digest(content), conversionKey: conversionKey(this.settings) }; }
  currentSnapshot() { try { return this.snapshotFromView(this.currentView()); } catch (error) { console.error("count-by-msword-obsidian: cannot read editor", error); return null; } }
  scheduleActivation() {
    clearTimeout(this.activationTimer);
    this.activationTimer = setTimeout(() => {
      this.activationTimer = null;
      if (this.disposed) return;
      try { this.activate(); } catch (error) { this.showError(error); }
    }, 120);
  }
  activate() {
    const view = this.currentView(), snapshot = this.snapshotFromView(view);
    const identity = snapshot ? { view, path: snapshot.filePath } : null;
    const changed = !identity || !this.activeIdentity || identity.view !== this.activeIdentity.view || identity.path !== this.activeIdentity.path;
    this.activeIdentity = identity;
    if (!snapshot) { this.pending = null; this.renderStatus(); return; }
    const first = !this.seen.has(snapshot.filePath); this.seen.add(snapshot.filePath);
    this.renderStatus(snapshot);
    if (changed && this.settings.autoCount && (first || this.settings.updateOnReturn)) this.requestCount(snapshot, false);
    else if (changed) this.pending = null;
  }
  validCache(snapshot) {
    return this.cache.get(cacheId(snapshot)) || null;
  }
  removePathCache(filePath) {
    for (const [key, entry] of this.cache) if (entry.filePath === filePath || entry.filePath.startsWith(filePath + "/")) this.cache.delete(key);
  }
  refresh() {
    if (this.disposed) return;
    const snapshot = this.currentSnapshot();
    if (!snapshot) { new Notice("请先打开一篇 Markdown 笔记。"); return; }
    this.requestCount(snapshot, true);
  }
  sameSnapshot(a, b) { return a && b && a.filePath === b.filePath && a.contentHash === b.contentHash && a.conversionKey === b.conversionKey; }
  requestCount(snapshot, force) {
    if (this.disposed) return;
    if (!force && this.validCache(snapshot)) { this.pending = null; this.renderStatus(); return; }
    if (this.running && this.running.epoch === this.cacheEpoch && this.sameSnapshot(this.running.snapshot, snapshot)) { this.pending = null; this.renderStatus(); return; }
    this.pending = { snapshot, settings: { ...this.settings }, epoch: this.cacheEpoch };
    this.errors.delete(snapshot.filePath);
    this.renderStatus();
    if (!this.pumpPromise) {
      this.pumpPromise = this.pump().catch(error => { if (!this.disposed) this.showError(error); }).finally(() => { this.pumpPromise = null; });
    }
  }
  async pump() {
    while (this.pending && !this.disposed) {
      const job = this.pending; this.pending = null;
      const controller = new AbortController(); this.running = { ...job, controller };
      let tempDirectory = null;
      try {
        if (process.platform !== "win32" || (Platform && Platform.isMobile)) throw new Error("本插件仅支持 Windows 桌面版 Obsidian 和 Microsoft Word。");
        this.renderStatus();
        const docx = await markdownToDocx(job.snapshot.content, job.settings);
        tempDirectory = await fsp.mkdtemp(path.join(os.tmpdir(), "count-by-msword-obsidian-"));
        const temporaryPath = path.join(tempDirectory, "count.docx");
        await fsp.writeFile(temporaryPath, docx.buffer, { flag: "wx" });
        if (this.disposed) controller.abort();
        const result = await wordStatistics(temporaryPath, job.settings.timeoutSeconds, controller.signal);
        if (!this.disposed && job.epoch === this.cacheEpoch && job.snapshot.conversionKey === conversionKey(this.settings)) {
          const entry = { filePath: job.snapshot.filePath, contentHash: job.snapshot.contentHash, conversionKey: job.snapshot.conversionKey, wordCount: result.statistics.words, statisticsVersion: STATISTICS_VERSION, statistics: result.statistics, statisticsErrors: result.statisticsErrors, calculatedAt: Date.now() };
          this.cache.delete(cacheId(entry)); this.cache.set(cacheId(entry), entry);
          while (this.cache.size > 100) this.cache.delete(this.cache.keys().next().value);
          this.errors.delete(entry.filePath); this.persist();
        }
      } catch (error) {
        if (!this.disposed && !error.cancelled && job.epoch === this.cacheEpoch) {
          this.errors.set(job.snapshot.filePath, { filePath: job.snapshot.filePath, contentHash: job.snapshot.contentHash, conversionKey: job.snapshot.conversionKey, message: error.message, detail: error.detail || "", missingWord: !!error.missingWord });
          while (this.errors.size > 100) this.errors.delete(this.errors.keys().next().value);
          if (this.sameSnapshot(this.currentSnapshot(), job.snapshot)) this.showError(error);
        }
      } finally {
        if (tempDirectory) {
          try { await removeTempDirectory(tempDirectory); }
          catch (error) { if (!this.disposed) this.showError(new Error("临时文档清理失败：" + error.message)); else console.error("count-by-msword-obsidian: temporary document cleanup failed", error); }
        }
        this.running = null; if (!this.disposed) this.renderStatus();
      }
    }
  }
  createStatisticsHover() {
    const doc = this.status.ownerDocument;
    this.hoverDocument = doc;
    const element = (tag, parent, className, text) => {
      const el = doc.createElement(tag);
      if (className) el.className = className;
      if (text) el.textContent = text;
      parent.appendChild(el); return el;
    };
    this.statisticsHover = element("div", doc.body, "count-by-msword-obsidian-hover");
    this.statisticsHover.id = "word-statistics-" + crypto.randomBytes(6).toString("hex");
    this.statisticsHover.hidden = true;
    this.statisticsHover.setAttribute("role", "dialog");
    this.statisticsHover.setAttribute("aria-modal", "false");
    this.status.setAttribute("aria-controls", this.statisticsHover.id);
    this.status.setAttribute("aria-haspopup", "dialog");
    this.status.setAttribute("aria-expanded", "false");
    const heading = element("div", this.statisticsHover, "count-by-msword-obsidian-hover-heading", "Word 文档统计");
    heading.id = this.statisticsHover.id + "-heading";
    this.statisticsHover.setAttribute("aria-labelledby", heading.id);
    this.hoverState = element("div", this.statisticsHover, "count-by-msword-obsidian-hover-meta");
    const table = element("table", this.statisticsHover, "count-by-msword-obsidian-hover-table");
    const tbody = element("tbody", table);
    this.hoverValues = new Map();
    for (const [key, label] of STATISTICS) {
      const row = element("tr", tbody), heading = element("th", row, "", label);
      heading.setAttribute("scope", "row");
      this.hoverValues.set(key, element("td", row, "", "—"));
    }
    this.hoverTime = element("div", this.statisticsHover, "count-by-msword-obsidian-hover-meta");
    this.hoverPath = element("div", this.statisticsHover, "count-by-msword-obsidian-hover-path");
    this.hoverError = element("div", this.statisticsHover, "count-by-msword-obsidian-error");
    element("div", this.statisticsHover, "count-by-msword-obsidian-hover-meta", "行数、页数按导出文档排版统计。点击状态栏刷新。");
    this.hoverExport = element("button", this.statisticsHover, "count-by-msword-obsidian-hover-export", "点击导出 DOCX");
    this.hoverExport.type = "button";
    this.registerDomEvent(this.hoverExport, "click", () => {
      const file = this.currentView()?.file;
      this.hideStatisticsHover();
      if (file && file.extension === "md") void this.exportNote(file);
    });
    const show = () => {
      if (this.disposed) return;
      clearTimeout(this.hoverHideTimer);
      this.statisticsHover.hidden = false;
      this.status.setAttribute("aria-expanded", "true");
      this.renderStatus(); this.positionStatisticsHover();
    };
    const scheduleHide = () => {
      clearTimeout(this.hoverHideTimer);
      this.hoverHideTimer = setTimeout(() => {
        if (!this.hoverOverStatus && !this.hoverOverPanel && !this.status.contains(doc.activeElement) && !this.statisticsHover.contains(doc.activeElement)) this.hideStatisticsHover();
      }, 180);
    };
    this.registerDomEvent(this.status, "mouseenter", () => { this.hoverOverStatus = true; show(); });
    this.registerDomEvent(this.status, "mouseleave", () => { this.hoverOverStatus = false; scheduleHide(); });
    this.registerDomEvent(this.statisticsHover, "mouseenter", () => { this.hoverOverPanel = true; clearTimeout(this.hoverHideTimer); });
    this.registerDomEvent(this.statisticsHover, "mouseleave", () => { this.hoverOverPanel = false; scheduleHide(); });
    this.registerDomEvent(this.status, "focusin", show);
    this.registerDomEvent(this.status, "focusout", scheduleHide);
    this.registerDomEvent(this.statisticsHover, "focusout", scheduleHide);
    const escape = event => {
      if (event.key === "Escape" && !this.statisticsHover.hidden) {
        event.preventDefault(); this.status.focus(); this.hideStatisticsHover();
      }
    };
    this.registerDomEvent(this.status, "keydown", escape);
    this.registerDomEvent(this.statisticsHover, "keydown", escape);
    if (doc.defaultView) this.registerDomEvent(doc.defaultView, "resize", () => this.positionStatisticsHover());
  }
  hideStatisticsHover() {
    clearTimeout(this.hoverHideTimer);
    this.hoverOverStatus = false; this.hoverOverPanel = false;
    if (this.statisticsHover) this.statisticsHover.hidden = true;
    if (this.status) this.status.setAttribute("aria-expanded", "false");
  }
  positionStatisticsHover() {
    const panel = this.statisticsHover;
    if (!panel || panel.hidden || this.disposed) return;
    const rect = this.status.getBoundingClientRect(), size = panel.getBoundingClientRect();
    const win = this.hoverDocument.defaultView;
    const width = win ? win.innerWidth : this.hoverDocument.documentElement.clientWidth;
    const height = win ? win.innerHeight : this.hoverDocument.documentElement.clientHeight;
    panel.style.left = Math.max(8, Math.min(rect.right - size.width, width - size.width - 8)) + "px";
    panel.style.top = Math.max(8, Math.min(rect.top - size.height - 6, height - size.height - 8)) + "px";
  }
  renderStatisticsHover(snapshot, entry, valid, error, busy) {
    if (!this.statisticsHover) return;
    this.hoverState.textContent = !snapshot ? "打开 Markdown 笔记后可统计字数。" :
      busy ? (entry ? "正在更新 · 下方为上次统计结果" : "正在统计…") :
      valid ? "缓存有效" : entry ? "缓存已过期 · 下方为上次统计结果" : "尚无统计结果";
    for (const [key] of STATISTICS) {
      const cell = this.hoverValues.get(key), value = entry?.statistics?.[key];
      cell.textContent = Number.isSafeInteger(value) ? value.toLocaleString("en-US") : entry ? "不可用" : "—";
      // Metric errors are already shown below the table; no nested tooltips.
      cell.removeAttribute("title");
    }
    this.hoverTime.textContent = "上次统计：" + (entry ? new Date(entry.calculatedAt).toLocaleString() : "无");
    this.hoverPath.textContent = "文档：" + (snapshot ? snapshot.filePath : "无");
    const metricErrors = entry ? STATISTICS.filter(([key]) => entry.statisticsErrors?.[key]).map(([key, label]) => label + "：" + entry.statisticsErrors[key]) : [];
    this.hoverError.textContent = [error ? "最近错误：" + error.message + (this.settings.detailedErrors && error.detail ? "\n" + error.detail : "") : "", ...metricErrors].filter(Boolean).join("\n");
    this.hoverError.hidden = !this.hoverError.textContent;
    this.hoverExport.disabled = !snapshot || !!this.exporting;
    this.positionStatisticsHover();
  }
  renderStatus(snapshot = this.currentSnapshot()) {
    if (!this.status || this.disposed) return;
    // Obsidian creates its own tooltip from aria-label, in addition to title.
    // The status button is named by its visible text; the dialog by its heading.
    this.status.removeAttribute("title");
    this.status.removeAttribute("aria-label");
    if (!snapshot) { this.status.setText("Word: —"); this.renderStatisticsHover(null, null, null, null, false); return; }
    const valid = this.validCache(snapshot);
    let entry = valid;
    if (!entry) for (const cached of this.cache.values()) if (cached.filePath === snapshot.filePath) entry = cached;
    const error = this.errors.get(snapshot.filePath), currentError = error && this.sameSnapshot(error, snapshot);
    const busy = (this.running && this.sameSnapshot(this.running.snapshot, snapshot)) || (this.pending && this.sameSnapshot(this.pending.snapshot, snapshot));
    let text;
    if (busy) text = "Word: 统计中…";
    else if (currentError) text = error.missingWord ? "Word: 未安装" : "Word: 统计失败";
    else if (entry) text = "Word: " + entry.wordCount.toLocaleString("en-US") + " 字" + (valid ? "" : " · 待更新");
    else text = "Word: 待统计";
    if (this.settings.showTime && entry && !busy) text += " · " + new Date(entry.calculatedAt).toLocaleTimeString();
    this.status.setText(text);
    this.renderStatisticsHover(snapshot, entry, valid, error, busy);
  }
  showError(error) {
    if (this.disposed) return;
    const detail = this.settings.detailedErrors && error.detail ? "\n" + error.detail : "";
    new Notice("Count By MSWord：" + (error.message || String(error)) + detail, 10000);
    if (this.settings.detailedErrors) console.error("count-by-msword-obsidian", error);
  }
  persist() {
    const state = { settings: { ...this.settings }, cache: [...this.cache.values()] };
    this.saveTail = this.saveTail.then(() => this.saveData(state)).catch(error => { if (!this.disposed) this.showError(new Error("保存插件设置或缓存失败：" + error.message)); });
    return this.saveTail;
  }
  clearCache(notify = true) {
    this.cacheEpoch++; this.cache.clear(); this.errors.clear(); this.pending = null;
    this.persist(); this.renderStatus(); if (notify) new Notice("Word 字数缓存已清除。");
  }
  async changeSetting(key, value) {
    const before = conversionKey(this.settings);
    this.settings = validatedSettings({ ...this.settings, [key]: value });
    if (conversionKey(this.settings) !== before) { this.cacheEpoch++; this.cache.clear(); this.errors.clear(); this.pending = null; }
    if (key === "autoCount" && !value) this.pending = null;
    await this.persist(); this.renderStatus();
  }
  async snapshotForFile(file) {
    const current = this.currentView();
    if (current && current.file?.path === file.path) return this.snapshotFromView(current);
    const views = this.app.workspace.getLeavesOfType("markdown").map(leaf => leaf.view).filter(view => view instanceof MarkdownView && view.file?.path === file.path);
    views.sort((a, b) => (this.edits.get(b) || 0) - (this.edits.get(a) || 0));
    if (views.length) return this.snapshotFromView(views[0]);
    return this.makeSnapshot(file.path, await this.app.vault.read(file));
  }
  async exportNote(file) {
    if (this.disposed || this.exporting) return;
    this.exporting = true;
    let stage = null;
    try {
      if (process.platform !== "win32") throw new Error("DOCX 导出位置选择仅支持 Windows 本地路径。");
      if (!file) file = this.currentView()?.file;
      if (!file || file.extension !== "md") { new Notice("请先打开一篇 Markdown 笔记。"); return; }
      const snapshot = await this.snapshotForFile(file), settings = { ...this.settings };
      const adapter = this.app.vault.adapter;
      const base = typeof adapter.getBasePath === "function" ? adapter.getBasePath() : null;
      let directory = settings.exportDirectory.trim();
      if (!directory) {
        if (!base) throw new Error("当前仓库未提供本地目录，请先在设置中填写默认导出目录。");
        directory = path.join(base, path.dirname(file.path));
      }
      if (!/^[a-z]:[\\/]/i.test(directory)) throw new Error("默认导出目录必须是带盘符的 Windows 本地绝对路径。");
      const target = await new PathModal(this.app, path.join(directory, safeFilename(file.basename)), this.modals).openResult();
      if (!target || this.disposed) return;
      const parent = await fsp.stat(path.dirname(target));
      if (!parent.isDirectory()) throw new Error("导出目录不存在。");
      const docx = await markdownToDocx(snapshot.content, settings);
      let existing = null;
      try { existing = await fsp.lstat(target); } catch (error) { if (error.code !== "ENOENT") throw error; }
      if (existing) {
        if (!existing.isFile() || existing.isSymbolicLink()) throw new Error("目标不是普通文件，请选择其他导出路径。");
        if (!await new ConfirmModal(this.app, target, this.modals).openResult() || this.disposed) return;
        stage = path.join(path.dirname(target), ".wce-" + crypto.randomBytes(12).toString("hex") + ".tmp");
        await fsp.writeFile(stage, docx.buffer, { flag: "wx" });
        const latest = await fsp.lstat(target);
        if (!latest.isFile() || !sameFileStat(existing, latest)) throw new Error("目标文件在确认后发生变化，请重新导出。");
        if (this.disposed) return;
        await fsp.rename(stage, target); stage = null;
      } else {
        // Exclusive create protects a file created while the export dialog was open.
        if (this.disposed) return;
        await fsp.writeFile(target, docx.buffer, { flag: "wx" });
      }
      if (!this.disposed) {
        new Notice("Word 文档已导出：" + target, 7000);
        if (docx.warnings.length) new Notice(docx.warnings.join("\n"), 9000);
      }
    } catch (error) { if (!this.disposed) this.showError(new Error("导出失败：" + error.message)); }
    finally {
      if (stage) { try { await fsp.unlink(stage); } catch (error) { console.error("count-by-msword-obsidian: export staging cleanup failed", error); } }
      this.exporting = false;
    }
  }
  onunload() {
    this.disposed = true; clearTimeout(this.activationTimer); this.pending = null;
    this.hideStatisticsHover();
    if (this.statisticsHover) this.statisticsHover.remove();
    for (const modal of [...this.modals]) modal.close();
    if (this.running) this.running.controller.abort();
    // Obsidian removes registered events/DOM events and the status element.
    // The already-started asynchronous pump finishes owned-process/temp cleanup.
  }
}

class CountByMSWordSettings extends PluginSettingTab {
  constructor(app, plugin) { super(app, plugin); this.plugin = plugin; }
  display() {
    const { containerEl } = this, p = this.plugin; containerEl.empty();
    const toggle = (name, description, key) => new Setting(containerEl).setName(name).setDesc(description).addToggle(control => control.setValue(p.settings[key]).onChange(value => { void p.changeSetting(key, value); }));
    const text = (name, description, key) => new Setting(containerEl).setName(name).setDesc(description).addText(control => {
      control.setValue(String(p.settings[key]));
      // Commit on blur, avoiding cache writes for every typed character.
      control.inputEl.addEventListener("change", () => { void p.changeSetting(key, control.getValue().trim()); });
    });
    const number = (name, description, key, min, max) => new Setting(containerEl).setName(name).setDesc(description).addText(control => {
      control.setValue(String(p.settings[key])); control.inputEl.type = "number"; control.inputEl.min = String(min); control.inputEl.max = String(max); control.inputEl.step = "0.5";
      control.inputEl.addEventListener("change", () => {
        const value = Number(control.getValue());
        if (!Number.isFinite(value) || value < min || value > max) { new Notice("请输入 " + min + " 到 " + max + " 之间的数值。"); control.setValue(String(p.settings[key])); return; }
        void p.changeSetting(key, value);
      });
    });
    containerEl.createEl("h2", { text: "Word 字数统计" });
    toggle("自动统计", "首次打开或切换笔记时统计；编辑时仅标记待更新。", "autoCount");
    toggle("返回笔记时更新", "返回已打开笔记时检查内容，未变化则使用缓存。", "updateOnReturn");
    toggle("显示上次统计时间", "在状态栏附加时间；悬停提示始终显示完整时间。", "showTime");
    new Setting(containerEl).setName("点击状态栏强制刷新").setDesc("始终可用，忽略已有缓存；相同内容正在统计时合并请求。");
    new Setting(containerEl).setName("清除字数缓存").setDesc("最多持久保存最近 100 篇笔记。清除后可点击状态栏重新统计。").addButton(button => button.setButtonText("清除缓存").onClick(() => p.clearCache()));
    containerEl.createEl("h2", { text: "DOCX 导出" });
    text("默认导出目录", "留空使用当前笔记文件夹；可填写带盘符的 Windows 本地绝对路径。导出时可修改完整路径。", "exportDirectory");
    text("中文字体", "默认宋体，需在本机安装该字体。", "chineseFont");
    text("英文字体", "默认 Times New Roman。", "englishFont");
    new Setting(containerEl).setName("标点符号形式").setDesc("默认中文全角，中文引号使用中文字体；可选英文半角。转换常用标点，代码、公式、YAML、网址、邮箱和数字格式保留原样。").addDropdown(control => control
      .addOption("full", "全角（中文标点）")
      .addOption("half", "半角（英文标点）")
      .setValue(p.settings.punctuationMode)
      .onChange(value => { void p.changeSetting("punctuationMode", value); }));
    toggle("保留空行", "默认关闭：Markdown 段落之间的空行只用于分段，不生成空白段落；开启后保留。代码块、公式和 YAML 内的空行始终保留。", "keepBlankLines");
    number("正文字号 (pt)", "默认 12 pt。", "fontSize", 6, 72);
    number("行距（倍）", "默认 1.5 倍。", "lineSpacing", 1, 4);
    number("首行缩进（汉字）", "默认 2 个汉字；标题、列表、引用、代码不缩进。", "firstLineIndent", 0, 8);
    toggle("包含 YAML", "将 Frontmatter 内容作为原样文本加入正文；默认排除。", "includeYaml");
    containerEl.createEl("h2", { text: "高级设置" });
    number("Word 统计超时（秒）", "默认 30 秒；包含 Word 启动、统计和正常关闭。超时后另留最多约 7 秒清理已确认的实例。", "timeoutSeconds", 5, 180);
    toggle("显示详细错误信息", "在提示、悬停文字和开发者控制台中显示 COM 错误详情。", "detailedErrors");
    containerEl.createEl("h3", { text: "Markdown 转换规则" });
    const rules = containerEl.createEl("ul", { cls: "count-by-msword-obsidian-rules" });
    for (const rule of RULES) rules.createEl("li", { text: rule });
  }
}

module.exports = CountByMSWordObsidian;
