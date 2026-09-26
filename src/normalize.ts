import type { NormalizeOptions } from './types.js';

/**
 * Markdown 中文/标点/排版规范管线。
 *
 * 面向 RAG 的目标：统一全角字符、补齐中英空位、清理零宽字符，
 * 同时不破坏 Markdown 结构（代码块、表格、行内 code 保持原样）。
 */

export const normalizeOptionDefaults = (partial: NormalizeOptions = {}): NormalizeOptions => ({
  cjkLatinSpacing: true,
  fullWidthToHalfWidth: true,
  punctuationToHalfWidth: false,
  collapseBlankLines: true,
  normalizeBullets: true,
  normalizeOrderedMarkers: true,
  ideographicSpace: true,
  stripInvisible: true,
  ...partial,
});

const CJK = '\\p{Script=Han}';
// 注意：V8 字符类中只能使用 \p{Script=...} / \p{...=...} 形式
const LATIN_OR_DIGIT = '\\p{Script=Latin}\\p{Nd}';

/** CJK 与 ASCII 之间补空格；数字与 % 等符号跳动不再错位。 */
function insertCjkLatinSpacing(line: string): string {
  // CJK → Latin/digit：避免断在已有空格两侧
  let s = line.replace(
    new RegExp(`(${CJK})(?=[${LATIN_OR_DIGIT}])`, 'gu'),
    '$1 ',
  );
  // Latin/digit → CJK：接在 ASCII 后
  s = s.replace(
    new RegExp(`([${LATIN_OR_DIGIT}])(?=${CJK})`, 'gu'),
    '$1 ',
  );
  return s;
}

/**
 * 全角字符转半角。
 * - 默认：仅字母数字（Ａ→A、１→1）与全角斜杠。
 * - punctuationToHalfWidth：额外将中文标点转英文标点。
 */
function toHalfWidth(line: string, punctuation: boolean): string {
  let s = line.replace(/[\uFF21-\uFF3A]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0)); // Ａ-Ｚ
  s = s.replace(/[\uFF41-\uFF5A]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0)); // ａ-ｚ
  s = s.replace(/[\uFF10-\uFF19]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0)); // ０-９
  if (punctuation) {
    s = s
      .replace(/，/g, ',')
      .replace(/。/g, '.')
      .replace(/；/g, ';')
      .replace(/：/g, ':')
      .replace(/！/g, '!')
      .replace(/？/g, '?')
      .replace(/（/g, '(')
      .replace(/）/g, ')')
      .replace(/【/g, '[')
      .replace(/】/g, ']')
      .replace(/‘/g, "'")
      .replace(/’/g, "'")
      .replace(/“/g, '"')
      .replace(/”/g, '"')
      .replace(/＜/g, '<')
      .replace(/＞/g, '>')
      .replace(/＝/g, '=')
      .replace(/＋/g, '+')
      .replace(/－/g, '-')
      .replace(/％/g, '%')
      .replace(/＆/g, '&')
      .replace(/＠/g, '@')
      .replace(/＃/g, '#')
      .replace(/＊/g, '*')
      .replace(/￣/g, '~');
  }
  return s;
}

/** 清理不可见字符。 */
function stripInvisibleChars(s: string): string {
  return s
    // 零宽字符 + BOM + 软连字符
    .replace(/[\u200B\u200C\u200D\uFEFF\u00AD\u2060\u200E\u200F]/g, '')
    // NBSP → 普通空格
    .replace(/\u00A0/g, ' ')
    .replace(/\u2007/g, ' ')
    .replace(/\u202F/g, ' ');
}

/** 全角空格（常见于中文排版缩进）→ 半角空格（可选保留为缩进语义）。 */
function convertIdeographicSpace(s: string): string {
  return s.replace(/\u3000/g, ' ');
}

/** 行首项目符号归一化。 */
function normalizeBulletMarkers(s: string): string {
  // 必须后跟空格/制表符才算列表符号，避免误伤 "•abc"
  return s.replace(/^([ \t]*)[•·▪●○‣⟶→▸►■□✦✧▶◆]([ \t])/, '$1-$2');
}

/** 中文序号列表 `1、` `1）` `1．` `1. ` → 有序列表 `1. `（`. ` 需后跟空格，避免误伤日期）。 */
function normalizeOrderedMarkers(s: string): string {
  return s.replace(/^([ \t]*)(\d{1,5})([、)．]|[.](?=[ \t]|$))/, (_m, ws: string, n: string) => {
    return `${ws}${n}. `;
  });
}

function fenceOpen(line: string): boolean {
  return /^\s*(`{3,}|~{3,})\s*$/.test(line) || /^\s*(`{3,}|~{3,})\S/.test(line);
}

function fenceClose(line: string): boolean {
  return /^\s*(`{3,}|~{3,})\s*$/.test(line);
}

function isTableLine(line: string): boolean {
  const t = line.trim();
  if (!t) return false;
  // gfm 表格行列：以 | 开头或结尾，或含 | 且行内出现在 4 个以上
  if (t.startsWith('|') || t.endsWith('|')) return true;
  return (t.match(/\|/g) ?? []).length >= 2 && /^\s*\|.*\|.*\s*$/.test(t);
}

/**
 * 对整篇 Markdown 做规范化。
 */
export function normalizeMarkdown(input: string, userOptions: NormalizeOptions = {}): string {
  const opt = normalizeOptionDefaults(userOptions);
  const src = String(input ?? '').replace(/\r\n?/g, '\n');
  const lines = src.split('\n');

  const out: string[] = [];
  let fence = false;

  for (const rawLine of lines) {
    if (!fence && fenceOpen(rawLine)) {
      fence = true;
      out.push(rawLine);
      continue;
    }
    if (fence && fenceClose(rawLine)) {
      fence = false;
      out.push(rawLine);
      continue;
    }
    if (fence) {
      out.push(rawLine);
      continue;
    }

    let line = rawLine;
    if (opt.stripInvisible) line = stripInvisibleChars(line);
    if (opt.ideographicSpace) line = convertIdeographicSpace(line);

    if (isTableLine(line)) {
      // 表格行：只清尾部空格与行前控制空格
      out.push(line.trimEnd());
      continue;
    }

    if (opt.fullWidthToHalfWidth) line = toHalfWidth(line, opt.punctuationToHalfWidth ?? false);
    if (opt.cjkLatinSpacing) line = insertCjkLatinSpacing(line);
    if (opt.normalizeBullets) line = normalizeBulletMarkers(line);
    if (opt.normalizeOrderedMarkers) line = normalizeOrderedMarkers(line);

    // 列表标记行保留缩进（嵌套列表需要）；普通段落去掉行首空格。
    const listLike = /^(?:[-*+]|\d{1,5}[.)]|>)\s/.test(line.trimStart());
    if (!listLike) {
      line = line.replace(/^[ \t]+/, '');
    }
    // 行内 2+ 连续空格压缩为一个（保留行首缩进，避免破坏嵌套列表）
    const lead = (/^[ \t]*/.exec(line) ?? [''])[0];
    const rest = line.slice(lead.length).replace(/[ \t]{2,}/g, (m) => (/\s$/.test(m) ? m[m.length - 1] : ' '));
    // 行尾空格清理（只清 rest）
    line = lead + rest.trimEnd();
    out.push(line);
  }

  let result = out.join('\n');

  if (opt.collapseBlankLines) {
    result = result.replace(/\n{3,}/g, '\n\n');
  }
  // 首尾不留空行
  return result.replace(/^\n+/, '').replace(/\n+$/, '') + '\n';
}

/**
 * 对单段“纯文本”（非 markdown）做中文规范化。
 * 供解析器在拼段落时直接用，或在 title/alt 等字段上调用。
 */
export function normalizeText(input: string, userOptions: NormalizeOptions = {}): string {
  const opt = normalizeOptionDefaults(userOptions);
  let s = String(input ?? '');
  if (opt.stripInvisible) s = stripInvisibleChars(s);
  if (opt.ideographicSpace) s = convertIdeographicSpace(s);
  if (opt.fullWidthToHalfWidth) s = toHalfWidth(s, opt.punctuationToHalfWidth ?? false);
  if (opt.cjkLatinSpacing) s = insertCjkLatinSpacing(s);
  return s.replace(/\s+/g, ' ').trim();
}