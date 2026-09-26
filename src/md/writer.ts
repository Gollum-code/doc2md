import type { Block, Doc, Inline } from '../types.js';

export interface MdOptions {
  tableStyle: 'gfm' | 'html' | 'none';
  escape: boolean;
  headingOffset: number;
  emptyCell: string;
  keepEmptyParagraphs: boolean;
}

/** 把 Doc IR 渲染为 Markdown。 */
export function renderMarkdown(doc: Doc, options: Partial<MdOptions> = {}): string {
  const o: MdOptions = {
    tableStyle: 'gfm',
    escape: true,
    headingOffset: 0,
    emptyCell: '',
    keepEmptyParagraphs: false,
    ...options,
  };
  const lines: string[] = [];
  renderBlocks(lines, doc.blocks, o);
  return lines.join('\n');
}

function renderBlocks(out: string[], blocks: Block[], o: MdOptions): void {
  let prevWasBlock = false;
  for (const block of blocks) {
    const buf: string[] = [];
    renderBlock(buf, block, o);
    const text = buf.join('\n').trimEnd();
    if (!text) continue;
    if (prevWasBlock) out.push('');
    out.push(text);
    prevWasBlock = true;
  }
}

function renderBlock(out: string[], block: Block, o: MdOptions): void {
  switch (block.type) {
    case 'heading': {
      const level = clamp(block.level + o.headingOffset, 1, 6);
      out.push(`${'#'.repeat(level)} ${inlineText(block.children, o)}`);
      break;
    }
    case 'paragraph': {
      const t = inlineText(block.children, o);
      if (!t && !o.keepEmptyParagraphs) break;
      if (block.align === 'center') {
        out.push(`<p align="center">${t}</p>`);
      } else if (block.align === 'right') {
        out.push(`<p align="right">${t}</p>`);
      } else {
        out.push(t || '');
      }
      break;
    }
    case 'quote': {
      const inner: string[] = [];
      renderBlocks(inner, block.children, o);
      for (const line of inner.join('\n').split('\n')) {
        out.push(line ? `> ${line}` : '>');
      }
      break;
    }
    case 'list': {
      for (const item of block.items) {
        renderListItems(out, item, block.ordered, block.start, 0, o);
      }
      break;
    }
    case 'table':
      renderTable(out, block, o);
      break;
    case 'code':
      out.push('```' + (block.lang ?? '') + '\n' + block.value.replace(/\n+$/g, '') + '\n```');
      break;
    case 'image':
      out.push(escapeImageRef(block.src, o) + (block.alt ? `\n\n*${o.escape ? escapeInlineText(block.alt) : block.alt}*` : ''));
      break;
    case 'hr':
      out.push('---');
      break;
    case 'section': {
      const metaLine = sectionMeta(block.meta);
      if (metaLine) out.push(metaLine);
      const level = clamp(block.level ?? 2, 1, 6);
      out.push(`${'#'.repeat(level)} ${inlineText([{ kind: 'text', value: block.title }], o)}`);
      renderBlocks(out, block.blocks, o);
      break;
    }
    case 'footnotes': {
      out.push(`## ${block.title}`);
      for (const item of block.items) {
        const body: string[] = [];
        renderBlocks(body, item.blocks, o);
        out.push(`[^${item.label}]: ${body.join(' ')}`);
      }
      break;
    }
    case 'preformatted':
      out.push(block.value);
      break;
  }
}

function sectionMeta(meta?: Record<string, string>): string | undefined {
  if (!meta) return undefined;
  const parts: string[] = [];
  for (const [k, v] of Object.entries(meta)) {
    if (v) parts.push(`${k}: ${v}`);
  }
  return parts.length ? `<sup>${parts.join(' · ')}</sup>` : undefined;
}

function renderListItems(
  out: string[],
  item: Block[],
  ordered: boolean,
  start: number,
  indent: number,
  o: MdOptions,
): void {
  const itemLines: string[] = [];
  const nestedBuf: string[] = [];
  for (const block of item) {
    if (block.type === 'list') {
      // 子列表：延后到父项自己的文本之后输出，并层级 +1
      const tmp: string[] = [];
      for (const sub of block.items) {
        renderListItems(tmp, sub, block.ordered, block.start, indent + 1, o);
      }
      nestedBuf.push(...tmp);
      continue;
    }
    const buf: string[] = [];
    renderBlock(buf, block, o);
    for (const line of buf.join('\n').split('\n')) {
      itemLines.push(line);
    }
  }
  const marker = ordered ? `${start}.` : '-';
  const pad = '  '.repeat(indent);
  if (itemLines.length) {
    out.push(`${pad}${marker} ${itemLines[0]}`);
    for (let i = 1; i < itemLines.length; i++) {
      out.push(`${pad}  ${itemLines[i]}`);
    }
  } else {
    out.push(`${pad}${marker}`);
  }
  // 父项文本输出完毕后，再输出子列表
  for (const line of nestedBuf) {
    out.push(line);
  }
}

function renderTable(out: string[], table: Block & { type: 'table' }, o: MdOptions): void {
  if (o.tableStyle === 'none') {
    for (const row of [table.header, ...table.rows]) {
      out.push(row.map((c) => escapeTableCell(c, o)).join(' | '));
    }
    return;
  }
  if (o.tableStyle === 'html') {
    out.push('<table>');
    if (table.header.length) {
      out.push('<thead><tr>' + table.header.map((c) => `<th>${escapeHtml(c)}</th>`).join('') + '</tr></thead>');
    }
    out.push('<tbody>');
    for (const row of table.rows) {
      out.push('<tr>' + row.map((c) => `<td>${escapeHtml(c)}</td>`).join('') + '</tr>');
    }
    out.push('</tbody></table>');
    return;
  }
  // gfm 管道表格
  const colCount = Math.max(
    table.header.length,
    ...table.rows.map((r) => r.length),
  );
  const cells = (row: string[]): string[] => {
    const c = row.slice();
    while (c.length < colCount) c.push('');
    return c;
  };
  const rows: string[][] = [];
  if (table.header.length) rows.push(cells(table.header));
  for (const row of table.rows) rows.push(cells(row));
  const separator = (): string =>
    Array.from({ length: colCount }, (_, i) => {
      const a = table.align?.[i];
      const dashes = '---';
      if (a === 'center') return ':---:';
      if (a === 'right') return '---:';
      return dashes;
    }).join(' | ');
  const fmtRow = (row: string[]): string => row.map((c) => escapeTableCell(c, o)).join(' | ');
  rows.forEach((row, i) => {
    out.push(`| ${fmtRow(row)} |`);
    if (i === 0 && table.header.length) out.push(`| ${separator()} |`);
  });
}

function escapeTableCell(cell: string, o: MdOptions): string {
  if (!o.escape) return cell;
  return cell.replace(/\r?\n/g, '<br>').replace(/\|/g, '\\|');
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function escapeImageRef(src: string, o: MdOptions): string {
  if (!o.escape) return `![${src}](${src})`;
  const s = src.replace(/[()]/g, (m) => (m === '(' ? '%28' : '%29'));
  return `![${s}](${s})`;
}

/** 把 Inline[] 拼成一段文本（转义 + 格式化）。 */
export function inlineText(children: Inline[], o: MdOptions): string {
  if (!children.length) return '';
  let out = '';
  for (const node of children) {
    out += inlineNode(node, o);
  }
  return out;
}

function inlineNode(node: Inline, o: MdOptions): string {
  switch (node.kind) {
    case 'text':
      return o.escape ? escapeInlineText(node.value) : node.value;
    case 'code':
      return '`' + node.value.replace(/`/g, '``').replace(/\n/g, ' ') + '`';
    case 'strong':
      return `**${inlineText(node.children, o)}**`;
    case 'em':
      return `*${inlineText(node.children, o)}*`;
    case 'del':
      return `~~${inlineText(node.children, o)}~~`;
    case 'sup':
      return `<sup>${inlineText(node.children, o)}</sup>`;
    case 'sub':
      return `<sub>${inlineText(node.children, o)}</sub>`;
    case 'link':
      return `[${inlineText(node.children, o)}](${o.escape ? escapeUrl(node.href) : node.href})`;
    case 'image':
      return `![${node.alt ?? ''}](${o.escape ? escapeUrl(node.src) : node.src})`;
    case 'break':
      return '\n';
    case 'pageBreak':
      return '\n\n---\n';
  }
}

function escapeUrl(url: string): string {
  return url.replace(/[()\s]/g, (m) => {
    if (m === '(') return '%28';
    if (m === ')') return '%29';
    if (m === ' ') return '%20';
    return m;
  });
}

/**
 * 段落文本转义 —— 只转义会破坏 Markdown 结构的字符，保持文本可读。
 * 不转义普通出现的 `*`/`_`，只在行首危险位置处理。
 */
export function escapeInlineText(text: string): string {
  if (!text) return text;
  let s = text.replace(/`/g, '\\`');
  // 转义行首的 markdown 标记（#、>、-、+、数字.、*、_、=、[、|、~）
  s = s.replace(/(^|\n)([ \t]*)(#{1,6})(\s)/g, (_m, lead: string, ws: string, hashes: string, sp: string) => {
    return `${lead}${ws}\\${hashes}${sp}`;
  });
  s = s.replace(/(^|\n)([ \t]*)(>+)(\s)/g, (_m, lead: string, ws: string, g: string, sp: string) => {
    return `${lead}${ws}\\${g}${sp}`;
  });
  // 列表符号本身不转义（保留列表结构），但 "- " 在引用里会歧义——这里仅转义横线段落(hr)
  s = s.replace(/(^|\n)([ \t]*)(- ?(- ?-)+)(\s*)$/g, '\\$&');
  // 数字开头的行，若非有序列表上下文，也转义防自动编号干扰
  s = s.replace(
    /(^|\n)([ \t]*)([0-9]{1,9})([.)])([ \t])/g,
    (_m, lead: string, ws: string, n: string, dot: string, sp: string) =>
      `${lead}${ws}${n}\\${dot}${sp}`,
  );
  return s;
}

function clamp(n: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, n));
}
