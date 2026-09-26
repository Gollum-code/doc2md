import type { Block, ConvertOptions, Doc } from '../types.js';
import { normalizeMarkdown } from '../normalize.js';
import { renderMarkdown, type MdOptions } from './writer.js';

export interface ComposeOptions extends ConvertOptions {
  /** 标题 H1 的层级，缺省自动（1）。 */
  titleLevel?: number;
  /** 强制插入标题 H1（默认自动：正文已有 H1 时不插）。 */
  forceTitle?: boolean;
}

/**
 * 统一出图管线：渲染 → 标题/front-matter → 中文规范化。
 */
export function composeDocument(doc: Doc, options: ComposeOptions = {}): string {
  const mdOpts: MdOptions = {
    tableStyle: options.tableStyle ?? 'gfm',
    escape: options.escape !== false,
    headingOffset: options.headingOffset ?? 0,
    emptyCell: options.emptyCell ?? '',
    keepEmptyParagraphs: options.keepEmptyParagraphs === true,
  };

  let body = renderMarkdown(doc, mdOpts);
  const head: string[] = [];

  if (options.frontMatter) {
    head.push(renderFrontMatter(doc));
  }

  const title = doc.meta.title?.trim();
  if (title) {
    const level = Math.min(6, Math.max(1, options.titleLevel ?? 1));
    const alreadyHas = hasHeadingAtLevel(doc.blocks, level);
    if (options.forceTitle || !alreadyHas) {
      head.push(`${'#'.repeat(level)} ${title}`);
    }
  }

  const parts = [...head, body].filter((p) => p && p.trim());
  let md = parts.join('\n\n');

  if (options.normalize !== false) {
    md = normalizeMarkdown(md, options.normalizeOptions);
  } else {
    md = md.replace(/\n{3,}/g, '\n\n').trimEnd() + '\n';
  }
  return md;
}

function hasHeadingAtLevel(blocks: Block[], level: number): boolean {
  for (const b of blocks) {
    if (b.type === 'heading' && b.level === level) return true;
    if (b.type === 'section' && (b.level ?? 2) === level) return true;
    if (b.type === 'quote' && hasHeadingAtLevel(b.children, level)) return true;
  }
  return false;
}

function renderFrontMatter(doc: Doc): string {
  const lines: string[] = ['---'];
  const put = (k: string, v: unknown): void => {
    if (v === undefined || v === null) return;
    const s = Array.isArray(v) ? v.join(', ') : String(v);
    if (!s.trim()) return;
    const quoted = /[:#\[\]{}>|&*!?'"%@`\n]/.test(s) ? JSON.stringify(s) : s;
    lines.push(`${k}: ${quoted}`);
  };
  put('title', doc.meta.title);
  put('author', doc.meta.author);
  put('subject', doc.meta.subject);
  put('keywords', doc.meta.keywords);
  put('created', doc.meta.created);
  put('modified', doc.meta.modified);
  put('format', doc.format);
  put('slides', doc.meta.slides);
  put('sheets', doc.meta.sheets);
  lines.push('---');
  return lines.join('\n');
}
