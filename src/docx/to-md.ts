import type { ConvertOptions, Doc } from '../types.js';
import { composeDocument } from '../md/compose.js';

/** DOCX Doc → Markdown。 */
export function docxToMarkdown(doc: Doc, options: ConvertOptions = {}): string {
  return composeDocument(doc, options);
}

/** DOCX Doc → JSON 结构。 */
export function docxToJson(doc: Doc): string {
  return JSON.stringify({ format: doc.format, meta: doc.meta, blocks: doc.blocks }, null, 2);
}