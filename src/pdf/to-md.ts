import type { ConvertOptions, Doc } from '../types.js';
import { composeDocument } from '../md/compose.js';

/** PDF Doc → Markdown。 */
export function pdfToMarkdown(doc: Doc, options: ConvertOptions = {}): string {
  return composeDocument(doc, options);
}

/** PDF Doc → JSON 结构。 */
export function pdfToJson(doc: Doc): string {
  return JSON.stringify({ format: doc.format, meta: doc.meta, blocks: doc.blocks }, null, 2);
}