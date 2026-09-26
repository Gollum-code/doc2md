import type { ConvertOptions, Doc } from '../types.js';
import { composeDocument } from '../md/compose.js';

/** XLSX Doc → Markdown。 */
export function xlsxToMarkdown(doc: Doc, options: ConvertOptions = {}): string {
  return composeDocument(doc, options);
}

/** XLSX Doc → JSON 结构。 */
export function xlsxToJson(doc: Doc): string {
  return JSON.stringify({ format: doc.format, meta: doc.meta, blocks: doc.blocks }, null, 2);
}
