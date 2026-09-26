import type { ConvertOptions, Doc } from '../types.js';
import { composeDocument } from '../md/compose.js';

/** PPTX Doc → Markdown。 */
export function pptxToMarkdown(doc: Doc, options: ConvertOptions = {}): string {
  return composeDocument(doc, options);
}

/** PPTX Doc → JSON 结构（供二次处理）。 */
export function pptxToJson(doc: Doc): string {
  return JSON.stringify({ format: doc.format, meta: doc.meta, blocks: doc.blocks }, null, 2);
}
