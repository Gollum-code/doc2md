import { readFile, mkdir, writeFile } from 'node:fs/promises';
import { dirname, extname, resolve, basename } from 'node:path';
import type { ConvertOptions, ConvertResult, Doc, DocFormat } from './types.js';
import { OoxmlPackage } from './core/ooxml.js';
import { createParseContext } from './core/context.js';
import { parsePptx } from './pptx/parser.js';
import { parseDocx } from './docx/parser.js';
import { parseXlsx } from './xlsx/parser.js';
import { composeDocument } from './md/compose.js';

/** doc2md 错误。 */
export class Doc2mdError extends Error {
  constructor(
    message: string,
    readonly code:
      | 'UNSUPPORTED_FORMAT'
      | 'PARSE_ERROR'
      | 'FILE_NOT_FOUND'
      | 'NOT_IMPLEMENTED'
      | 'PACKAGE_ERROR'
      | 'IO_ERROR',
  ) {
    super(message);
    this.name = 'Doc2mdError';
  }
}

export type { Doc, DocFormat, ConvertOptions, ConvertResult, Block, Inline, DocMeta } from './types.js';
export { normalizeMarkdown, normalizeText, normalizeOptionDefaults } from './normalize.js';
export { renderMarkdown } from './md/writer.js';
export { composeDocument } from './md/compose.js';
export type { PendingAsset } from './core/context.js';

/** 魔数嗅探格式。 */
export function detectFormat(data: Uint8Array, fileName?: string): DocFormat {
  const ext = (fileName ? extname(fileName) : '').toLowerCase();
  if (ext === '.md' || ext === '.markdown' || ext === '.txt' || ext === '.mdx') return 'markdown';
  if (isZip(data)) {
    try {
      const pkg = OoxmlPackage.fromBufferSync(data);
      if (pkg.has('ppt/presentation.xml')) return 'pptx';
      if (pkg.has('word/document.xml')) return 'docx';
      if (pkg.has('xl/workbook.xml')) return 'xlsx';
      return 'unknown';
    } catch {
      return 'unknown';
    }
  }
  if (data.length >= 5) {
    const head = String.fromCharCode(...data.subarray(0, 5));
    if (head === '%PDF-') return 'pdf';
  }
  return 'unknown';
}

function isZip(data: Uint8Array): boolean {
  return data.length >= 4 && data[0] === 0x50 && data[1] === 0x4b && (data[2] === 0x03 || data[2] === 0x05 || data[2] === 0x07);
}

/** 解析文档为 IR（同步；图片引用已分配但未写盘）。 */
export function parseDocument(data: Uint8Array, options: ConvertOptions = {}): { doc: Doc; assets: ReturnType<typeof createParseContext>['pending']; warnings: string[] } {
  const format = options.format ?? detectFormat(data, options.fileName);
  const bundle = createParseContext(options);

  let doc: Doc;
  try {
    switch (format) {
      case 'pptx': {
        const pkg = OoxmlPackage.fromBufferSync(data);
        doc = parsePptx(pkg, bundle.ctx);
        break;
      }
      case 'docx': {
        const pkg = OoxmlPackage.fromBufferSync(data);
        doc = parseDocx(pkg, bundle.ctx);
        break;
      }
      case 'xlsx': {
        const pkg = OoxmlPackage.fromBufferSync(data);
        doc = parseXlsx(pkg, bundle.ctx);
        break;
      }
      case 'markdown': {
        const text = new TextDecoder().decode(data);
        doc = markdownAsDoc(text);
        break;
      }
      case 'pdf':
        throw new Doc2mdError('PDF 解析器尚未实现（Roadmap M4），当前支持 pptx / docx / xlsx。', 'NOT_IMPLEMENTED');
      case 'unknown':
        throw new Doc2mdError(
          `无法识别文件格式${options.fileName ? `：${options.fileName}` : ''}。支持 pptx / docx / xlsx。`,
          'UNSUPPORTED_FORMAT',
        );
    }
  } catch (err) {
    if (err instanceof Doc2mdError) throw err;
    throw new Doc2mdError(`解析失败：${(err as Error).message}`, 'PARSE_ERROR');
  }

  return { doc, assets: bundle.pending, warnings: bundle.warnings };
}

/** 转换：解析 + 渲染 + 写图片。 */
export async function convert(data: Uint8Array | ArrayBuffer | string, options: ConvertOptions = {}): Promise<ConvertResult> {
  const buf = toBytes(data);
  const { doc, assets, warnings } = parseDocument(buf, options);

  const markdown = composeDocument(doc, options);

  const written: { file: string; source: string }[] = [];
  if (assets.length) {
    await mkdir(dirname(assets[0].absPath), { recursive: true });
    for (const asset of assets) {
      await writeFile(asset.absPath, asset.data);
      written.push({ file: asset.absPath, source: asset.source });
    }
  }

  return {
    format: doc.format,
    fileName: options.fileName,
    markdown,
    doc,
    meta: doc.meta,
    assets: written,
    warnings,
  };
}

/** 转换文件。 */
export async function convertFile(filePath: string, options: ConvertOptions = {}): Promise<ConvertResult> {
  let data: Uint8Array;
  try {
    data = new Uint8Array(await readFile(filePath));
  } catch (err) {
    throw new Doc2mdError(`无法读取文件 ${filePath}：${(err as Error).message}`, 'FILE_NOT_FOUND');
  }
  return convert(data, { ...options, fileName: options.fileName ?? basename(filePath) });
}

function toBytes(data: Uint8Array | ArrayBuffer | string): Uint8Array {
  if (typeof data === 'string') return new TextEncoder().encode(data);
  if (data instanceof Uint8Array) return data;
  return new Uint8Array(data);
}

/** 纯 Markdown 文件透传。 */
function markdownAsDoc(text: string): Doc {
  return {
    format: 'markdown',
    meta: {},
    blocks: [{ type: 'preformatted', value: text }],
  };
}

/** 工具：解析输出路径，缺省为输入同目录 .md。 */
export function resolveOutputPath(userPath: string | undefined, fileName: string): string {
  if (userPath) return resolve(userPath);
  return resolve(fileName.replace(/\.[^.]+$/, '') + '.md');
}
