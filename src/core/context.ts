import { basename, resolve } from 'node:path';
import type { ConvertOptions, ParseContext } from '../types.js';
import { extLower } from './ooxml.js';

export interface PendingAsset {
  /** 写进 markdown 的引用。 */
  ref: string;
  /** 绝对落盘路径。 */
  absPath: string;
  /** 源部件路径。 */
  source: string;
  data: Uint8Array;
}

export interface ParseContextBundle {
  ctx: ParseContext;
  pending: PendingAsset[];
  warnings: string[];
}

/**
 * 解析上下文工厂：
 * - 决定图片导出目录与 markdown 引用前缀
 * - 解析器通过 emitAsset 收集待写图片，转换完成后统一落盘
 */
export function createParseContext(options: ConvertOptions, mdFileDir?: string): ParseContextBundle {
  const fileName = options.fileName;
  const base = fileName ? basename(fileName).replace(/\.[^.]*$/, '') : 'document';
  const defaultRelDir = `${base}.assets`;
  // 优先使用 outputDir（md 输出目录），其次用调用方传入的 mdFileDir
  const mdDir = options.outputDir ?? mdFileDir;

  let absDir: string;
  let prefix: string;

  if (options.assetsDir) {
    // 显式 assetsDir：若为绝对路径，引用前缀取 basename；否则相对 cwd 落盘
    const isAbs = /^[a-zA-Z]:[\\/]|^\//.test(options.assetsDir);
    absDir = isAbs ? options.assetsDir : resolve(options.assetsDir);
    prefix = options.imageBaseUrl ?? (isAbs ? basename(options.assetsDir) : toPosix(options.assetsDir));
  } else {
    // 默认放到 md 同级的 <base>.assets
    const rel = toPosix(defaultRelDir);
    absDir = mdDir ? resolve(mdDir, rel) : resolve(rel);
    prefix = options.imageBaseUrl ?? rel;
  }

  const pending: PendingAsset[] = [];
  const warnings: string[] = [];
  const sourceRefs = new Map<string, string>();
  const usedNames = new Set<string>();

  const allocateName = (source: string): string => {
    const existing = sourceRefs.get(source);
    if (existing) return existing;
    const ext = extLower(source);
    const rawBase = basename(source);
    let stem = rawBase.replace(/\.[^.]+$/, '');
    // 清理危险字符与控制字符
    stem = stem
      .replace(/[\\/:*?"<>|\u0000-\u001f\u007f]/g, '_')
      .replace(/^\.+/, '')
      .slice(0, 80);
    if (!stem) stem = 'image';
    let name = ext ? `${stem}.${ext}` : stem;
    let i = 2;
    while (usedNames.has(name)) {
      name = ext ? `${stem}-${i}.${ext}` : `${stem}-${i}`;
      i += 1;
    }
    usedNames.add(name);
    return name;
  };

  const ctx: ParseContext = {
    options: {
      images: true,
      notes: true,
      hiddenSlides: false,
      keepEmptyParagraphs: false,
      maxRows: 2000,
      xlsxSheetTitles: true,
      emptyCell: '',
      fileName,
      ...options,
    },
    emitAsset(source: string, data: Uint8Array): string {
      if (!ctx.options.images) return '';
      const name = allocateName(source);
      sourceRefs.set(source, name);
      const ref = `${prefix}/${name}`;
      const absPath = resolve(absDir, name);
      pending.push({ ref, absPath, source, data });
      return ref;
    },
    warn(message: string): void {
      warnings.push(message);
    },
  };

  return { ctx, pending, warnings };
}

function toPosix(p: string): string {
  return p.replace(/\\/g, '/');
}

/** 图片类型由扩展名推断（仅用于文案，不影响内容）。 */
export function inferImageKind(ext: string): string | undefined {
  switch (ext.toLowerCase()) {
    case 'png':
    case 'jpg':
    case 'jpeg':
    case 'gif':
    case 'webp':
    case 'bmp':
    case 'svg':
      return 'image';
    case 'emf':
    case 'wmf':
      return 'vector';
    default:
      return undefined;
  }
}

/** 供 CLI 计算默认输出路径。 */
export function defaultOutputFor(filePath: string): string {
  return filePath.replace(/\.[^.]+$/, '') + '.md';
}
