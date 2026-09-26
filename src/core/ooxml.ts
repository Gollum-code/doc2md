import { unzipSync, type Unzipped } from 'fflate';
import { XMLParser } from 'fast-xml-parser';
import { join, posix } from 'node:path';

/** OOXML 关系。 */
export interface OoxmlRel {
  id: string;
  type: string;
  /** 相对 part 目录的 target；External 时为绝对 URL。 */
  target: string;
  /** 是否外部引用（URL 等）。 */
  external: boolean;
  /** 解析后的包内部件路径（internal 时）。 */
  targetPart?: string;
}

/** 关系类型常量。 */
export const REL = {
  officeDocument: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  slide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
  notesSlide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
  slideLayout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  chart: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
  worksheet: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
  styles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  sharedStrings: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
  numbering: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
  footnotes: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes',
  endnotes: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes',
  coreProperties: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  extendedProperties: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
  nextSlide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
} as const;

/** 一个打开的 OOXML 包（pptx/docx/xlsx 都是 zip）。 */
export class OoxmlPackage {
  static PARSER = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    parseTagValue: false,
    parseAttributeValue: false,
    trimValues: false,
    processEntities: true,
    preserveOrder: false,
  });

  /** 规范化的条目表：无前导 `/`，`/` 分隔。 */
  readonly entries: Map<string, Uint8Array>;
  private readonly cache = new Map<string, ReturnType<typeof OoxmlPackage.parseText>>();

  private constructor(zip: Unzipped) {
    this.entries = new Map();
    for (const raw of Object.keys(zip)) {
      const key = normalizeEntryName(raw);
      this.entries.set(key, zip[raw]);
    }
  }

  static async fromBuffer(data: Uint8Array): Promise<OoxmlPackage> {
    const zip = unzipSync(data);
    return new OoxmlPackage(zip);
  }

  static fromBufferSync(data: Uint8Array): OoxmlPackage {
    return new OoxmlPackage(unzipSync(data));
  }

  static parseText(text: string): unknown {
    return OoxmlPackage.PARSER.parse(text);
  }

  has(path: string): boolean {
    return this.entries.has(normalizeEntryName(path));
  }

  readBlob(path: string): Uint8Array | undefined {
    const entry = this.entries.get(normalizeEntryName(path));
    return entry;
  }

  readText(path: string): string | undefined {
    const blob = this.readBlob(path);
    if (!blob) return undefined;
    return decodeXmlText(blob);
  }

  /** 解析 XML 并返回根元素（fast-xml-parser 对象）。 */
  readXml(path: string): unknown {
    const key = `xml:${normalizeEntryName(path)}`;
    if (this.cache.has(key)) return this.cache.get(key);
    const raw = this.readText(path);
    if (raw === undefined) {
      throw new Error(`part not found: ${path}`);
    }
    const parsed = unwrapRoot(OoxmlPackage.parseText(raw));
    this.cache.set(key, parsed);
    return parsed;
  }

  /** 读取 part 的关系（`part/_rels/<name>.rels`）。 */
  readRels(partPath: string): OoxmlRel[] {
    const p = normalizeEntryName(partPath);
    const relPath = join(posixDirname(p), '_rels', `${posixBasename(p)}.rels`);
    const raw = this.readText(relPath);
    if (raw === undefined) return [];
    const root = OoxmlPackage.parseText(raw) as Record<string, unknown>;
    const list = (root as { Relationships?: { Relationship?: unknown } }).Relationships?.Relationship;
    const arr = list === undefined ? [] : Array.isArray(list) ? list : [list];
    const out: OoxmlRel[] = [];
    for (const item of arr) {
      const r = item as Record<string, string>;
      const id = r['@_Id'];
      const type = r['@_Type'] ?? '';
      const target = r['@_Target'] ?? '';
      const external = (r['@_TargetMode'] ?? '') === 'External';
      const targetPart = external ? undefined : resolveTarget(posixDirname(p), target);
      out.push({ id, type, target, external, targetPart });
    }
    return out;
  }

  /** 按 rel id 找关系。 */
  getRel(partPath: string, relId: string): OoxmlRel | undefined {
    return this.readRels(partPath).find((r) => r.id === relId);
  }

  /** 按 type 找关系。 */
  getRelsByType(partPath: string, type: string): OoxmlRel[] {
    return this.readRels(partPath).filter((r) => r.type === type);
  }

  /** 列出包内所有 media 部件。 */
  listMedia(): string[] {
    return [...this.entries.keys()].filter((k) => k.includes('/media/'));
  }

  /** 读 docProps/core.xml 与 docProps/app.xml 的元数据。 */
  readCoreProps(): Record<string, string | undefined> {
    const out: Record<string, string | undefined> = {};
    const core = this.entries.has('docProps/core.xml')
      ? (this.readXml('docProps/core.xml') as Record<string, unknown>)
      : undefined;
    // dc:title 等带前缀，剥前缀取最后一段作为 key
    const pick = (obj: unknown, keys: string[]): string | undefined => {
      if (!obj || typeof obj !== 'object') return undefined;
      const rec = obj as Record<string, unknown>;
      for (const k of keys) {
        const v = rec[k];
        if (v !== undefined) {
          const s = extractText(v);
          return s === undefined ? undefined : s.trim() || undefined;
        }
      }
      return undefined;
    };
    if (core) {
      out.title = pick(core, ['dc:title', 'title']);
      out.author = pick(core, ['dc:creator', 'creator']);
      out.subject = pick(core, ['dc:subject', 'subject']);
      out.description = pick(core, ['dc:description', 'description']);
      out.keywordsRaw = pick(core, ['cp:keywords', 'keywords']);
      out.created = pick(core, ['dcterms:created', 'created']);
      out.modified = pick(core, ['dcterms:modified', 'modified']);
      out.lastModifiedBy = pick(core, ['cp:lastModifiedBy', 'lastModifiedBy']);
      out.revision = pick(core, ['cp:revision', 'revision']);
    }
    const app = this.entries.has('docProps/app.xml')
      ? (this.readXml('docProps/app.xml') as Record<string, unknown>)
      : undefined;
    if (app) {
      out.company = pick(app, ['Company']);
    }
    return out;
  }
}

/**
 * fast-xml-parser 会把 XML 声明当作 `?xml` 键，导致真正的根元素被嵌套一层。
 * 这里把根元素提上来，供 `readXml` 返回。
 */
function unwrapRoot(parsed: unknown): unknown {
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) return parsed;
  const rec = parsed as Record<string, unknown>;
  const keys = Object.keys(rec);
  if (keys.includes('?xml')) {
    const others = keys.filter((k) => k !== '?xml');
    if (others.length === 1) return rec[others[0]];
  }
  return parsed;
}

/** 关系 target 解析成包的规范部件路径。 */
function resolveTarget(partDir: string, target: string): string {
  if (target.startsWith('/')) {
    return normalizeEntryName(target.slice(1));
  }
  return normalizeEntryName(join(partDir, target));
}

/** 规范化 zip 条目路径：去 `.`/`..`、去前导 `/`、统一 `/`。 */
export function normalizeEntryName(name: string): string {
  let n = name.replace(/\\/g, '/');
  while (n.startsWith('/')) n = n.slice(1);
  const parts: string[] = [];
  for (const seg of n.split('/')) {
    if (seg === '' || seg === '.') continue;
    if (seg === '..') {
      parts.pop();
      continue;
    }
    parts.push(seg);
  }
  return parts.join('/');
}

function posixDirname(p: string): string {
  return posix.dirname(p).replace(/\\/g, '/');
}

function posixBasename(p: string): string {
  return posix.basename(p);
}

/** 尝试多种编码解码 OOXML 文本部件。 */
export function decodeXmlText(data: Uint8Array): string {
  if (data.length >= 2) {
    if (data[0] === 0xff && data[1] === 0xfe) {
      return decodeUtf16(data, true);
    }
    if (data[0] === 0xfe && data[1] === 0xff) {
      return decodeUtf16(data, false);
    }
  }
  if (data.length >= 3 && data[0] === 0xef && data[1] === 0xbb && data[2] === 0xbf) {
    return new TextDecoder('utf-8').decode(data.subarray(3));
  }
  // 有 BOM 的 utf-8 之外，先按 utf-8 解析，替换非法序列而不是中断
  return new TextDecoder('utf-8', { fatal: false }).decode(data);
}

function decodeUtf16(data: Uint8Array, littleEndian: boolean): string {
  const words = new Uint16Array(data.buffer, data.byteOffset, data.byteLength >> 1);
  if (words.length > 0 && words[0] === 0xfeff) {
    // 跳过 BOM
    const rest = words.subarray(1);
    return String.fromCharCode(...rest);
  }
  // 无 BOM：直接按 endianness 读
  const out: string[] = [];
  for (let i = littleEndian ? 0 : 1; i + 1 < data.length; i += 2) {
    out.push(String.fromCharCode(littleEndian ? data[i] | (data[i + 1] << 8) : data[i + 1] | (data[i] << 8)));
  }
  return out.join('');
}

/** 从 fast-xml-parser 节点中抽取纯文本（递归所有后代 `#text`）。 */
export function extractText(node: unknown): string | undefined {
  if (node === null || node === undefined) return undefined;
  if (typeof node === 'string') return node;
  if (typeof node === 'number' || typeof node === 'boolean') return String(node);
  if (Array.isArray(node)) {
    let s = '';
    for (const item of node) {
      const t = extractText(item);
      if (t !== undefined) s += t;
    }
    return s;
  }
  const rec = node as Record<string, unknown>;
  if (typeof rec !== 'object') return undefined;
  let s = '';
  for (const key of Object.keys(rec)) {
    if (key.startsWith('@') || key === '#text') {
      if (key === '#text') s += extractText(rec[key]);
      continue;
    }
    const t = extractText(rec[key]);
    if (t !== undefined) s += t;
  }
  return s;
}

/** 读 fast-xml-parser 属性节点下带前缀的属性。 */
export function attr(node: unknown, name: string): string | undefined {
  if (!node || typeof node !== 'object') return undefined;
  const rec = node as Record<string, unknown>;
  for (const key of [name, `@_${name}`, `@_${toCamel(name)}`, name.replace(/^@_/, '')]) {
    const v = rec[key];
    if (v !== undefined) {
      return typeof v === 'string' || typeof v === 'number' ? String(v) : undefined;
    }
  }
  // 带命名空间前缀的属性，如 @_w:val
  const suffix = name.split(':').pop();
  for (const key of Object.keys(rec)) {
    if (key.startsWith('@_') && key.endsWith(`:${suffix}`)) {
      const v = rec[key];
      return typeof v === 'string' || typeof v === 'number' ? String(v) : undefined;
    }
  }
  return undefined;
}

function toCamel(s: string): string {
  return s.replace(/[-_]([a-z])/g, (_m, c) => c.toUpperCase());
}

/** 扩展名小写。 */
export function extLower(path: string): string {
  const i = path.lastIndexOf('.');
  return i < 0 ? '' : path.slice(i + 1).toLowerCase();
}