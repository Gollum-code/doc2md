import { attr, extractText } from './ooxml.js';

/** fast-xml-parser 输出节点（结构松散，这里保持宽松类型）。 */
export type XmlNode = any;

/** 读取子节点，兼容有无命名空间前缀。 */
export function child(node: XmlNode, tag: string): XmlNode | undefined {
  if (!node || typeof node !== 'object') return undefined;
  if (node[tag] !== undefined) return node[tag];
  const local = tag.includes(':') ? tag.slice(tag.indexOf(':') + 1) : tag;
  if (node[local] !== undefined) return node[local];
  // 前缀不匹配时按后缀匹配
  for (const key of Object.keys(node)) {
    if (key.startsWith('@') || key === '#text') continue;
    if (key.endsWith(':' + local)) return node[key];
  }
  return undefined;
}

/** 读取子节点数组。 */
export function children(node: XmlNode, tag: string): XmlNode[] {
  const v = child(node, tag);
  if (v === undefined || v === null) return [];
  return Array.isArray(v) ? v : [v];
}

/** 读取子节点纯文本。 */
export function childText(node: XmlNode, tag: string): string | undefined {
  const v = child(node, tag);
  if (v === undefined || v === null) return undefined;
  const t = extractText(v);
  return t === undefined ? undefined : t;
}

/** 读取自身纯文本（`<a:t>text</a:t>`）。 */
export function textOf(node: XmlNode): string {
  const t = extractText(node);
  return t === undefined ? '' : t;
}

/** 读取属性为数字。 */
export function numAttr(node: XmlNode, name: string, dflt = 0): number {
  const raw = attr(node, name);
  if (raw === undefined) return dflt;
  const n = Number(raw);
  return Number.isFinite(n) ? n : dflt;
}

/** 读取属性为布尔（OOXML 里 "1"/"0"/"true"/"false"）。 */
export function boolAttr(node: XmlNode, name: string, dflt = false): boolean {
  const raw = attr(node, name);
  if (raw === undefined) return dflt;
  return raw === '1' || raw === 'true' || raw === 'on';
}

/** 读取属性为字符串。 */
export function strAttr(node: XmlNode, name: string): string | undefined {
  return attr(node, name);
}
