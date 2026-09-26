import type { Block } from '../types.js';

export interface ListItemInput {
  /** 层级（0 起）。 */
  level: number;
  ordered: boolean;
  /** 有序列表起始数字。 */
  start?: number;
  /** 条目内容。 */
  content: Block[];
}

type ListBlock = { type: 'list'; ordered: boolean; start: number; items: Block[][] };

/**
 * 嵌套列表构建器：把连续 `level` 分层的列表项组装成嵌套列表树。
 * 非列表段落出现时应调用 `flush()`，或将两个列表间的段落先送入正文。
 * 注意：打开新层级时不预先挂载；只在关闭时挂到父级末项，避免重复。
 */
export class ListTreeBuilder {
  private stack: ListBlock[] = [];

  constructor(private readonly sink: Block[]) {}

  /** 追加一个列表项；level 相对当前根。 */
  add(item: ListItemInput): void {
    const level = Math.max(0, Math.min(item.level, 8));
    // 收拢更深层级（挂回父级）
    while (this.stack.length > level + 1) this.closeTop();
    // 打开缺失层级
    while (this.stack.length < level + 1) {
      this.stack.push({ type: 'list', ordered: false, start: 1, items: [] });
    }
    const cur = this.stack[level];
    if (item.ordered) cur.ordered = true;
    if (item.ordered && cur.items.length === 0) cur.start = item.start ?? 1;
    cur.items.push(item.content);
  }

  /** 结束所有打开的列表并挂回正文。 */
  flush(): void {
    while (this.stack.length > 1) this.closeTop();
    while (this.stack.length) this.sink.push(this.stack.pop()!);
  }

  private closeTop(): void {
    const top = this.stack.pop()!;
    const parent = this.stack[this.stack.length - 1];
    if (!parent) return;
    const lastItem = parent.items[parent.items.length - 1];
    if (lastItem) lastItem.push(top);
    else parent.items.push([top]);
  }
}