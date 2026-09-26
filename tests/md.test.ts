import { describe, expect, it } from 'vitest';
import type { Doc } from '../src/types.js';
import { renderMarkdown } from '../src/md/writer.js';

const doc: Doc = {
  format: 'pptx',
  meta: { title: '演示' },
  blocks: [
    { type: 'heading', level: 1, children: [{ kind: 'text', value: '第一章 概述' }] },
    {
      type: 'paragraph',
      children: [
        { kind: 'text', value: '这是正文，包含 ' },
        { kind: 'code', value: 'const x = 1' },
        { kind: 'text', value: ' 与 ' },
        { kind: 'strong', children: [{ kind: 'text', value: '加粗' }] },
        { kind: 'text', value: ' 和 ' },
        { kind: 'em', children: [{ kind: 'text', value: '斜体' }] },
      ],
    },
    {
      type: 'list',
      ordered: false,
      start: 1,
      items: [
        [{ type: 'paragraph', children: [{ kind: 'text', value: '第一项' }] }],
        [
          { type: 'paragraph', children: [{ kind: 'text', value: '第二项' }] },
          {
            type: 'list',
            ordered: true,
            start: 1,
            items: [[{ type: 'paragraph', children: [{ kind: 'text', value: '子项 1' }] }]],
          },
        ],
      ],
    },
    { type: 'table', header: ['名称', '数量'], rows: [['键盘', '12'], ['鼠标 | 无线', '3'], ['显示器', '1']] },
    { type: 'code', lang: 'ts', value: 'interface X {}' },
    { type: 'quote', children: [{ type: 'paragraph', children: [{ kind: 'text', value: '引用语' }] }] },
    { type: 'hr' },
    {
      type: 'section',
      title: '第二节',
      meta: { 页码: '2' },
      blocks: [{ type: 'paragraph', children: [{ kind: 'text', value: '分段内容' }] }],
    },
    { type: 'footnotes', title: '脚注', items: [{ label: '1', blocks: [{ type: 'paragraph', children: [{ kind: 'text', value: '说明文字' }] }] }] },
  ],
};

describe('renderMarkdown', () => {
  it('渲染基础结构', () => {
    const md = renderMarkdown(doc);
    expect(md).toContain('# 第一章 概述');
    expect(md).toContain('这是正文，包含 `const x = 1` 与 **加粗** 和 *斜体*');
  });

  it('嵌套列表缩进', () => {
    const md = renderMarkdown(doc);
    expect(md).toContain('- 第一项');
    expect(md).toContain('- 第二项');
    expect(md).toContain('  1. 子项 1\n');
  });

  it('表格转义管道符', () => {
    const md = renderMarkdown(doc);
    expect(md).toContain('| 鼠标 \\| 无线 | 3 |');
  });

  it('代码块、引用、hr、节、脚注', () => {
    const md = renderMarkdown(doc);
    expect(md).toContain('```ts\ninterface X {}\n```');
    expect(md).toContain('> 引用语');
    expect(md).toContain('---');
    expect(md).toContain('## 第二节');
    expect(md).toContain('<sup>页码: 2</sup>');
    expect(md).toContain('## 脚注');
    expect(md).toContain('[^1]: 说明文字');
  });

  it('行首 # 正文被转义避免变标题', () => {
    const md = renderMarkdown({
      format: 'docx',
      meta: {},
      blocks: [{ type: 'paragraph', children: [{ kind: 'text', value: '# 这是正文，不是标题' }] }],
    });
    expect(md).toContain('\\# 这是正文');
  });

  it('行首数字+点被转义避免开启有序列表', () => {
    const md = renderMarkdown({
      format: 'docx',
      meta: {},
      blocks: [{ type: 'paragraph', children: [{ kind: 'text', value: '3. 这个版本说明' }] }],
    });
    expect(md).toContain('3\\. 这个版本说明');
  });

  it('escape:false 保留原样', () => {
    const md = renderMarkdown(
      { format: 'docx', meta: {}, blocks: [{ type: 'paragraph', children: [{ kind: 'text', value: '| 未转义' }] }] },
      { escape: false },
    );
    expect(md).toContain('| 未转义');
  });

  it('headingOffset 偏移层级', () => {
    const md = renderMarkdown(doc, { headingOffset: 1 });
    expect(md).not.toMatch(/^# 第一章/m);
    expect(md).toContain('## 第一章 概述');
  });

  it('空段落默认不输出，keepEmptyParagraphs 时输出空行', () => {
    const emptyDoc: Doc = {
      format: 'docx',
      meta: {},
      blocks: [{ type: 'paragraph', children: [] }],
    };
    expect(renderMarkdown(emptyDoc).trim()).toBe('');
    expect(renderMarkdown(emptyDoc, { keepEmptyParagraphs: true }).trim()).toBe('');
  });
});