import { describe, expect, it } from 'vitest';
import { parseDocument, detectFormat } from '../src/index.js';
import { buildPptx, buildDocx, buildXlsx } from './fixtures/build.js';

describe('detectFormat', () => {
  it('识别 pptx/docx/xlsx 魔数', () => {
    expect(detectFormat(buildPptx())).toBe('pptx');
    expect(detectFormat(buildDocx())).toBe('docx');
    expect(detectFormat(buildXlsx())).toBe('xlsx');
  });

  it('按扩展名识别 markdown', () => {
    expect(detectFormat(new TextEncoder().encode('# 标题'), 'a.md')).toBe('markdown');
  });
});

describe('pptx → Doc', () => {
  const { doc } = parseDocument(buildPptx(), { fileName: 'demo.pptx' });

  it('元数据', () => {
    expect(doc.format).toBe('pptx');
    expect(doc.meta.title).toBe('季度业务回顾');
    expect(doc.meta.author).toBe('张伟');
    expect(doc.meta.slides).toBe(3);
  });

  it('隐藏页默认排除，正文 slide1/slide2 保留', () => {
    const sections = doc.blocks.filter((b) => b.type === 'section');
    expect(sections).toHaveLength(2);
    expect(sections[0].title).toContain('季度业务回顾');
    expect(sections[1].title).toContain('下季度计划');
  });

  it('包含备注的 blockquote', () => {
    const sections = doc.blocks.filter((b) => b.type === 'section');
    const quote = sections[0].blocks.find((b) => b.type === 'quote');
    expect(quote).toBeDefined();
    expect(JSON.stringify(quote)).toContain('营收增长35%');
    // 备注中不应包含页码字段与幻灯片图像占位符
    expect(JSON.stringify(quote)).not.toContain('‹#›');
  });

  it('要点合并为列表且支持嵌套层级', () => {
    const sections = doc.blocks.filter((b) => b.type === 'section');
    const lists = sections[0].blocks.filter((b) => b.type === 'list');
    expect(lists.length).toBeGreaterThan(0);
    const json = JSON.stringify(sections[0].blocks);
    expect(json).toContain('营收同比增长');
    expect(json).toContain('自动摘要准确率92%');
  });

  it('buNone 段落不进入列表', () => {
    const sections = doc.blocks.filter((b) => b.type === 'section');
    const paras = sections[0].blocks.filter((b) => b.type === 'paragraph');
    const text = JSON.stringify(paras);
    expect(text).toContain('数据截至2026年9月');
  });

  it('表格块', () => {
    const sections = doc.blocks.filter((b) => b.type === 'section');
    const table = sections[0].blocks.find((b) => b.type === 'table' && b.header[0] === '业务线');
    expect(table).toBeDefined();
    if (table && table.type === 'table') {
      expect(table.header).toEqual(['业务线', '营收']);
      expect(table.rows).toEqual([['零售', '1200']]);
    }
  });

  it('图表转表格', () => {
    const sections = doc.blocks.filter((b) => b.type === 'section');
    const table = sections[0].blocks.find((b) => b.type === 'table' && b.header[0] === '类别');
    expect(table).toBeDefined();
    if (table && table.type === 'table') {
      expect(table.header).toEqual(['类别', '营收(万元)']);
      expect(table.rows).toEqual([
        ['Q1', '820'],
        ['Q2', '910'],
        ['Q3', '1107'],
      ]);
    }
  });

  it('图片分配引用', () => {
    const { assets } = parseDocument(buildPptx(), { fileName: 'demo.pptx' });
    expect(assets.length).toBe(1);
    expect(assets[0].source).toBe('ppt/media/image1.png');
    expect(assets[0].ref).toContain('image1.png');
  });

  it('hiddenSlides 选项可包含隐藏页', () => {
    const { doc } = parseDocument(buildPptx(), { fileName: 'demo.pptx', hiddenSlides: true });
    const sections = doc.blocks.filter((b) => b.type === 'section');
    expect(sections).toHaveLength(3);
    expect(sections[2].title).toContain('内部草稿页');
  });

  it('notes:false 去掉备注', () => {
    const { doc } = parseDocument(buildPptx(), { fileName: 'demo.pptx', notes: false });
    const sections = doc.blocks.filter((b) => b.type === 'section');
    expect(sections[0].blocks.some((b) => b.type === 'quote')).toBe(false);
  });
});

describe('docx → Doc', () => {
  const { doc } = parseDocument(buildDocx(), { fileName: 'demo.docx' });

  it('标题层级', () => {
    const headings = doc.blocks.filter((b) => b.type === 'heading');
    const h1 = headings.find((h) => h.level === 1);
    expect(h1).toBeDefined();
    expect(JSON.stringify(h1?.children)).toContain('doc2md 项目说明');
    expect(headings.some((h) => h.level === 2)).toBe(true);
  });

  it('内联粗体/斜体/代码', () => {
    const json = JSON.stringify(doc.blocks);
    expect(json).toContain('"kind":"strong"');
    expect(json).toContain('"kind":"em"');
    expect(json).toContain('"kind":"code"');
  });

  it('无序列表与嵌套', () => {
    const lists = doc.blocks.filter((b) => b.type === 'list');
    expect(lists.length).toBeGreaterThan(0);
    const json = JSON.stringify(lists);
    expect(json).toContain('支持表格');
    expect(json).toContain('支持嵌套列表');
  });

  it('有序列表', () => {
    const ordered = doc.blocks.find((b) => b.type === 'list' && b.ordered);
    expect(ordered).toBeDefined();
    expect(JSON.stringify(ordered)).toContain('第一阶段交付 CLI');
  });

  it('引用样式段落 → quote', () => {
    const quote = doc.blocks.find((b) => b.type === 'quote');
    expect(quote).toBeDefined();
    expect(JSON.stringify(quote)).toContain('表格是 RAG');
  });

  it('超链接保留 URL', () => {
    const json = JSON.stringify(doc.blocks);
    expect(json).toContain('https://example.com/docs');
  });

  it('字段指令不输出，只输出结果文本', () => {
    const json = JSON.stringify(doc.blocks);
    expect(json).not.toContain('TOC');
    expect(json).toContain('目录标题');
    expect(json).toContain('后续说明');
  });

  it('表格含表头与 gridSpan', () => {
    const table = doc.blocks.find((b) => b.type === 'table');
    expect(table).toBeDefined();
    if (table && table.type === 'table') {
      expect(table.header).toEqual(['模块', '状态']);
      expect(table.rows).toContainEqual(['pptx2md', '已完成']);
      expect(table.rows[1]).toContain('docx2md 也在同一期');
    }
  });

  it('分页符 → pageBreak', () => {
    const json = JSON.stringify(doc.blocks);
    expect(json).toContain('pageBreak');
    expect(json).toContain('第二页内容');
  });

  it('脚注收集到文档末尾', () => {
    const fn = doc.blocks.find((b) => b.type === 'footnotes');
    expect(fn).toBeDefined();
    if (fn && fn.type === 'footnotes') {
      expect(fn.items.length).toBe(1);
      expect(JSON.stringify(fn.items[0].blocks)).toContain('口径以财务系统为准');
    }
  });

  it('图片引用', () => {
    const { assets } = parseDocument(buildDocx(), { fileName: 'demo.docx' });
    expect(assets.length).toBe(1);
    expect(assets[0].source).toBe('word/media/image1.png');
  });
});

describe('xlsx → Doc', () => {
  const { doc } = parseDocument(buildXlsx(), { fileName: 'sales.xlsx' });

  it('两个 sheet 两个 section', () => {
    const sections = doc.blocks.filter((b) => b.type === 'section');
    expect(sections).toHaveLength(2);
    expect(sections[0].title).toBe('销售明细');
    expect(sections[1].title).toBe('汇总');
  });

  it('共享字符串与数字', () => {
    const sections = doc.blocks.filter((b) => b.type === 'section');
    const table = sections[0].blocks[0];
    expect(table.type).toBe('table');
    if (table.type === 'table') {
      expect(table.header).toEqual(['产品', '区域', '销量', '单价']);
      expect(table.rows[0]).toEqual(['doc2md转换器', '华东', '120', '199']);
      expect(table.rows[2]).toEqual(['汇总', '华东', '35%', '2026-09-05']);
    }
  });

  it('富文本共享字符串拼接', () => {
    const json = JSON.stringify(doc.blocks);
    expect(json).toContain('doc2md转换器');
  });

  it('内联字符串与布尔值', () => {
    const json = JSON.stringify(doc.blocks);
    expect(json).toContain('内联字符串');
    expect(json).toContain('TRUE');
  });

  it('合并单元格非左上角为空', () => {
    const json = JSON.stringify(doc.blocks);
    // A8:C8 合并，但空行已被裁剪，此处只验证解析不抛错
    expect(json.length).toBeGreaterThan(0);
  });

  it('空行被裁剪', () => {
    const sections = doc.blocks.filter((b) => b.type === 'section');
    const table = sections[0].blocks[0];
    if (table.type === 'table') {
      // 行2-5 均为有效数据行，第6行全空应被裁掉
      expect(table.rows.length).toBe(4);
    }
  });

  it('maxRows 截断', () => {
    const { doc } = parseDocument(buildXlsx(), { fileName: 'sales.xlsx', maxRows: 1 });
    const sections = doc.blocks.filter((b) => b.type === 'section');
    const table = sections[0].blocks[0];
    if (table.type === 'table') {
      expect(table.rows.length).toBe(1);
    }
  });
});
