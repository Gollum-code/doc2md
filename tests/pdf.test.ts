import { describe, expect, it } from 'vitest';
import { PDFDocument, StandardFonts, rgb } from 'pdf-lib';
import { detectColumns, lineToText, linesToTable, orderTextItems, type PdfTextItem } from '../src/pdf/parser.js';

function item(str: string, x: number, y: number, w = 10, h = 12): PdfTextItem {
  return { str, x, y, w, h };
}

describe('pdf 阅读顺序（纯函数）', () => {
  it('按行聚类、行内按 x 排序', () => {
    const items = [
      item('右', 120, 720),
      item('左', 40, 720),
      item('第二行', 40, 700),
    ];
    const lines = orderTextItems(items);
    expect(lines).toHaveLength(2);
    expect(lines[0].map((i) => i.str)).toEqual(['左', '右']); // 同一行按 x
    expect(lines[1][0].str).toBe('第二行');
  });

  it('行拼接保留词间距', () => {
    const line = [item('docs', 40, 720, 30), item('to', 76, 720, 14), item('md', 96, 720, 16)];
    expect(lineToText(line)).toBe('docs to md');
  });

  it('紧密排列不插入空格', () => {
    const line = [item('doc', 40, 720, 15), item('2md', 56, 720, 15)];
    expect(lineToText(line)).toBe('doc2md');
  });
});

describe('pdf 表格启发式', () => {
  it('检测两列对齐', () => {
    const lines = [
      [item('Name', 40, 720), item('Score', 200, 720)],
      [item('Alice', 40, 700), item('92', 200, 700)],
      [item('Bob', 40, 680), item('88', 200, 680)],
    ];
    const cols = detectColumns(lines, { minSupport: 2 });
    expect(cols).toBeDefined();
    expect(cols!.x).toEqual([40, 200]);
  });

  it('单列文本不误判为表格', () => {
    const lines = [
      [item('一行标题说明', 40, 720)],
      [item('第二段正文', 40, 700)],
      [item('第三段正文', 40, 680)],
    ];
    expect(detectColumns(lines)).toBeUndefined();
  });

  it('linesToTable 组装网格', () => {
    const lines = [
      [item('Name', 40, 720), item('Score', 200, 720)],
      [item('Alice', 40, 700), item('92', 200, 700)],
    ];
    const cols = detectColumns(lines)!;
    const grid = linesToTable(lines, cols);
    expect(grid).toEqual([
      ['Name', 'Score'],
      ['Alice', '92'],
    ]);
  });
});

describe('pdf 集成（pdf-lib 生成）', () => {
  async function makePdf(draw: (page: Awaited<ReturnType<PDFDocument['addPage']>>, font: Awaited<ReturnType<PDFDocument['embedStandardFont']>>) => void): Promise<Uint8Array> {
    const pdf = await PDFDocument.create();
    const font = await pdf.embedStandardFont(StandardFonts.Helvetica);
    const page = pdf.addPage([612, 792]);
    draw(page, font);
    return pdf.save();
  }

  it('抽取文本并按页输出', async () => {
    const bytes = await makePdf((page, font) => {
      page.drawText('Quarterly Report', { x: 60, y: 720, size: 16, font, color: rgb(0, 0, 0) });
      page.drawText('Revenue grows 35% YoY', { x: 60, y: 700, size: 11, font });
    });
    const { convert } = await import('../src/index.js');
    const result = await convert(bytes, { fileName: 'demo.pdf' });
    expect(result.format).toBe('pdf');
    expect(result.markdown).toContain('Quarterly Report');
    expect(result.markdown).toContain('Revenue grows 35% YoY');
    expect(result.markdown).toContain('第 1 页');
  });

  it('列对齐表格还原', async () => {
    const bytes = await makePdf((page, font) => {
      page.drawText('Month', { x: 60, y: 700, size: 11, font });
      page.drawText('Sales', { x: 250, y: 700, size: 11, font });
      page.drawText('Jan', { x: 60, y: 680, size: 11, font });
      page.drawText('120', { x: 250, y: 680, size: 11, font });
      page.drawText('Feb', { x: 60, y: 660, size: 11, font });
      page.drawText('190', { x: 250, y: 660, size: 11, font });
    });
    const { convert } = await import('../src/index.js');
    const result = await convert(bytes, { fileName: 'table.pdf' });
    expect(result.markdown).toContain('| Month | Sales |');
    expect(result.markdown).toMatch(/Jan\s*\|\s*120/);
  });
});