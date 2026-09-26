import { describe, expect, it } from 'vitest';
import { normalizeMarkdown, normalizeText } from '../src/normalize.js';

describe('normalizeMarkdown', () => {
  it('CRLF → LF 且去掉多余空行', () => {
    const out = normalizeMarkdown('第一段\r\n\r\n\r\n\r\n第二段');
    expect(out).toBe('第一段\n\n第二段\n');
  });

  it('中英之间补空格', () => {
    expect(normalizeMarkdown('使用Node.js开发RAG应用').trim()).toBe('使用 Node.js 开发 RAG 应用');
  });

  it('数字与中文之间补空格', () => {
    expect(normalizeMarkdown('共3个业务线，转化率92%').trim()).toBe('共 3 个业务线，转化率 92%');
  });

  it('全角字母数字转半角', () => {
    expect(normalizeMarkdown('ＡＰＩ ｖ２')).toContain('API v2');
  });

  it('默认保留中文标点', () => {
    const out = normalizeMarkdown('他说：“好。”然后离开了。');
    expect(out).toContain('：');
    expect(out).toContain('。');
  });

  it('开启 punctuationToHalfWidth 后中文标点转半角', () => {
    const out = normalizeMarkdown('他说：“好，好。”', { punctuationToHalfWidth: true });
    expect(out).toContain('"');
    expect(out).toContain(',');
  });

  it('清理零宽字符与 NBSP', () => {
    const out = normalizeMarkdown('零\u200b宽\u00a0字符');
    expect(out.trim()).toBe('零宽 字符');
  });

  it('全角空格转半角并去行尾空格', () => {
    const out = normalizeMarkdown('\u3000缩进段落   ');
    expect(out).toBe('缩进段落\n');
  });

  it('项目符号归一化', () => {
    expect(normalizeMarkdown('• 第一点\n· 第二点').trim()).toBe('- 第一点\n- 第二点');
  });

  it('中文序号归一化为有序列表', () => {
    expect(normalizeMarkdown('1、第一步\n2、第二步').trim()).toBe('1. 第一步\n2. 第二步');
  });

  it('代码块内不做中英空格与标点处理', () => {
    const src = ['```js', 'const a = 1;', '```'].join('\n');
    expect(normalizeMarkdown(src).trim()).toBe(src);
  });

  it('表格行不被中英间距规则污染', () => {
    const src = ['| 名称 | 数量 |', '| --- | --- |', '| 键盘 | 12 |'].join('\n');
    expect(normalizeMarkdown(src).trim()).toBe(src);
  });

  it('不会把普通段落转成列表', () => {
    expect(normalizeMarkdown('2026.09.26 发布').trim()).toBe('2026.09.26 发布');
  });

  it('可用开关关闭全部规范化', () => {
    const src = 'ＡＢＣ中文abc';
    const out = normalizeMarkdown(src, { cjkLatinSpacing: false, fullWidthToHalfWidth: false });
    expect(out.trim()).toBe(src);
  });
});

describe('normalizeText', () => {
  it('单段文本压缩空白并规范中英混排', () => {
    expect(normalizeText('  转换   PPTX 为   Markdown  ')).toBe('转换 PPTX 为 Markdown');
  });

  it('空字符串安全', () => {
    expect(normalizeText('')).toBe('');
  });
});
