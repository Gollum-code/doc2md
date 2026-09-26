import { describe, expect, it } from 'vitest';
import { mkdtemp, mkdir, readFile, readdir, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { run } from '../src/cli.js';
import { buildDocx, buildPptx, buildXlsx } from './fixtures/build.js';

async function tempDir(): Promise<string> {
  return mkdtemp(join(tmpdir(), 'doc2md-'));
}

describe('cli', () => {
  it('转换单个 docx 并写出 md + 图片目录', async () => {
    const dir = await tempDir();
    const file = join(dir, 'demo.docx');
    await writeFile(file, buildDocx());

    const code = await run([file]);
    expect(code).toBe(0);

    const md = await readFile(join(dir, 'demo.md'), 'utf8');
    expect(md).toContain('# doc2md 项目说明');
    expect(md).toContain('**干净的 Markdown**');
    // 脚注引用与定义
    expect(md).toContain('[^1]');
    // 表格
    expect(md).toContain('| 模块 | 状态 |');

    const assets = await readdir(join(dir, 'demo.assets'));
    expect(assets).toContain('image1.png');

    await rm(dir, { recursive: true, force: true });
  });

  it('--stdout 只输出 markdown 不落盘', async () => {
    const dir = await tempDir();
    const file = join(dir, 'demo.pptx');
    await writeFile(file, buildPptx());

    const origWrite = process.stdout.write.bind(process.stdout);
    let captured = '';
    // 捕获 stdout
    (process.stdout as unknown as { write: (c: string) => boolean }).write = ((c: string) => {
      captured += c;
      return true;
    }) as never;
    const code = await run([file, '--stdout']);
    (process.stdout as unknown as { write: typeof origWrite }).write = origWrite;

    expect(code).toBe(0);
    expect(captured).toContain('# 季度业务回顾');
    expect(captured).toContain('演讲备注');
    // 不应落盘
    await expect(readFile(join(dir, 'demo.md'), 'utf8')).rejects.toThrow();

    await rm(dir, { recursive: true, force: true });
  });

  it('--json 输出结构化结果', async () => {
    const dir = await tempDir();
    const file = join(dir, 'demo.pptx');
    await writeFile(file, buildPptx());

    const origLog = console.log;
    const logs: string[] = [];
    console.log = (...args: unknown[]) => {
      logs.push(args.map(String).join(' '));
    };
    const code = await run([file, '--json']);
    console.log = origLog;

    expect(code).toBe(0);
    const parsed = JSON.parse(logs.join('\n'));
    expect(parsed.format).toBe('pptx');
    expect(parsed.meta.title).toBe('季度业务回顾');
    expect(parsed.markdown).toContain('季度业务回顾');
    expect(parsed.assets.length).toBe(1);

    await rm(dir, { recursive: true, force: true });
  });

  it('目录批量转换', async () => {
    const dir = await tempDir();
    const sub = join(dir, 'docs');
    await mkdir(sub, { recursive: true });
    await writeFile(join(dir, 'a.pptx'), buildPptx());
    await writeFile(join(sub, 'b.docx'), buildDocx());
    await writeFile(join(sub, 'c.xlsx'), buildXlsx());

    const code = await run([dir]);
    expect(code).toBe(0);
    const md = await readFile(join(sub, 'b.md'), 'utf8');
    expect(md).toContain('设计目标');
    const md2 = await readFile(join(sub, 'c.md'), 'utf8');
    expect(md2).toContain('销售明细');
    const md3 = await readFile(join(dir, 'a.md'), 'utf8');
    expect(md3).toContain('季度业务回顾');

    await rm(dir, { recursive: true, force: true });
  });

  it('不支持的格式返回非零退出码', async () => {
    const dir = await tempDir();
    const file = join(dir, 'x.bin');
    await writeFile(file, new Uint8Array([1, 2, 3]));
    const code = await run([file]);
    expect(code).toBe(1);
    await rm(dir, { recursive: true, force: true });
  });
});