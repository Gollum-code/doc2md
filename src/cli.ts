#!/usr/bin/env node

import { cac } from 'cac';
import { mkdir, readdir, readFile, stat, writeFile } from 'node:fs/promises';
import { basename, dirname, extname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import type { ConvertOptions, DocFormat } from './types.js';
import { convertFile } from './index.js';

async function versionFromPkg(): Promise<string> {
  try {
    const pkgPath = join(dirname(fileURLToPath(import.meta.url)), '..', 'package.json');
    const pkg = JSON.parse(await readFile(pkgPath, 'utf8')) as { version?: string };
    return pkg.version ?? '0.0.0';
  } catch {
    return '0.0.0';
  }
}

export async function run(argv: string[]): Promise<number> {
  const cli = cac('doc2md');

  cli
    .command('[input...]', '转换文件或目录（.pptx/.docx/.xlsx/.pdf/.md）')
    .option('-o, --output <path>', '输出文件（单文件模式）')
    .option('-d, --out-dir <dir>', '输出目录（多文件/目录模式）')
    .option('--format <fmt>', '强制指定输入格式 pptx|docx|xlsx|pdf|markdown')
    .option('-j, --json', '输出 JSON（含 markdown 与元数据）')
    .option('--stdout', '只输出 markdown 到 stdout')
    .option('--assets', '导出图片', { default: true })
    .option('--assets-dir <dir>', '图片输出目录')
    .option('--image-base-url <url>', '图片引用前缀（默认相对 assets 目录名）')
    .option('--notes', '导出 PPT 演讲者备注', { default: true })
    .option('--hidden-slides', '包含隐藏幻灯片')
    .option('--front-matter', '输出 YAML front-matter')
    .option('--heading-offset <n>', '标题层级整体偏移', { default: 0 })
    .option('--table-style <style>', '表格样式 gfm|html|none', { default: 'gfm' })
    .option('--max-rows <n>', 'XLSX 每 sheet 最多行数', { default: 2000 })
    .option('--normalize', '中文/标点规范化', { default: true })
    .option('--half-punct', '中文标点也转半角')
    .option('--escape', 'Markdown 转义', { default: true })
    .option('--verbose', '打印详细日志');

  cli.help();
  // cac 内部按 process.argv 形态解析（argv[0]=node, argv[1]=script），这里补全前缀
  cli.parse(['doc2md', 'doc2md', ...argv], { run: false });

  if (cli.options.version) {
    console.log(await versionFromPkg());
    return 0;
  }

  const inputs = (cli.args ?? []) as string[];
  if (cli.options.help || !inputs.length) {
    cli.outputHelp();
    return 0;
  }

  const opts = cli.options as Record<string, unknown>;
  if (opts.output && inputs.length > 1) {
    console.error('✗ -o/--output 只能用于单个输入文件');
    return 1;
  }

  const convertOptions: ConvertOptions = {
    format: (opts.format as DocFormat) || undefined,
    images: opts.assets !== false,
    assetsDir: opts.assetsDir as string | undefined,
    imageBaseUrl: opts.imageBaseUrl as string | undefined,
    notes: opts.notes !== false,
    hiddenSlides: !!opts.hiddenSlides,
    frontMatter: !!opts.frontMatter,
    headingOffset: toInt(opts.headingOffset, 0),
    tableStyle: (opts.tableStyle as 'gfm' | 'html' | 'none') || 'gfm',
    maxRows: toInt(opts.maxRows, 2000),
    normalize: opts.normalize !== false,
    escape: opts.escape !== false,
    normalizeOptions: {
      punctuationToHalfWidth: !!opts.halfPunct,
    },
  };

  const jsonMode = !!opts.json;
  const stdoutMode = !!opts.stdout;
  const outDir = opts.outDir ? resolve(String(opts.outDir)) : undefined;

  const files = await expandInputs(inputs);
  if (!files.length) {
    console.error('✗ 未找到可转换的文件');
    return 1;
  }

  let exitCode = 0;
  for (const file of files) {
    try {
      const outputPath = resolveOutput(file, opts);
      const targetDir = outDir ?? dirname(outputPath);
      const result = await convertFile(file, {
        ...convertOptions,
        fileName: basename(file),
        outputDir: targetDir,
      });

      if (jsonMode) {
        console.log(
          JSON.stringify(
            {
              file,
              format: result.format,
              meta: result.meta,
              markdown: result.markdown,
              assets: result.assets.map((a) => ({ file: a.file, source: a.source })),
              warnings: result.warnings,
            },
            null,
            2,
          ),
        );
      } else if (stdoutMode) {
        process.stdout.write(result.markdown);
      } else {
        await mkdir(targetDir, { recursive: true });
        await writeFile(outputPath, result.markdown, 'utf8');
        const extra = result.assets.length ? `（含 ${result.assets.length} 张图）` : '';
        console.log(`✓ ${file} → ${outputPath}${extra}`);
      }

      if (opts.verbose && result.warnings.length) {
        for (const w of result.warnings) console.warn(`  ⚠ ${w}`);
      }
    } catch (err) {
      exitCode = 1;
      console.error(`✗ ${file}: ${(err as Error).message}`);
    }
  }
  return exitCode;
}

function resolveOutput(file: string, opts: Record<string, unknown>): string {
  const explicitOut = opts.output as string | undefined;
  const outDir = opts.outDir as string | undefined;
  if (explicitOut) return resolve(explicitOut);
  const base = file.replace(/\.[^.]+$/, '') + '.md';
  return outDir ? join(resolve(outDir), basename(base)) : base;
}

async function expandInputs(inputs: string[]): Promise<string[]> {
  const out: string[] = [];
  const SUPPORTED = new Set(['.pptx', '.docx', '.xlsx', '.pdf', '.md', '.markdown', '.txt']);
  for (const input of inputs) {
    const p = resolve(input);
    let st;
    try {
      st = await stat(p);
    } catch {
      console.error(`✗ 路径不存在：${p}`);
      continue;
    }
    if (st.isDirectory()) {
      for (const c of await listDirRecursive(p)) {
        if (SUPPORTED.has(extname(c).toLowerCase())) out.push(c);
      }
    } else {
      out.push(p);
    }
  }
  return out;
}

async function listDirRecursive(dir: string): Promise<string[]> {
  const out: string[] = [];
  for (const e of await readdir(dir, { withFileTypes: true })) {
    const p = join(dir, e.name);
    if (e.isDirectory()) out.push(...(await listDirRecursive(p)));
    else out.push(p);
  }
  return out;
}

function toInt(v: unknown, dflt: number): number {
  const n = Number(v);
  return Number.isFinite(n) ? n : dflt;
}

function isMain(entry: string | undefined): boolean {
  if (!entry) return false;
  try {
    return import.meta.url === pathToFileURL(entry).href;
  } catch {
    return false;
  }
}

if (isMain(process.argv[1])) {
  run(process.argv.slice(2)).then((code) => {
    process.exitCode = code;
  });
}
