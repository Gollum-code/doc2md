/**
 * doc2md 中间表示（IR）。
 *
 * 所有格式解析器（pptx / docx / xlsx / pdf）都输出同一棵 `Doc` 树，
 * 由渲染器统一转成 Markdown。这样新增格式只需实现一个 parser。
 */

/** 段落/文本对齐方式。 */
export type Align = 'left' | 'center' | 'right';

/** 文本内联片段。 */
export type Inline =
  | { kind: 'text'; value: string }
  | { kind: 'code'; value: string }
  | { kind: 'strong'; children: Inline[] }
  | { kind: 'em'; children: Inline[] }
  | { kind: 'del'; children: Inline[] }
  | { kind: 'sup'; children: Inline[] }
  | { kind: 'sub'; children: Inline[] }
  | { kind: 'link'; href: string; children: Inline[] }
  | { kind: 'image'; src: string; alt?: string }
  | { kind: 'break' }
  | { kind: 'pageBreak' };

/** 块级节点。 */
export type Block =
  /** 标题，level 1~6。 */
  | { type: 'heading'; level: number; children: Inline[] }
  /** 段落。 */
  | { type: 'paragraph'; children: Inline[]; align?: Align }
  /** 引用块。 */
  | { type: 'quote'; children: Block[] }
  /** 列表。 */
  | { type: 'list'; ordered: boolean; start: number; items: Block[][] }
  /** 表格。header 为空表示无表头。 */
  | { type: 'table'; header: string[]; rows: string[][]; align?: Align[] }
  /** 代码块。 */
  | { type: 'code'; value: string; lang?: string }
  /** 图片（独占一块）。 */
  | { type: 'image'; src: string; alt?: string }
  /** 水平分隔线。 */
  | { type: 'hr' }
  /** 小节标题（如 PPT 的一页、Excel 的一个 sheet）。level 1~6，默认 2。 */
  | { type: 'section'; title: string; level?: number; blocks: Block[]; meta?: Record<string, string> }
  /** 脚注 / 尾注容器。 */
  | { type: 'footnotes'; title: string; items: { label: string; blocks: Block[] }[] }
  /** 原始文本块（页脚、批注等低置信度内容）。 */
  | { type: 'preformatted'; value: string };

/** 文档级元数据。 */
export interface DocMeta {
  title?: string;
  author?: string;
  subject?: string;
  description?: string;
  keywords?: string[];
  created?: string;
  modified?: string;
  lastModifiedBy?: string;
  revision?: string;
  company?: string;
  /** 幻灯片总数（pptx）。 */
  slides?: number;
  /** 工作表名（xlsx）。 */
  sheets?: string[];
  [key: string]: unknown;
}

/** 一篇完整文档。 */
export interface Doc {
  /** 源格式。 */
  format: DocFormat;
  meta: DocMeta;
  blocks: Block[];
}

export type DocFormat = 'pptx' | 'docx' | 'xlsx' | 'pdf' | 'markdown' | 'unknown';

/** 转换选项。 */
export interface ConvertOptions {
  /** 显式指定源格式，缺省按扩展名 / 魔数嗅探。 */
  format?: DocFormat;
  /** 源文件名，仅用于日志与元数据。 */
  fileName?: string;
  /** 图片输出目录；不传则默认写到 md 文件同级的 `<basename>.assets/`。 */
  assetsDir?: string;
  /** 是否导出图片（默认 true）。 */
  images?: boolean;
  /** 图片引用相对 md 文件的路径前缀（默认相对 assets 目录名）。 */
  imageBaseUrl?: string;
  /** 是否输出 PPT 演讲者备注（默认 true）。 */
  notes?: boolean;
  /** 是否包含隐藏幻灯片（pptx，默认 false）。 */
  hiddenSlides?: boolean;
  /** 是否输出 front-matter（默认 false）。 */
  frontMatter?: boolean;
  /** 标题层级整体偏移，RAG 拼接多篇文档时用（默认 0）。 */
  headingOffset?: number;
  /** 中文/标点规范化开关（默认 true）。 */
  normalize?: boolean;
  /** 细粒度规范化选项。 */
  normalizeOptions?: NormalizeOptions;
  /** Markdown 转义开关（默认 true）。 */
  escape?: boolean;
  /** 表格输出方式：`gfm` 管道表格 / `html` / `none`（默认 gfm）。 */
  tableStyle?: 'gfm' | 'html' | 'none';
  /** XLSX 每个 sheet 最多输出多少行（默认 2000）。 */
  maxRows?: number;
  /** XLSX 是否输出空行分隔的 sheet 名（默认 true）。 */
  xlsxSheetTitles?: boolean;
  /** PPT 幻灯片之间是否插入 `---`（默认 true）。 */
  slideSeparator?: boolean;
  /** 幻灯片标题是否带序号（默认 true）。 */
  slideNumbers?: boolean;
  /** 空段落是否保留（默认 false）。 */
  keepEmptyParagraphs?: boolean;
  /** 表格空单元格占位符（默认空字符串）。 */
  emptyCell?: string;
  /** 缩进层级起始（DOCX 列表，保留原始层级）。 */
  listIndent?: 'flat' | 'nested';
  /** 输出 markdown 文件所在目录（用于解析图片落盘位置，CLI 内部使用）。 */
  outputDir?: string;
}

/** 规范化选项。 */
export interface NormalizeOptions {
  /** 中英文之间自动补空格（默认 true）。 */
  cjkLatinSpacing?: boolean;
  /** 全角 ASCII 转半角（默认 true）。 */
  fullWidthToHalfWidth?: boolean;
  /** 中文标点转英文标点（默认 false，保持中文排版）。 */
  punctuationToHalfWidth?: boolean;
  /** 连续空行压缩为一个（默认 true）。 */
  collapseBlankLines?: boolean;
  /** 项目符号归一化：`•`/`·` → `-`（默认 true）。 */
  normalizeBullets?: boolean;
  /** `1)` `1、` `1．` → `1.`（默认 true）。 */
  normalizeOrderedMarkers?: boolean;
  /** 全角空格转半角（默认 true）。 */
  ideographicSpace?: boolean;
  /** 是否清理零宽字符与 NBSP（默认 true）。 */
  stripInvisible?: boolean;
}

/** 转换结果。 */
export interface ConvertResult {
  /** 源格式。 */
  format: DocFormat;
  /** 源文件名。 */
  fileName?: string;
  /** 渲染出的 Markdown。 */
  markdown: string;
  /** 中间表示，便于二次处理 / 分块。 */
  doc: Doc;
  /** 元数据。 */
  meta: DocMeta;
  /** 导出的图片文件。 */
  assets: { file: string; source: string }[];
  /** 非致命告警。 */
  warnings: string[];
}

/** 解析上下文：包读取器 + 选项 + 资产收集。 */
export interface ParseContext {
  options: Required<
    Pick<
      ConvertOptions,
      'images' | 'notes' | 'hiddenSlides' | 'keepEmptyParagraphs' | 'maxRows' | 'xlsxSheetTitles' | 'emptyCell'
    >
  > &
    ConvertOptions & { fileName?: string };
  /** 写出图片，返回可写入 markdown 的路径。 */
  emitAsset(name: string, data: Uint8Array): string;
  /** 记录告警。 */
  warn(message: string): void;
}
