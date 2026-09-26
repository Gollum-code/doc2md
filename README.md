# doc2md

> 把 **PPTX / DOCX / XLSX / PDF** 转成 LLM 能直接读的干净 Markdown —— **中文友好**、表格与阅读顺序优先，零配置即用。

[![npm version](https://img.shields.io/npm/v/@gollum-code/doc2md)](https://www.npmjs.com/package/@gollum-code/doc2md)
[![npm downloads](https://img.shields.io/npm/dm/@gollum-code/doc2md)](https://www.npmjs.com/package/@gollum-code/doc2md)
[![CI](https://img.shields.io/github/actions/workflow/status/Gollum-code/doc2md/ci.yml?branch=main&label=CI)](https://github.com/Gollum-code/doc2md/actions)
[![License](https://img.shields.io/npm/l/doc2md)](LICENSE)

```
$ doc2md 季度业务汇报.pptx
✓ 季度业务汇报.pptx → 季度业务汇报.md（含 12 张图）
```

RAG / AI 应用里「上传文档 → 喂给大模型」一直靠 Python 生态（markitdown），**Node / npm 没有一个高质量的多格式转 Markdown 库**。doc2md 补上了这个空白：解码 OOXML 结构、还原表格/列表/阅读顺序、导出图片，并做专门的中文排版规范化，输出对 LLM 越干净越好用的文本。

---

## ✨ 特性

| 格式 | 支持内容 |
|---|---|
| **PPTX** | 标题与要点层级、嵌套列表、**表格**、**图表转表格**、**演讲者备注**、图片导出、隐藏页开关、按位置重排阅读顺序 |
| **DOCX** | 标题结构、**有序/无序/嵌套列表**、表格（含合并单元格）、链接、**脚注**、内联加粗/斜体/代码、图片、域代码处理、分页符 |
| **XLSX** | 多工作表、共享字符串、富文本、**日期/百分比格式**、合并单元格、空行修剪 |
| **PDF** ⚠️ | 文本抽取（pdf.js）+ **阅读顺序重排** + 列对齐表格启发式；扫描件 OCR 与公式暂未支持 |
| **Markdown** | 透传 `.md` / `.txt` |

- **中文优化**：中英混排自动补空格、全角字母数字转半角、项目符号/中文序号归一化、零宽字符清理，且**不破坏代码块与表格**
- **信息密度优先**：表格保结构、图表直接展开成 Markdown 表格、备注可选保留
- **结构化输出**：`--json` 导出 IR，可直接做分块（chunking）
- **CLI 批量**：目录递归、`--stdout`、front-matter 输出、标题层级偏移

## 📦 安装

```bash
npm install -g @gollum-code/doc2md   # CLI
# 或作为库：
npm install @gollum-code/doc2md
```

要求 Node ≥ 18.17。

## 🚀 快速开始（CLI）

```bash
# 单文件，默认输出到同目录 .md
doc2md 报告.pptx

# 指定输出
doc2md 报告.docx -o docs/report.md

# 目录（含子目录）批量
doc2md ./知识库/

# 只输出到 stdout（管道到其它工具）
doc2md 报告.pptx --stdout --no-notes

# 输出 JSON（含结构化 IR，便于分块）
doc2md 报告.pptx --json

# 常见选项
doc2md 演示.pptx \
  --no-notes           # 去掉演讲者备注
  --hidden-slides      # 包含隐藏幻灯片
  --front-matter       # 输出 YAML front-matter
  --heading-offset 1   # 标题整体 +1 级（多文档合并成 RAG 语料时）
  --no-assets          # 不导出图片
```

完整选项：`doc2md --help`。

## 📖 作为库使用

```ts
import { convert, convertFile, parseDocument } from 'doc2md';

// Buffer / ArrayBuffer / string
const result = await convert(buffer, { fileName: '汇报.pptx' });
console.log(result.markdown);          // 渲染后的 Markdown
console.log(result.doc.blocks);        // 结构化 IR，二次分块
console.log(result.meta.title);        // 元数据
console.log(result.assets);            // 已导出的图片

// 只拿 IR，不写图片（自己控制落盘）
const { doc, assets } = parseDocument(buffer);
```

### ConvertOptions 摘要

| 选项 | 默认 | 说明 |
|---|---|---|
| `format` | 嗅探 | 强制 `pptx / docx / xlsx / markdown` |
| `assetsDir` / `imageBaseUrl` | `<name>.assets` | 图片导出目录 / 引用前缀 |
| `images` | `true` | 是否导出图片 |
| `notes` | `true` | 是否保留 PPT 演讲者备注 |
| `hiddenSlides` | `false` | 是否包含隐藏幻灯片 |
| `frontMatter` | `false` | 输出 YAML front-matter |
| `headingOffset` | `0` | 标题层级偏移（合并语料用） |
| `tableStyle` | `'gfm'` | `gfm / html / none` |
| `normalize` + `normalizeOptions` | `true` | 中文规范化管线开关与细节 |
| `keepEmptyParagraphs` | `false` | 保留空段落 |
| `maxRows` | `2000` | XLSX 每 sheet 上限行数 |

### IR 一览

```
Doc { format, meta, blocks: Block[] }
Block = heading | paragraph | list | table | code | quote | image | hr
      | section（PPT 页 / XLSX sheet） | footnotes | preformatted
Inline = text | code | strong | em | del | sup | sub | link | image | break
```

## 🧭 输入示例 → 输出示例

**输入**：一个带标题、要点、表格、演讲者备注的 PPT
**输出**：

```markdown
# 季度业务回顾

<sup>页码: 1</sup>
## 季度业务回顾
- **营收同比增长** 35%
- AI 助手接入 3 个业务线
  - 自动摘要准确率 92%

![assets/image1.png](assets/image1.png)

| 业务线 | 营收 |
| --- | --- |
| 零售 | 1200 |

> **演讲备注**：本页重点强调营收增长 35%，并说明 AI 助手的落地情况。
```

## 🗺 Roadmap

| 阶段 | 内容 | 状态 |
|---|---|---|
| M1 | pptx2md + CLI + 表格 | ✅ 已实现 |
| M2 | docx2md + 图片导出 + 脚注 | ✅ 已实现 |
| M3 | xlsx2md + 中文优化 + 批量 | ✅ 已实现 |
| M4 | pdf2md：文本抽取 + 阅读顺序 + 列对齐表格 | 🚧 已起步（扫描件 OCR、公式、双栏深度检测待做） |

## 🧪 试试（examples/）

仓库自带示例文件，克隆后一行命令即可看到效果：

```bash
node dist/cli.js examples/demo.pptx   # → examples/demo.md
node dist/cli.js examples/demo.pdf    # → examples/demo.pdf.md
```

`examples/` 下的输入文件由 `tests/fixtures/build.ts` 程序化生成，无任何外部素材。

## 🧠 设计

```
doc2md (TypeScript 库 + CLI)
  ├─ parser/       每一格式一个：pptx / docx / xlsx（同步）+ pdf（异步，pdf.js）→ 统一 IR（Doc）
  ├─ core/         OOXML 包读取（zip+rels+编码）、嵌套列表构建、图片资产分配
  ├─ normalize.ts  中文/标点/中英混排规范化（Markdown 感知）
  ├─ md/           IR → Markdown 渲染（表格/转义/脚注）与 front-matter 编排
  └─ cli.ts        cac 驱动：单文件 / 目录批量 / --json / --stdout
```

- 每种格式都输出**同一棵 IR 树**，新增格式只写 parser，渲染管线通用
- 解析器纯同步（fflate 解压 + fast-xml-parser），零原生依赖，浏览器可复用核心；PDF 走异步 pdf.js
- 图片引用延迟分配：解析器只登记 `source → ref`，转换完成统一落盘

## ⚖️ 对比

| 工具 | 语言 | 说明 |
|---|---|---|
| **doc2md** | Node.js | 多格式、中文友好、表格/阅读顺序优先、npm 原生 |
| markitdown | Python | 中文差、无 Node 生态 |
| mammoth.js | Node | 仅 docx，且只转正文、无表格结构 |
| docx-preview | Node | 渲染预览，不输出 Markdown |
| pdf.js | Node | 解析原语，doc2md 将其封装为 Markdown |

## 📝 开发

```bash
npm install
npm run typecheck   # 类型检查
npm test            # 测试（纯代码构造 OOXML 夹具，无二进制）
npm run build       # 产出 dist/（ESM + d.ts + CLI shebang）
node dist/cli.js 你的演示.pptx --stdout
```

## 🔒 合规

只转换**用户提供的文档**，不涉及爬取或破解。输出中保留源文件元数据仅供参考。

## 📄 License

MIT。本项目与微软 markitdown 无关联。