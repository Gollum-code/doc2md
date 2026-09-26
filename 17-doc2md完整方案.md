# 17. doc2md（多格式文档 → Markdown）— 完整方案

> ⭐ 格式转换 | 面向 AI/RAG 的风口 | 微软 markitdown 验证了方向但它是 Python 且中文差，npm 生态无高质量等价物
> 切入点：先做 ppt2md（最快出效果），再扩展多格式，PDF 放最后

---

## 一、定位

一个 Node.js 库 + CLI，把 **docx / pptx / xlsx / pdf → 高质量 Markdown**，主打**中文友好 + 表格/公式/阅读顺序**，输出为 LLM/RAG 能直接消费的干净文本。

## 二、为什么要做（机会证据）

1. 做 RAG / AI 应用时，"用户上传 PPT/Word/PDF → 转成干净文本喂 LLM"是**人人头疼的痛点**。
2. 微软 `markitdown` 已验证这个方向，但：
   - 它是 **Python**（Node 生态没有等价物）
   - 中文支持差、表格/公式/阅读顺序处理烂
3. **npm 生态没有高质量的多格式转 Markdown 库**——这是明确空白。
4. RAG/AI 应用爆发 → 文档解析需求持续增长。

## 三、目标用户

- 做 RAG、AI 助手、知识库的 JS/TS 团队
- 需要批量文档转文本的数据处理团队
- 用 Node 生态做 AI 应用的开发者

## 四、技术架构

```
doc2md (TypeScript 库 + CLI)
  ├─ 解析器（每种格式一个）
  │    ├─ pptx2md：读 PPTX (XML) → 每页标题/要点/表格/备注
  │    ├─ docx2md：读 DOCX (XML) → 段落/标题/表格/列表/图片引用
  │    ├─ xlsx2md：读 XLSX → 工作表 → Markdown 表格
  │    └─ pdf2md：读 PDF → 文本抽取 + 表格还原 + 阅读顺序（最难，放最后）
  ├─ 输出管线
  │    ├─ Markdown 渲染（表格、代码块、图片引用）
  │    ├─ 中文优化（标点/换行/中英混排）
  │    └─ 图片处理（导出为本地文件或引用路径）
  └─ CLI：`doc2md input.pptx -o output.md`
```

## 五、MVP 功能清单（先做 pptx2md + docx2md）

| 模块 | 工具 |
|---|---|
| PPTX | 解析幻灯片：标题层级、要点、表格、演讲者备注、图片导出 |
| DOCX | 解析段落/标题结构、表格、列表、链接、图片 |
| XLSX | 读取所有 sheet，转为 Markdown 表格 |
| 输出 | Markdown 渲染 + 中文标点/换行优化 |
| 图片 | 从文档中导出图片到目录，Markdown 引用 |
| CLI | 单文件/目录批量转换、`--gfm` / `--中文` 选项 |

### PDF（M4 之后，最难的）
- 文本抽取（pdf.js / pdf-lib）
- 表格还原（列对齐启发式）
- 阅读顺序（双栏检测）

## 六、Roadmap

| 阶段 | 时间 | 交付 |
|---|---|---|
| M1 | 1~2 周 | pptx2md（最快出效果）+ CLI + 表格 |
| M2 | 1~2 周 | docx2md + 图片导出 |
| M3 | 1~2 周 | xlsx2md + 中文优化 + 批量 |
| M4 | 2~4 周 | pdf2md（文本+表格+阅读顺序） |

## 七、关键难点与验证

| 难点 | 应对 |
|---|---|
| 中文排版（换行/标点/中英混排） | 专门的文本规范化管线 |
| 表格还原 | XML 结构天然有表格信息（PPTX/DOCX 容易）；PDF 用启发式 |
| PPT 备注/隐藏页 | 可选参数控制是否导出 |
| PDF 阅读顺序（双栏） | 用坐标分析，放最后做 |

## 八、竞品分析

| 工具 | 语言 | 差异 |
|---|---|---|
| markitdown | Python | Node 生态空白，中文差 |
| mammoth.js | Node | 只 docx，且弱（只转正文） |
| docx-preview | Node | 渲染预览，不输出 markdown |
| pdf.js | Node | 只是 PDF 解析原语，不封装为 markdown |

## 九、拿星策略

- 标签：`markdown` `pdf` `docx` `pptx` `rag` `llm` `npm`
- README 核心：「把任意文档转成 LLM 能读的干净 Markdown——中文友好」
- 演示：一张"真实 PPT → 结构清晰 Markdown"对比图
- 发 npm + 写《为什么 markitdown 不适合中文场景》引流

## 十、目录结构（MVP）

```
doc2md/
├─ src/
│  ├─ index.ts          # 入口
│  ├─ pptx/
│  │  ├─ parser.ts
│  │  └─ to-md.ts
│  ├─ docx/
│  │  ├─ parser.ts
│  │  └─ to-md.ts
│  ├─ xlsx/
│  ├─ normalize.ts      # 中文/标点规范化
│  └─ cli.ts
├─ tests/
├─ README.md
└─ package.json
```

## 十一、风险
- 文档格式细节繁多，维护成本中高（每种格式的边界案例）
- PDF 最难的表格/阅读顺序，放最后做（先用 pptx/docx 建立口碑）
- markitdown 可能发布 JS 版（跟进保持差异：中文 + npm 原生）
- 版权/合规：只转"用户提供"的文档，不做爬取/破解