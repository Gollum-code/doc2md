# doc2md 项目说明

## 设计目标

把 PPTX、Word 文档转成**干净的 Markdown**，方便喂给 RAG 检索链路。*中文排版不乱码。*

安装命令：`npm i doc2md`

1. 支持表格
  - 支持嵌套列表
1. 支持图片导出
1. 第一阶段交付 CLI
1. 第二阶段交付库

> 表格是 RAG 里最容易丢失的结构，必须保住。

参考与本页数据[^1][官方文档](https://example.com/docs)

目录标题后续说明

---
第二页内容

![demo.assets/image1.png](demo.assets/image1.png)

*流程图*

| 模块 | 状态 |
| --- | --- |
| pptx2md | 已完成 |
| docx2md 也在同一期 |  |

## 脚注
[^1]: 口径以财务系统为准。
