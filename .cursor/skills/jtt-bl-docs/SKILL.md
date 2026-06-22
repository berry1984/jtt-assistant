---
name: jtt-bl-docs
description: Generates bill of lading PDFs and telex release guarantee xlsx from Excel 提单信息 sheet. Use when working on 提单, 电放保函, gen_bl_docs.py, B/L, or 退税资料.
---

# 提单 + 电放保函

## Quick Start

Web: `POST /generate_bl_docs` with `excel_file` → ZIP download.

```python
from gen_bl_docs import generate_bl_docs
zip_path, telex_count, bl_count = generate_bl_docs('data.xlsx')
```

## Input

Excel 含 Sheet `提单信息`（如 `5月提单信息`），列：`JTT no.` / `引用模板` / `B/L No.` / `渠道` / `箱数` 等。

## Template Mapping

| 引用模板 | PDF |
|----------|-----|
| By sea | 提单By sea.pdf |
| By train | 提单By train.pdf |
| By truck | 提单By truck.pdf |

- 跳过 Place of receipt = "查验"
- 按 B/L No. 合并；文件名含 JTT号+渠道+箱数

## Key Files

| 文件 | 用途 |
|------|------|
| `TR账单自动生成/gen_bl_docs.py` | Web 模块 |
| `生成提单及电放保函/generate_docs.py` | CLI 脚本 |
| `TR账单自动生成/templates_bl/` | PDF/xlsx 模板 |

## Additional Resources

- [docs/06-bl-docs.md](../../docs/06-bl-docs.md)
