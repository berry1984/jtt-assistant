---
name: jtt-bl-docs
description: Generates bill of lading PDFs and telex release guarantee xlsx from Excel 提单信息 sheet. Use when working on 提单, 电放保函, gen_bl_docs.py, B/L, 按B/L合并, 退税资料, or /bl_docs feature.
---

# 提单 + 电放保函（2026-06 新增）

## Quick Start

Web: `GET /bl_docs` → `POST /generate_bl_docs` with `excel_file` → ZIP 下载。

```python
# 在 TR账单自动生成/ 目录下
from gen_bl_docs import generate_bl_docs
zip_path, telex_ok, bl_ok = generate_bl_docs('data.xlsx')
# 响应头 X-Telex-Count / X-Bl-Count 为生成数量
```

## 输入 Excel

Sheet 自动检测（**不限文件名/月份**）：
1. 含 `提单信息` → 2. 含 `提单` → 3. 第一个 sheet

必填列：`JTT no.` / `引用模板` / `B/L No.` / `渠道` / `cartons` / `KGS` / `CBM` 等。

**数据预处理**：
- 向下填充：B/L No、渠道、引用模板、船名航次等空白行继承上行
- 跳过 `Place of receipt = "查验"`
- 跳过非 `JTT` 开头的行（底部备注）

## 引用模板 → PDF

| 引用模板 | 模板文件 |
|----------|----------|
| By sea | `templates_bl/提单By sea.pdf` |
| By train | `templates_bl/提单By train.pdf` |
| By truck | `templates_bl/提单By truck.pdf` |

## 按 B/L No 合并

同一 B/L 多票 JTT → **合并输出 1 份提单 + 1 份电放保函**：
- cartons / KGS / CBM 累加
- Marks & Description 多行 `\n` 拼接
- 多 JTT 文件名：`JTT202605000364,353`（共享前缀 + 逗号序号）

## 输出文件名

```
{JTT}{渠道}{N}件提单.pdf
{JTT}{渠道}{N}件电放保函.xlsx
```

ZIP：`提单电放保函_{YYYYMMDD_HHMMSS}.zip`

## PDF / 电放保函细节

- **字体**：Arial 10.5（替代原 Calibri/SimSun）；跨平台回退 macOS Arial → Linux Liberation/DejaVu → helv
- **长文本**：超框自动缩小字号（船名航次等）
- **KGS**：显示为 `{值}\nKGS`，插入点右移避免遮挡蓝线
- **电放保函**：A16/A18 为固定标签文案，不写入 shipper/consignee 变量

## Key Files

| 文件 | 用途 |
|------|------|
| `TR账单自动生成/gen_bl_docs.py` | Web 主模块 |
| `TR账单自动生成/templates_bl/` | PDF + 电放保函模板（已纳入 git） |
| `生成提单及电放保函/generate_docs.py` | 早期 CLI 脚本 |
| `TR账单自动生成/app.py` | 路由 `/bl_docs`, `/generate_bl_docs` |

## Additional Resources

- [docs/06-bl-docs.md](../../docs/06-bl-docs.md)
