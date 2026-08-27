# 功能6：提单 + 电放保函

> 2026-06 新增，Web 入口 `/bl_docs`

## 概述

上传含 **提单信息** 的 Excel，自动生成 **提单 PDF** + **电放保函 xlsx**，打包 ZIP 下载。

## 输入 Excel

Sheet 自动检测（**不限文件名/月份**）：

1. 含 `提单信息` → 2. 含 `提单` → 3. 第一个 sheet

关键列：`JTT no.` / `引用模板` / `B/L No.` / `渠道` / `cartons` / `KGS` / `CBM` / `Ocean Vessel` 等。

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

## Web 接口

- 页面：`GET /bl_docs`
- 提交：`POST /generate_bl_docs`（字段：`excel_file`）
- 响应头：`X-Telex-Count` / `X-Bl-Count` 为生成数量

## 命令行 / 模块调用

```bash
cd "TR账单自动生成"
python3 -c "from gen_bl_docs import generate_bl_docs; print(generate_bl_docs('data.xlsx'))"
```

早期独立脚本：`生成提单及电放保函/generate_docs.py`（功能已整合进 Web 模块）

## 关键代码

| 文件 | 用途 |
|------|------|
| `TR账单自动生成/gen_bl_docs.py` | Web 主模块 |
| `TR账单自动生成/templates_bl/` | PDF + 电放保函模板（已纳入 git） |
| `TR账单自动生成/app.py` | 路由 `/bl_docs`, `/generate_bl_docs` |
