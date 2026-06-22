# 功能6：提单 + 电放保函

## 概述

上传含 **提单信息** 的 Excel，自动生成 **提单 PDF** + **电放保函 xlsx**，打包 ZIP 下载。

## 输入要求

- Excel 含 Sheet `提单信息`（如 `5月提单信息`）
- 关键列：`JTT no.` / `引用模板` / `B/L No.` / `渠道` / `箱数` / `Ocean Vessel` 等

## 引用模板 → PDF

| 引用模板 | 提单模板 |
|----------|----------|
| By sea | 提单By sea.pdf |
| By train | 提单By train.pdf |
| By truck | 提单By truck.pdf |

## 规则

- 跳过 `Place of receipt = "查验"` 的货件
- 按 B/L No. 合并同类提单
- 文件名：`{JTT号}_{渠道}_{箱数}箱_提单.pdf` / `..._电放保函.xlsx`
- 固定 Shipper / Consignee 信息写入 PDF

## 输出

ZIP 包内含全部提单 PDF 和电放保函 xlsx。

## 命令行

```bash
cd 生成提单及电放保函
python3 generate_docs.py
# 或 Flask 模块调用：
python3 -c "from TR账单自动生成.gen_bl_docs import generate_bl_docs; print(generate_bl_docs('data.xlsx'))"
```

## Web 接口

- 页面：`GET /bl_docs`
- 提交：`POST /generate_bl_docs`（字段：`excel_file`）

## 关键代码

- Web 模块：`TR账单自动生成/gen_bl_docs.py`
- CLI 脚本：`生成提单及电放保函/generate_docs.py`
- 模板：`TR账单自动生成/templates_bl/` 或 `生成提单及电放保函/模版/`
