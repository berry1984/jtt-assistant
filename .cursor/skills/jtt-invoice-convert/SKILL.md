---
name: jtt-invoice-convert
description: Converts TR/思锐/赛诺吉 invoices to 天图, 航乐, or 美琦 supplier templates. Use when working on 发票转换, convert_invoice.py, TR→天图, 航乐发票, 美琦发票, 产品图片不显示, twoCellAnchor, or Page1 sheet mapping.
---

# 发票转换 TR→天图/航乐/美琦

## Quick Start

```bash
cd 发票转换
python3 convert_invoice.py <源发票.xlsx> <输出.xlsx> [--target 天图|航乐-uk|航乐-eu|美琦]
```

Web: `POST /invoice_convert` with `invoice_file`, `target`.

## Source Format

- Sheet `Page1`，Row 1-16 header，Row 18+ 数据行
- 含货箱编号、品名、申报单价/数量、重量尺寸、产品图片

## Conversion Rules

- **订单列表匹配（美琦/天图/航乐通用）**：可另传「订单列表 excel」（含「运单号」「仓库代码」「供应商服务」列）。匹配优先级：① 源「客户订单号」== 订单列表「运单号」；② 源「地址库编码」（为空取「收件人姓名」）== 订单列表「仓库代码」。命中行回填两项：**运单号**→客户订单号（美琦 B1、天图 B14、航乐 输出文件名）；**供应商服务**→服务/渠道（**渠道抓取**：美琦 B3、天图 B1、航乐 I7）。未传/未命中/无该列则保留源值。CLI 用 `--order-list`。
- **天图**：B3-B13 收件人直填原值；产品总价 = 单价×数量；B1 服务 = 订单列表「供应商服务」（未命中回退源服务）并追加到 Sheet2 下拉
- **航乐**：输出名 `{客户名} {订单号} {欧洲|英国}发票.xlsx`；I7 渠道 = 订单列表「供应商服务」（未命中回退源服务）
- **美琦**：收件人信息按地址库编码从 `亚马逊仓库代码` sheet 查表；**海关编码保持源原值（不加小数点）**；报关方式含「退税」→ 一般贸易；渠道未映射时追加到 `服务渠道` 下拉；数据列 A-R，O=产品图片（源图嵌入）、P=PO Number（源 V 列）、Q=物品箱号（单行总箱数）、R=物品FBA ID（货箱编号 `U00000` 前 12 位）

## 产品图片（2026-06 修复）

**问题**：转换后图片不显示 / 尺寸过大。

**方案**：`_embed_images_as_cell_images()` 写入标准 Excel 绘图，不用 IMAGE() 公式为主：

| 组件 | 说明 |
|------|------|
| `xl/media/image_N.png` | 图片二进制 |
| `xl/drawings/drawing1.xml` | **twoCellAnchor**，图片跟随单元格大小 |
| worksheet `<drawing>` | 关联 drawing rels |

**源图提取**（`TRInvoice`）：
- 标准 Excel 嵌入图（openpyxl `_images`）
- WPS `cellimages.xml` + DISPIMG 公式（`_extract_wps_cell_images`）

**天图** → M 列；**航乐** → W 列；**美琦** → O 列（同时保留模板表头图 Row≤17）。行高设为 80 以容纳图片。

Web 版 `app.py` 仍提取图片到 `/temp_images/`，但 CLI/本地转换以 twoCellAnchor 嵌入为准，打开 xlsx 即可见图。

## Key Files

| 文件 | 用途 |
|------|------|
| `发票转换/convert_invoice.py` | 转换引擎 + `_embed_images_as_cell_images` |
| `发票转换/check_format.py` | 验证 drawing/media 是否写入 |
| `TR转天图发票_转换规则说明.md` | 完整字段映射 |

## Additional Resources

- [docs/03-invoice-convert.md](../../docs/03-invoice-convert.md)
