# 功能3：发票转换 TR→天图/航乐

## 概述

将 TR/思锐/赛诺吉 客户下单发票转换为目标供应商模板。

## 支持格式

| 源 | 目标 |
|----|------|
| TR/思锐/赛诺吉发票 | 天图下单发票 |
| 同上 | 航乐-英国发票 |
| 同上 | 航乐-欧洲发票 |
| 同上 | 美琦美线发票 |

## 美琦转换要点

- 模板：`美琦美线发票模版.xlsx`（3 sheet：`发票`/`亚马逊仓库代码`/`服务渠道`）
- 头部 Row 1-17（A=标签/B=值），数据列头 Row 18（A-N），数据行 Row 19 起
- **收件人信息**按「地址库编码」从 `亚马逊仓库代码` sheet 查表（联系人/地址一/城市/省洲/邮编/国家），查不到回退源字段原值
- 地址库编码 = 源「地址库编码」，为空取源「收件人姓名」（仓库代码，如 IND9）
- 客户订单号/客户参考号 = 源「客户订单号」；电话/邮箱/公司名称留空
- 报关方式归一化（美线）：含「退税」→ 一般贸易，含「代理」→ 代理报关
- 海关编码格式化为 `XXXX.XX.XXXX`（如 9405429000 → 9405.42.9000）
- 货箱编号 = 运行式箱号区间（如 1-3）；不做图片嵌入（新版无图片列）

## 源文件要求

- Sheet 名 `Page1`
- Row 1-16/28 头部，Row 18+ 数据行
- 含货箱编号、品名、申报单价/数量、重量尺寸、产品图片

## 天图转换要点

- B3-B13 收件人信息直填客户原值，不做 VLOOKUP
- B1 服务留空（保留下拉验证）
- 产品单箱申报总价 = 单价 × 数量（计算值，非复制）

## 航乐命名规则

`{客户名称} {订单号} {欧洲|英国}发票.xlsx`

## 产品图片（2026-06 修复）

**问题**：转换后图片不显示 / 尺寸过大。

**方案**：`_embed_images_as_cell_images()` 写入标准 Excel 绘图（非 IMAGE() 公式为主）：

| 组件 | 说明 |
|------|------|
| `xl/media/image_N.png` | 图片二进制 |
| `xl/drawings/drawing1.xml` | **twoCellAnchor**，图片跟随单元格大小 |
| worksheet `<drawing>` | 关联 drawing rels |

**源图提取**（`TRInvoice`）：
- 标准 Excel 嵌入图（openpyxl `_images`）
- WPS `cellimages.xml` + DISPIMG 公式（`_extract_wps_cell_images`）

**列位置**：天图 → M 列；航乐 → W 列。数据行行高设为 80。

Web 版 `app.py` 仍可将图片提取到 `/temp_images/`，但本地/CLI 转换以 twoCellAnchor 嵌入为准，打开 xlsx 即可见图。

验证工具：`python3 发票转换/check_format.py <输出.xlsx>`

## 命令行

```bash
cd 发票转换
python3 convert_invoice.py <源发票.xlsx> <输出.xlsx> [--target 天图|航乐-uk|航乐-eu]
```

## Web 接口

- 页面：`GET /invoice`
- 提交：`POST /invoice_convert`（字段：`invoice_file`, `target`）

## 详细规则

见项目根目录 [TR转天图发票_转换规则说明.md](../TR转天图发票_转换规则说明.md)

## 关键代码

| 文件 | 用途 |
|------|------|
| `发票转换/convert_invoice.py` | 转换引擎 + `_embed_images_as_cell_images` |
| `发票转换/check_format.py` | 验证 drawing/media 是否写入 |
| `发票转换/天图单票专用模板*.xlsx` | 天图模板 |
| `发票转换/航乐*.xls` | 航乐模板 |
