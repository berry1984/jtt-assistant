---
name: jtt-invoice-convert
description: Converts TR/思锐/赛诺吉 invoices to 天图 or 航乐 supplier templates. Use when working on 发票转换, convert_invoice.py, TR→天图, 航乐发票, or Page1 sheet mapping.
---

# 发票转换 TR→天图/航乐

## Quick Start

```bash
cd 发票转换
python3 convert_invoice.py <源发票.xlsx> <输出.xlsx> [--target 天图|航乐-uk|航乐-eu]
```

Web: `POST /invoice_convert` with `invoice_file`, `target`.

## Source Format

- Sheet `Page1`, Row 1-16 header, Row 18+ data
- 含货箱编号、品名、申报单价/数量、重量尺寸、产品图片

## Conversion Rules

- **天图**：B3-B13 收件人直填原值；产品总价 = 单价×数量；B1 服务留空
- **航乐**：输出名 `{客户名} {订单号} {欧洲|英国}发票.xlsx`
- 图片提取为 PNG；Web 版用 IMAGE() 引用 HTTP URL

## Key Files

| 文件 | 用途 |
|------|------|
| `发票转换/convert_invoice.py` | 转换引擎 |
| `TR转天图发票_转换规则说明.md` | 完整字段映射 |

## Additional Resources

- [docs/03-invoice-convert.md](../../docs/03-invoice-convert.md)
