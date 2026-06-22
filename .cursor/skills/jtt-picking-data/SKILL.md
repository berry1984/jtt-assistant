---
name: jtt-picking-data
description: Generates 内部拣货数据参考值 by matching invoices with system picking exports and box history database. Use when working on 拣货数据, export_picking_data.py, 箱规历史, SO号匹配, or FBA ID mapping.
---

# 拣货数据参考值

## Quick Start

```bash
cd 拣货数据
python3 export_picking_data.py <发票.xlsx> <系统导出.xlsx> [输出.xlsx]
```

Web: `POST /picking_export` with `picking_invoice[]`, `picking_system`, optional `picking_quotation`, `picking_history`.

## Matching Logic

1. 货箱编号前12位 = FBA ID → 匹配系统扩展箱号 → SO号
2. 品名+重量尺寸 → 箱规历史数据库 → 标准箱规(V/W/X/Y)
3. 仓库代码 → 报价表 → 应收/应付单价、供应商渠道
4. 无匹配 → V/W/X/Y 留空标红

## Defaults

- 箱规历史：`拣货数据/箱规历史数据库.xlsx`
- 报价单：`拣货数据/报价表.xlsx`
- 输出模板：`拣货数据/内部拣货数据参考值模版.xlsx`

## Key Files

| 文件 | 用途 |
|------|------|
| `拣货数据/export_picking_data.py` | 生成引擎 |

## Additional Resources

- [docs/04-picking-data.md](../../docs/04-picking-data.md)
