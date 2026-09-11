---
name: jtt-picking-data
description: Generates 内部拣货数据参考值 by matching invoices with system picking exports and box history database. Use when working on 拣货数据, export_picking_data.py, 箱规历史, SO号匹配, or FBA ID mapping.
---

# 拣货数据参考值

## Quick Start

```bash
cd 拣货数据
python3 export_picking_data.py <发票.xlsx> <系统导出.xlsx> [输出.xlsx] [导出报价表模版.xlsx] [报价方式]
```

Web: `POST /picking_export` with `picking_invoice[]`, `picking_system`, `picking_quotation`(必传), optional `picking_history`, `quotation_source`.

## Matching Logic

1. 货箱编号前12位 = FBA ID → 匹配系统扩展箱号 → SO号
2. 品名+重量尺寸 → 箱规历史数据库 → 标准箱规(V/W/X/Y)
3. 报价来源三选一（`quotation_source` → `price_mode`）：
   - `quotation` → 仓库代码 → 报价表 → 应收/应付单价、供应商渠道
   - `weekly` → 渠道+仓点+计费重 → JTT每周渠道报价表 → 应收单价
   - `export_template` → SO号=运单号 → 导出报价表模版 → 服务/仓库/应收/应付/供应商服务
4. 无匹配 → V/W/X/Y 留空标红；报价模版未覆盖的 SO → C/F/G/H 留空、A–H 标红

## Defaults

- 箱规历史：`拣货数据/箱规历史数据库.xlsx`
- 报价单：`拣货数据/报价表.xlsx`（仅当引擎未传 quotation_file 时的兜底；Web 端报价文件必传）
- 输出模板：`拣货数据/内部拣货数据参考值模版.xlsx`

## Key Files

| 文件 | 用途 |
|------|------|
| `拣货数据/export_picking_data.py` | 生成引擎 |

## Additional Resources

- [docs/04-picking-data.md](../../docs/04-picking-data.md)
