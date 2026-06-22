---
name: jtt-sr-bill
description: Generates 思锐开票账单 from system bill (.xls) and optional order list. Use when working on SR账单, gen_sr_bill.py, 思锐开票, AB2汇率, or 运通系统账单.
---

# 思锐账单自动生成

## Quick Start

```bash
python3 gen_sr_bill.py 系统账单.xls [输出.xlsx]
```

Web: `POST /generate_sr` with `sr_bill_file`, optional `sr_order_file`, `ab2_rate` (default 0.1282).

## Core Rules

- 系统账单 Sheet `运单`，按运单号分组
- 订单列表匹配：AH列渠道 · C列FBA · D列件数 · J列收费重
- 费用映射：运费→O, 超品名→Q, 超重→R, 报关→V, 清关→Z, 保费→AA
- 税费关税(AN) = 系统税金 × AB2汇率，取不到显示「后补」
- 受益部门按国家(DE/GB) + FBA 自动判定

## Key Files

| 文件 | 用途 |
|------|------|
| `gen_sr_bill.py` | 生成引擎 |
| `SR账单自动生成/思锐账单模板*.xlsx` | 输出模板 |

## Additional Resources

- [docs/02-sr-bill.md](../../docs/02-sr-bill.md)
