---
name: jtt-tr-bill
description: Generates TR FBA billing Excel from order list, picking data, and price files. Use when working on TR账单, 拓锐账单, gen_bill.py, 分段开票, or uploading 订单列表+拣货数据+应收价格.
---

# TR账单自动生成

## Quick Start

```bash
cd "TR账单自动生成"
python3 gen_bill.py 订单列表.xlsx 拣货数据.xlsx 应收价格.xlsx [输出.xlsx]
```

Web: `POST /generate` with `order_file`, `pick_file`, `price_file`.

## Core Rules

- J列计费重 = ROUND(拣货收费重)，调整最大项匹配订单合计
- S列报关费 = 350/1.06，按走货渠道或 J 列合并报关组收取
- 排序：渠道 → SO → FBA
- 输出标题含周区间：`至：广州拓锐科技有限公司（M.D-M.D）`

## Key Files

| 文件 | 用途 |
|------|------|
| `TR账单自动生成/gen_bill.py` | 生成引擎 |
| `TR账单自动生成/账单模板.xlsx` | 输出模板 |
| `TR账单自动生成/app.py` | Flask 路由 `/` |

## Additional Resources

- [docs/01-tr-bill.md](../../docs/01-tr-bill.md)
