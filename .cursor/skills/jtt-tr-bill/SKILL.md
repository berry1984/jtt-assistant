---
name: jtt-tr-bill
description: Generates TR FBA billing Excel from order list, picking data, and price files. Use when working on TR账单, 拓锐账单, gen_bill.py, 分段开票, 账单命名, 中文文件名下载, or uploading 订单列表+拣货数据+应收价格.
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

## 输出文件名（2026-06 修复）

**动态月份**，不再硬编码「5月」：

```
{file_month}月拓锐FBA仓-分段开票账单-JTT({date_range_str}) RMB {total}.xlsx
```

- `file_month` = `date_range_str` 第一个 `.` 前的数字（发货周区间起始月）
- 无日期时回退到当前月
- `total` = 四舍五入到 1 位小数

示例：`6月拓锐FBA仓-分段开票账单-JTT(6.16-6.22) RMB 12345.6.xlsx`

## Web 下载（中文文件名）

- 后端：`send_file(..., download_name=output_name)` 直接传 UTF-8 中文名
- 前端 fetch 解析 `Content-Disposition`：**优先** `filename*=UTF-8''...`，再回退 `filename=`
- Railway 环境不要用 `filename*` 手动拼 header（不兼容）；依赖 Werkzeug `download_name`

## Key Files

| 文件 | 用途 |
|------|------|
| `TR账单自动生成/gen_bill.py` | 生成引擎（L723 动态月份） |
| `TR账单自动生成/账单模板.xlsx` | 输出模板 |
| `TR账单自动生成/app.py` | Flask 路由 `/` |
| `TR账单自动生成/templates/index.html` | fetch 下载 + filename* 解析 |

## Additional Resources

- [docs/01-tr-bill.md](../../docs/01-tr-bill.md)
