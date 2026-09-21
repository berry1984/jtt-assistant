---
name: jtt-tr-bill
description: Generates TR FBA billing Excel from a single 内部拣货数据参考值 file. Use when working on TR账单, 拓锐账单, gen_bill.py, 分段开票, 账单命名, 中文文件名下载, B2/合计行公式, or 报价表A 去重.
---

# TR账单自动生成

## Quick Start

```bash
cd "TR账单自动生成"
python3 gen_bill.py <内部拣货数据参考值.xlsx> [输出.xlsx]
```

Web: `POST /generate` with `ref_file`（页面 `GET /bill`）。

## Core Rules（2026-09-21 起：与参考值严格一一对应）

- **行序**：账单行 = 参考值表从上到下行序，逐行映射；不按 SO 归并、不重排、无回退匹配
- A 发货日期 = 下单时间；B/C/D/E/F = 系统SO号/FBA ID/客户渠道/仓库代码/总箱数(CTN)
- G/H/I 长宽高 = **参考长/参考宽/参考高**（参考尺寸，非发票实际尺寸）
- J 计费重 = 参考值「计费重」列（公式无缓存值时按模版公式复刻）；K 应收单价 = 「应收单价」列原值，为空即留空
- 组头列（下单时间/客户渠道/国家/仓库代码）为**合并单元格**时按区间取值，区间内每行同值
- S列报关费 = 350/1.06，**每一行**都收（不区分报关组）；T = S×0.06
- 计费重为 0 的行会被丢弃（报关费按行收），行号打印在日志里
- **合计行**逐列求和：F/J/O/P/Q/R/S/T/**U/V/W/X/Y**/AA
- **sheet1 B2** = 第2个sheet 合计行的 `Q+R+U+W+V+X+Y`（国际运费1+国际运费2+清关费+附加费+税金+其他费用+退税损失）
- **报价表A 五元组去重** `(月份, 周, 客户渠道, 仓库代码, 应收单价)`：同渠道+同仓点不同单价各列一条；
  五元组完全相同的重复行只留唯一一行（单价/渠道按归一化键比较：`13`/`'13 '`/`13.0` 同一条）
- 输出标题含周区间：`至：广州拓锐科技有限公司（M.D-M.D）`
- 参考值表按**表头名**取列，新增列（如拣货模块的 `D=VAT号`）不影响解析

## 输出文件名（2026-06 修复）

**动态月份**，不再硬编码「5月」：

```
{year}年{file_month}月拓锐FBA仓-分段开票账单-JTT({date_range_str}) RMB {total}.xlsx
```

- `year` = 发货周区间周一所在年份
- `file_month` = `date_range_str` 第一个 `.` 前的数字（发货周区间起始月）
- 无日期时回退到当前月
- `total` = 四舍五入到 1 位小数

示例：`2026年6月拓锐FBA仓-分段开票账单-JTT(6.16-6.22) RMB 12345.6.xlsx`

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
