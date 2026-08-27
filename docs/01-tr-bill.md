# 功能1：TR账单自动生成

## 概述

上传 **订单列表** + **拣货数据** + **应收价格** 三个 Excel，自动生成拓锐 FBA 分段开票账单。

## 输入文件

| 文件 | 说明 |
|------|------|
| 订单列表 `.xlsx` | 含运单号、创建日期、发货日期、工作日期等 |
| 拣货数据 `.xlsx` | 含 SO 号、FBA、收费重、走货渠道 |
| 应收价格 `.xlsx` | 渠道/仓库代码/单价/报关费；J 列标记合并报关组 |

## 核心规则

- **J列(计费重)** = ROUND(拣货收费重)，调整最大项使合计匹配订单列表
- **S列(报关费)** = 350/1.06，每个唯一走货渠道（或合并报关组）收取一次
- **排序**：按渠道分组 → SO → FBA
- **汇总公式**：O/P/Q/R/S/T 列按重量×单价比例计算（7%、6%、35%、58% 等）
- **标题**：`至：广州拓锐科技有限公司（M.D-M.D）` 按创建日期周区间（创建日期为空回退发货日期→工作日期）

## 输出文件名（2026-06 修复）

**动态月份**，不再硬编码「5月」：

```
{year}年{file_month}月拓锐FBA仓-分段开票账单-JTT({date_range_str}) RMB {total}.xlsx
```

| 字段 | 规则 |
|------|------|
| `year` | 发货周区间周一所在年份 |
| `file_month` | `date_range_str` 第一个 `.` 前的数字（发货周区间起始月） |
| `date_range_str` | 按订单**创建日期**所在周（周一~周日），如 `6.16-6.22` |
| `total` | 应收合计，四舍五入到 1 位小数 |

示例：`2026年6月拓锐FBA仓-分段开票账单-JTT(6.16-6.22) RMB 12345.6.xlsx`

无创建日期时回退发货日期→工作日期；均无则 `file_month` 回退到当前月。

## Web 下载（中文文件名）

- 后端：`send_file(..., download_name=output_name)` 直接传 UTF-8 中文名
- 前端 fetch 解析 `Content-Disposition`：**优先** `filename*=UTF-8''...`，再回退 `filename=`
- Railway 环境依赖 Werkzeug `download_name`，不要手动拼 `filename*` header

## 命令行

```bash
cd "TR账单自动生成"
python3 gen_bill.py 订单列表.xlsx 拣货数据.xlsx 应收价格.xlsx [输出.xlsx]
```

## Web 接口

- 页面：`GET /`
- 提交：`POST /generate`（字段：`order_file`, `pick_file`, `price_file`）

## 关键代码

| 文件 | 用途 |
|------|------|
| `TR账单自动生成/gen_bill.py` | 生成引擎（L723 动态月份） |
| `TR账单自动生成/账单模板.xlsx` | 输出模板 |
| `TR账单自动生成/app.py` | Flask 路由 `/` |
| `TR账单自动生成/templates/index.html` | fetch 下载 + filename* 解析 |
