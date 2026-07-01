# JTT电商AI助手 — 功能文档

> 在线版：[https://jtt-assistant-production.up.railway.app](https://jtt-assistant-production.up.railway.app)  
> 本地启动：`cd TR账单自动生成 && python3 app.py`

## 最近更新

| 日期 | 变更 |
|------|------|
| 2026-07-01 | [📘 开发规范](dev-spec.md) 新增：涵盖技术栈/目录/接口/已知坑/恢复指南 |
| 2026-07-01 | SR 账单：修复同费用类型覆盖→累加（税金多条分录只取最后一条 bug） |
| 2026-07-01 | SR 表单：修复 fetch 下载不可用（改原生提交） |
| 2026-06 | TR 账单：输出文件名动态月份；Web 中文文件名下载修复 |
| 2026-06 | 发票转换：图片改用 twoCellAnchor 标准绘图嵌入 |
| 2026-06 | 提单保函：新增功能（按 B/L 合并、Sheet 自动检测、ZIP 打包） |

## 六大功能

| # | 功能 | Web 路径 | 脚本 | 文档 |
|---|------|----------|------|------|
| 1 | TR账单（3个Excel） | `/` | `TR账单自动生成/gen_bill.py` | [01-tr-bill.md](01-tr-bill.md) |
| 2 | 思锐账单（2个Excel） | `/sr` | `gen_sr_bill.py` | [02-sr-bill.md](02-sr-bill.md) |
| 3 | 发票转换 TR→天图/航乐 | `/invoice` | `发票转换/convert_invoice.py` | [03-invoice-convert.md](03-invoice-convert.md) |
| 4 | 拣货数据参考值 | `/picking` | `拣货数据/export_picking_data.py` | [04-picking-data.md](04-picking-data.md) |
| 5 | 投保区间拆分（5区间） | `/insurance` | `投保区间拆分发票/split_insurance_v2.py` | [05-insurance-split.md](05-insurance-split.md) |
| 6 | 提单 + 电放保函 | `/bl_docs` | `TR账单自动生成/gen_bl_docs.py` | [06-bl-docs.md](06-bl-docs.md) |

## 项目结构

```
bb plan1/
├── TR账单自动生成/
│   ├── app.py              # Flask 主程序（聚合 6 大功能）
│   ├── gen_bill.py         # 功能1
│   └── gen_bl_docs.py      # 功能6
├── gen_sr_bill.py          # 功能2
├── 发票转换/               # 功能3
├── 拣货数据/               # 功能4
├── 投保区间拆分发票/       # 功能5
├── docs/                   # 本文档目录
└── .cursor/skills/         # Cursor Agent Skills
```

## 开发参考

→ [📘 **开发规范文档 (dev-spec.md)**](dev-spec.md)  
包含：技术栈版本、目录结构、各模块接口、已知坑（中文下载/xlrd延迟/坐标偏移等）、运行命令、紧急恢复指南。

## Cursor Skills

每个功能对应一个 Skill，Agent 在处理相关任务时会自动加载：

| Skill | 目录 |
|-------|------|
| `jtt-tr-bill` | `.cursor/skills/jtt-tr-bill/` |
| `jtt-sr-bill` | `.cursor/skills/jtt-sr-bill/` |
| `jtt-invoice-convert` | `.cursor/skills/jtt-invoice-convert/` |
| `jtt-picking-data` | `.cursor/skills/jtt-picking-data/` |
| `jtt-insurance-split` | `.cursor/skills/jtt-insurance-split/` |
| `jtt-bl-docs` | `.cursor/skills/jtt-bl-docs/` |
