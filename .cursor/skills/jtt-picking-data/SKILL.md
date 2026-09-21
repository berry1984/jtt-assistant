---
name: jtt-picking-data
description: Generates 内部拣货数据参考值 by matching invoices with system picking exports and box history database. Use when working on 拣货数据, export_picking_data.py, 箱规历史, SO号匹配, FBA ID mapping, or VAT号 列.
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
2. 品名+重量尺寸 → 箱规历史数据库 → 标准箱规（写入期 V/W/X/Y；输出 W/X/Y/Z）
3. **VAT号（D 列，2026-09-21 新增）**：每张发票各自识别——表头「VAT号*」右侧那一格
   （天图/TR：表头 E10、值 F10；航乐：表头 A7、值 B7），取 **DE 开头**的号；
   识别不到/未填写 → 留空，不回退其它发票。输出列序：A 下单时间 | B 系统SO号 |
   C 客户渠道 | **D VAT号** | E 国家 | F 仓库代码 | …（D 及以后 = 写入期 +1）
4. 报价来源三选一（`quotation_source` → `price_mode`）：
   - `quotation` → 仓库代码 → 报价表 → 应收/应付单价、供应商渠道
   - `weekly` → 渠道+仓点+计费重 → JTT每周渠道报价表 → 应收单价
   - `export_template` → SO号=运单号 → 导出报价表模版 → 服务/仓库/应收/应付/供应商服务
     （SO取不到时按收件人/仓点+渠道兜底，并回填 B 列 SO号；同键多价或
       运单号已被占用的则拒绝兜底）
   
   **报价来源以文件为准，不信下拉框**：`detect_quotation_file_type()` 嗅到
   「运单号」+「应收/成本运费单价」表头即强制走 `export_template`。选「报价单」
   （默认项）却传了模版时，模版会被当报价单解析——输出 G 列（应收单价）碰巧对、
   H/I 全空、仓库代码为空的行丢失，SO号/兜底逻辑完全不走，症状酷似「差一点没匹配上」。
   反向：选 export_template 却传非模版 → 告警 + 整表标红。
5. 无匹配 → V/W/X/Y（输出 W/X/Y/Z）留空标红；报价模版未覆盖 → C/F/G/H 留空、A–H 标红
   （输出：C/G/H/I 留空、A–I 标红）

> 列号两套：**写入期**（VAT号 插入前）/ **输出期**（D 及以后 +1）。
> openpyxl `insert_cols` 不翻译公式、不搬合并区、不重映射列宽 → 公式用 `_shift_formula_cols(f, 1)`，
> 列宽按 `idx>=4 → idx+1` 快照重排，数据区先 `unmerge`，标红行 D 列补红底。

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
