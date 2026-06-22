---
name: jtt-insurance-split
description: Splits 思锐 invoice Excel into 5 insurance range files by per-box RMB value. Use when working on 投保拆分, split_insurance_v2.py, 投保区间, or 每箱申报价值.
---

# 投保区间拆分

## Quick Start

```bash
cd 投保区间拆分发票
python3 split_insurance_v2.py <下单发票.xlsx> --out-dir ./输出
```

Web: `POST /insurance_split` with `insurance_file` → ZIP download.

## Core Formula

```
每箱RMB = 单箱子货值 × 汇率 × 1.1
```

汇率：英镑×9 · 欧元×8 · 美金×7

## 5 Ranges

不足5000 · 5000-10000 · 10000-20000 · 20000-30000 · 30000-40000 (RMB)

按 C1(货箱编号) 整箱分组；辅助 Sheet（FBA地址库/服务名称/换算）原样复制。

## Source Structure

- 主 Sheet `发票`：Row 1-26 头部，Row 27 表头，Row 28+ 数据

## Key Files

| 文件 | 用途 |
|------|------|
| `投保区间拆分发票/split_insurance_v2.py` | 拆分引擎 |
| `投保区间拆分发票/投保区间拆分规则说明_v2.md` | 完整规格 |

## Additional Resources

- [docs/05-insurance-split.md](../../docs/05-insurance-split.md)
