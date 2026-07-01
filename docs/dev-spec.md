# JTT 电商 AI 助手 — 开发规范文档

> 本文档面向开发者：在重构、崩溃恢复、新人接手时，可依此快速定位代码与逻辑。
> 最后更新：2026-07-01

---

## 一、项目概述与技术栈

一款**跨境物流一站式工具**，聚合 6 大业务功能于一个 Flask Web 应用，同时保持每个功能模块可 CLI 独立运行。

| 层级 | 技术 | 版本 | 说明 |
|------|------|------|------|
| 运行时 | Python | 3.12 (runtime.txt) / 3.11 (Docker) | Railway 用 Docker，本地用 runtime.txt |
| Web 框架 | Flask | ≥3.0 | 单进程开发服务器 |
| Excel 读写 | openpyxl | ≥3.1 | 全部 .xlsx 读写 |
| 旧版 .xls 读取 | xlrd | ≥2.0 | **只读不写**，延迟导入避免 Railway 崩溃 |
| PDF 生成 | PyMuPDF (fitz) | ≥1.23 | 提单 PDF 红划+插入 |
| 图片处理 | Pillow | ≥10.0 | 发票图片像素→EMU 转换 |
| 部署平台 | Railway | — | 自动从 GitHub main 分支部署 |
| 容器 | Docker | — | `python:3.11-slim` base image |

### 关键依赖路径

```
requirements.txt           # 全量依赖（Web 部署用）
TR账单自动生成/requirements.txt  # Flask+openpyxl+PyMuPDF 子集
runtime.txt                # Python 版本（Railway Buildpack 用）
Dockerfile                 # Docker 构建（pip install 两个 requirements.txt）
Procfile                   # web: cd TR账单自动生成 && python app.py
```

---

## 二、目录结构与各模块职责

```
bb plan1/                              # git 仓库根目录
├── TR账单自动生成/                     # 🔥 主应用（Flask app + 功能1/6）
│   ├── app.py                         # Flask 主程序，聚合 6 大功能的路由
│   ├── gen_bill.py                    # 功能1：TR 番茄钟账单生成（新版，支持报关组分组合并）
│   ├── gen_bl_docs.py                 # 功能6：提单 PDF + 电放保函 xlsx 生成
│   ├── requirements.txt               # Flask + openpyxl + PyMuPDF
│   ├── templates/
│   │   └── index.html                 # 单页应用：6 个 Tab 共用
│   └── templates_bl/                  # 提单 PDF 模板（By sea / By train / By truck）+ 电放保函.xlsx
│
├── 发票转换/                           # 功能3
│   ├── app.py                         # 独立 Web 应用（仅发票转换，端口 5001）
│   ├── convert_invoice.py             # 核心引擎：TRInvoice 解析 + 转天图/航乐
│   ├── convert_template_xls_to_xlsx.py# 辅助：模板格式转换 .xls→.xlsx
│   ├── check_format.py                # 格式验证脚本（无函数，纯脚本）
│   └── templates/
│       └── index.html                 # 发票转换页
│
├── SR账单自动生成/                     # 功能2 的数据目录
│   ├── 问题定位/                       # 2026-06 问题重现数据
│   ├── 思锐账单模板模板 思锐开票账单-JTT（5.1-5.31）.xlsx  # 模板（含汇率V2/AB2、银行信息）
│   ├── 系统账单-*-GDSR-*.xls          # 历史系统账单
│   ├── 思锐订单列表 *.xlsx             # 历史订单列表
│   └── 思锐账单生成规则.docx           # 业务规则文档
│
├── 拣货数据/                           # 功能4
│   ├── export_picking_data.py         # 核心引擎：发票→箱规匹配→拣货数据
│   ├── 内部拣货数据参考值模版.xlsx      # 输出模板
│   ├── 箱规历史数据库.xlsx             # 历史箱规数据
│   └── 报价表.xlsx                    # 渠道报价
│
├── 投保区间拆分发票/                    # 功能5
│   ├── split_insurance_v2.py          # 核心引擎：按RMB拆 5 个投保区间
│   └── 投保区间拆分规则说明_v2.md       # 业务规则文档
│
├── 生成提单及电放保函/                  # 功能6 的早期独立脚本（已整合进 Web，保留参考）
│   ├── generate_docs.py               # 旧版：硬编码路径、固定月份
│   ├── TR 退税资料明细.xlsx            # 测试数据
│   └── 模板/                          # PDF 模板
│
├── gen_sr_bill.py                     # 功能2：思锐(SR)账单生成（根目录，因 sys.path 导入）
├── gen_bill.py                        # 功能1：旧版 TR 账单生成（根目录，CLI 兼容保留）
├── requirements.txt                   # 全量依赖（含 xlrd/Pillow）
├── runtime.txt                        # Python 版本
├── Dockerfile                         # Docker 构建
├── Procfile                           # Railway 启动命令
├── 账单模板.xlsx                       # 功能1 TR 账单模板
├── pomodoro.html                      # 番茄钟（项目无关，历史遗留）
├── README-部署说明.md                  # Railway 部署教程
├── TR转天图发票_转换规则说明.md          # 发票转换规则
│
└── docs/                              # 功能文档
    ├── README.md                      # 总览
    ├── 01-tr-bill.md                  # TR 账单文档
    ├── 02-sr-bill.md                  # SR 账单文档
    ├── 03-invoice-convert.md          # 发票转换文档
    ├── 04-picking-data.md             # 拣货数据文档
    ├── 05-insurance-split.md          # 投保拆分文档
    ├── 06-bl-docs.md                  # 提单保函文档
    └── dev-spec.md                    # ← 本文档
```

---

## 三、已实现的模块

### 3.1 TR 账单自动生成（功能1）

**文件**：`TR账单自动生成/gen_bill.py`（新版）+ `gen_bill.py`（旧版 CLI）  
**入口**：Web `POST /generate` | CLI `python3 gen_bill.py <订单> <拣货> <价格> [输出]`  
**模板**：`账单模板.xlsx`（Sheet: `5月人民币账单（已调格式）`、`开票金额`、`报价表A`）

#### 输入文件

| 文件 | 格式 | 关键列 |
|------|------|--------|
| 订单列表 | .xlsx | 运单号、服务、发货日期、收费重、扩展单号、件数、收件人(仓库代码) |
| 拣货数据 | .xlsx | SO号、FBA箱号、长宽高、实重 |
| 应收价格 | .xlsx | 仓库代码、渠道、应收单价、报关费、J列(报关组) |

#### 核心函数

```python
def load_data(order_path, pick_path, price_path)
    -> (orders: dict, picks: list, prices: dict, price_rows_raw: list, declaration_groups: dict)

def build_rows(orders: dict, picks: list, prices: dict) -> list[dict]

def sort_rows(rows: list[dict], declaration_groups: dict=None) -> list[dict]

def generate_bill(rows, output_path, template_path=None, title_str=None,
                  date_range_str=None, price_rows_raw=None, year=None,
                  declaration_groups=None) -> bool
```

#### 输出 Excel 结构

| 列 | 内容 | 公式 |
|----|------|------|
| A | 日期 | |
| B | SO号（运单号） | |
| C | FBA号 | |
| D | 服务/渠道 | |
| E | 仓库代码 | |
| F | 件数 | |
| G-I | 长/宽/高 | |
| J | 计费重(kg) | ROUND 调整最大项使合计=订单列表 |
| K | 单价 | |
| L | 单价×0.07/1.06 | `=K*0.07/1.06` |
| M | 单价×0.35 | `=K*0.35` |
| N | 单价×0.58 | `=K*0.58` |
| O | 国内运费 | `=L*J` |
| P | 国内运费税额 | `=O*0.06` |
| Q | 国际运费A | `=M*J` |
| R | 国际运费B | `=N*J` |
| S | 报关费（合并显示） | `=350/1.06`，按报关组/渠道合并 |
| T | 报关费税额 | `=S*0.06` |
| U-Y | 预留 | 0 |
| Z | 币种 | RMB |
| AA | 合计 | `=SUM(O:Y)` |

#### 计费重调整逻辑（关键）

```
1. 每个 FBA 子行 weight = ROUND(weight_raw)
2. SUM(子行 weight) vs 订单列表总收费重  → diff
3. diff ≠ 0 → 从 weight_raw 最大的行开始 ±1，直到 diff=0
```

---

### 3.2 思锐(SR) 账单自动生成（功能2）

**文件**：`gen_sr_bill.py`  
**入口**：Web `POST /generate_sr` | CLI `python3 gen_sr_bill.py <系统账单.xls> [输出]`  
**模板**：`SR账单自动生成/思锐账单模板模板 思锐开票账单-JTT（5.1-5.31）.xlsx`

#### 输入文件

| 文件 | 格式 | 说明 |
|------|------|------|
| 系统账单 | .xls | 运通系统导出，Sheet `运单` |
| 思锐订单列表 | .xlsx | 可选，提供渠道/FBA/件数/收费重补充 |
| AB2 汇率 | float | 人民币→欧元，默认 0.1282 |

#### 核心函数

```python
def read_system_bill(xls_path) -> OrderedDict
    """返回 { 运单号: { date, country, weight, ext_no, has_fba,
                         fees: {费用类型: 金额}, fee_unit_prices: {} } }"""
    # ⚠️ 同费用类型自动累加（2026-07-01 修复）

def read_order_list(xlsx_path) -> dict
    """返回 { 运单号: { channel, country, fba, weight, pieces } }"""

def generate_bill(waybills, template_path, output_path,
                  order_list=None, ab2_rate=None) -> bool
```

#### 输出 Excel 列映射（Sheet `原始账单`）

| 列 | 内容 | 费用来源 |
|----|------|----------|
| C | FBA单号 | |
| D | 货代运单号 | |
| E | 渠道 | 订单列表 AH 列 |
| F | 国家 | |
| J | 受益部门 | 国家+FBA → 部门/经理 |
| K | 业务经理 | |
| L | 件数 | 订单列表 D 列 |
| M | 计费重 | ROUND(收费重) |
| N | 单价 | 运费单价 |
| O | 海运费（CNY） | =N×M，或直接取运费金额 |
| Q | 超品名费（CNY） | 超品名费 |
| R | 超尺寸费（CNY） | 超重费 |
| V | 报关费（CNY） | 出口(国内)报关费 |
| Z | 清关费（CNY） | 进口(海外)清关费 |
| AA | 保险费（CNY） | 保费 |
| AB | 目的港费用合计（CNY） | =SUM(Y:AA) |
| AC | 国内运费（CNY） | =U×0.1 |
| AD | 国内运费税费 | =AC×0.06 |
| AE | 国内运费最终开票 | =AC+AD |
| AJ | 国际运费（EUR） | =U×0.9×$AB$2 |
| AN | 税费关税（EUR） | 税金×AB2 或 "后补" |

---

### 3.3 发票转换 TR → 天图/航乐（功能3）

**文件**：`发票转换/convert_invoice.py`  
**入口**：Web `POST /invoice_convert`（主应用）| 独立 Web `python3 发票转换/app.py` | CLI `python3 convert_invoice.py`  
**模板**：`发票转换/天图单票专用模板20260601.xlsx`、`航乐-UK发票.xlsx`、`航乐-EU发票.xlsx`

#### 核心类

```python
class TRInvoice:
    def __init__(self, path: str)
        # 自动检测 TR 格式（Sheet Page1, Row 1-16 头部, Row 18+ 数据）
        # 或 思锐/赛诺吉格式（Sheet 发票, Row 1-26 头部, Row 28+ 数据）
    def get(self, key: str, default=None) -> str
    # 属性: .header (头部字段), .data_rows (数据行), .images (图片映射)
```

#### 转换函数

```python
def convert_to_tiantu(tr: TRInvoice, output_path: str,
                      image_url_base: str=None) -> bool

def convert_to_hangle(tr: TRInvoice, output_path: str,
                      region: str='uk', image_url_base: str=None) -> bool
```

#### 图片处理

TR 发票中的 WPS 图片（`DISPIMG` 公式 + `xl/cellimages.xml`）提取为标准 Excel 图片（`twoCellAnchor` 绘图），写入 `xl/drawings/drawing1.xml`。

---

### 3.4 拣货数据导出（功能4）

**文件**：`拣货数据/export_picking_data.py`  
**入口**：Web `POST /picking_export` | CLI `python3 export_picking_data.py`

#### 核心流程

```
发票(FBA箱号) → 提取前12位 FBA ID
                          ↘ 匹配 → 系统导出拣货数据(扩展箱号) → SO号(运单号)
箱规历史数据库(品名+尺寸)   ↗ 匹配 → 标准箱规参考值(V/W/X/Y)
报价表                    → 应收单价/应付单价/供应商渠道
输出：按SO归组，写入内部拣货数据参考值模版
```

#### 核心函数

```python
def parse_invoice(filepath) -> (list[dict], str, str)
    """返回 (data_rows, service, warehouse)"""

def parse_system_export(filepath) -> dict
    """返回 { FBA_ID_12chars: SO号 }"""

def find_history_match(records, name, weight, length, width, height) -> dict|None
    """匹配箱规历史（6种长宽高排列组合 + 重量±0.5）"""

def generate_picking_output(invoice_file, system_file, output_path, ...) -> int
    """返回总箱数"""

def generate_picking_output_multi(invoice_files, system_file, output_path, ...) -> (str, int)
    """支持多份发票合并输出"""
```

---

### 3.5 投保区间拆分（功能5）

**文件**：`投保区间拆分发票/split_insurance_v2.py`  
**入口**：Web `POST /insurance_split` | CLI `python3 split_insurance_v2.py <下单发票.xlsx>`

#### 核心逻辑

```
1. 解析发票：每行货箱编号 → 按箱号归组
2. 每箱RMB = 申报价值 × 汇率 × 1.1
3. 按RMB分配到5个区间：<5k / 5k-10k / 10k-20k / 20k-30k / 30k+
4. 每个区间生成独立 .xlsx（保留图片、样式）
```

#### 核心函数

```python
def parse_source(src_path) -> dict
    """返回 { currency, rate, service, header_rows, col_headers, data_rows,
              box_groups: {箱号: [data_rows]}, ranges, images }"""

def assign_ranges(box_groups) -> dict
    """返回 { range_name: [箱号列表] }"""

def create_range_output(src_data, range_name, box_groups, output_path) -> None

def split_invoice_to_ranges(src_path, output_dir=None) -> dict
    """返回 { range_name: output_file_path } — Web 入口"""
```

图片处理（来源文件图片 → 到各区间文件的过程：
1. `_get_image_data()`：提取源文件 media/drawing/cellimages/comments
2. `_embed_image_data()`：按保留行过滤 → 重映射行号 → 嵌入输出文件

---

### 3.6 提单 PDF + 电放保函（功能6）

**文件**：`TR账单自动生成/gen_bl_docs.py`  
**入口**：Web `POST /generate_bl_docs` | CLI `from gen_bl_docs import generate_bl_docs`  
**模板**：`TR账单自动生成/templates_bl/提单By sea.pdf`、`提单By train.pdf`、`提单By truck.pdf`、`电放保函.xlsx`

#### 核心函数

```python
def generate_bl_docs(excel_path: str, output_dir: str=None) -> (str, int, int)
    """返回 (zip_path, telex_count, bl_count)"""

def _load_shipments(excel_path: str) -> list[dict]
    """Sheet 自动检测：提单信息 > 提单 > 默认；向下填充空白；过滤查验/备注"""

def _merge_shipments(group: list[dict]) -> dict
    """按 B/L No 合并多票：cartons/KGS/CBM 累加，marks/desc 拼接"""

def _gen_bl(shipment: dict, out_dir: str, ...) -> str
def _gen_telex(shipment: dict, out_dir: str, ...) -> str
```

#### PDF 红划+插入

1. `_sea_train_fields()` / `_truck_fields()` 返回 `[(x0,y0,x1,y1), ...]`（清除矩形）和 `[((x,y), field_key, fontsize, max_width), ...]`（插入点）
2. `page.add_redact_annot()` 加白色遮罩 → `page.apply_redactions()` 擦除
3. `page.insert_text()` 写入新值
4. 长文本自动缩小字号（`_fit_text()`）

---

### 3.7 Flask Web 主应用

**文件**：`TR账单自动生成/app.py`

| 路由 | 方法 | 功能 | 输入字段 |
|------|------|------|----------|
| `/` | GET | TR账单页面 | — |
| `/generate` | POST | TR账单生成 | order_file, pick_file, price_file |
| `/sr` | GET | 思锐账单页面 | — |
| `/generate_sr` | POST | 思锐账单生成 | sr_bill_file, sr_order_file(可选), ab2_rate |
| `/invoice` | GET | 发票转换页面 | — |
| `/invoice_convert` | POST | 发票转换 | invoice_file, target |
| `/picking` | GET | 拣货数据页面 | — |
| `/picking_export` | POST | 拣货数据导出 | picking_invoice(multi), picking_system, picking_history(可选), picking_quotation(可选) |
| `/insurance` | GET | 投保拆分页面 | — |
| `/insurance_split` | POST | 投保拆分 | insurance_file |
| `/bl_docs` | GET | 提单保函页面 | — |
| `/generate_bl_docs` | POST | 提单保函生成 | excel_file |
| `/temp_images/<id>/<file>` | GET | 临时图片（IMAGE公式用） | — |
| `/debug_template` | GET | 调试信息 | — |

#### 响应方式

- 成功：`send_file(output_path, as_attachment=True, ...)` — 浏览器下载
- 失败：`flash(错误信息)` + `redirect(原页面)`
- 前端已从 fetch 下载改为多数表单使用原生提交（避免 Content-Disposition 头不可读）

---

### 3.8 发票转换独立 Web（功能3 子应用）

**文件**：`发票转换/app.py`（端口 5001，独立部署）

| 路由 | 方法 | 功能 |
|------|------|------|
| `/` | GET | 发票转换页面 |
| `/convert` | POST | 发票转换（字段：invoice_file, target） |

---

## 四、非目标 / 留后

以下问题**本次未做**，未来可考虑：

| 项目 | 原因 |
|------|------|
| 数据库持久化 | 所有数据来自上传 Excel，无用户系统，无持久存储需求 |
| 多用户 / 权限 | 单用户内部工具 |
| 异步任务队列 | 每个生成请求 < 5 秒，无需 Celery/RQ |
| 文件上传到 S3 | 临时文件本地 tempdir 生存期=HTTP 请求 |
| WebSocket 实时进度 | fetch 纯同步，无进度推送 |
| 单元测试 | 覆盖率低，仅 `test_price_fix.py` 有测试 |
| CI/CD | 仅有 Railway auto-deploy，无 GitHub Actions |
| 日志系统 | Flask 默认 log，无结构化日志/Error Tracking |
| API 版本控制 | 无第三方 API 消费者 |
| i18n | 全中文界面 |
| 权限分离（功能级） | 全部功能对所有人可见 |
| 模板管理 UI | 模板文件在 git 中，需手动替换 |
| .xls 输出 | 全部输出 .xlsx；xlrd 只读 2.0 不支持 .xlsx |
| 汇率自动获取 | AB2 汇率目前由用户手动输入（SR 账单） |

---

## 五、已知坑 / 绕过的 hack / 待重构项

### 5.1 中文文件名下载

- **问题**：Flask `send_file(download_name=中文)` 在 Railway 上 Content-Disposition 被 WPS/浏览器解析异常
- **Hack 1**（已 revert）：手动设置 `Content-Disposition` 头 → Werkzeug 版本升级导致 bug
- **Hack 2**（当前）：前端 `filename*`（RFC 5987）优先于 `filename`，由 JS 提取并 createObjectURL 下载
- **坑**：某些浏览器/代理可能不可读 Content-Disposition 头 → SR 表单已改为原生提交

### 5.2 xlrd 延迟导入

```python
# gen_sr_bill.py / gen_bill.py (旧版)
import xlrd  # 在函数内延迟导入
wb = xlrd.open_workbook(xls_path)
```

- **原因**：Railway 上缺依赖会导致整个 Flask 启动崩溃；延迟导入让 `ImportError` 不影响其他功能
- **坑**：xlrd 2.0+ **不支持 .xlsx**，仅支持 .xls；若用户上传 .xlsx 给 SR 表单会报错

### 5.3 模板文件路径硬编码

- `gen_sr_bill.py main()`：硬编码 `思锐账单模板模板 思锐开票账单-JTT（5.1-5.31）.xlsx`
- `generate_docs.py`（旧版）：硬编码绝对路径 `TR 退税资料明细.xlsx`
- **待重构**：统一用 `os.path.dirname(__file__)` 相对路径查找；模板名称不应绑定月份

### 5.4 时间选择 → 银影（异步调用

**2026-07-01 修复** 前使用 `import importlib; importlib.import_module('gen_sr_bill')` 延迟导入 SR 模块。

### Image、日志编码、 使用情景全局 / 持久存储

在 src - 工作全流程耗时共享后可用 `tempfile.mkdtemp` 清理，无清理。

### 5.5 发票转换图片嵌入

- **问题**：WPS DISPIMG 公式 + `cellimages.xml` 非标准 Excel 格式 → 转换为标准 `twoCellAnchor` drawing
- **Hack**：`_embed_images_as_cell_images()` 解压 xlsx → 操作 XML → 重打包
- **坑**：图片 EMU 尺寸计算依赖 PIL 读取 DPI；若图片无 DPI 信息默认 72 DPI
- **坑**：`image_url_base`（IMAGE 公式引用）仅开发模式有效，Railway 上需改用嵌入方式

### 5.6 模块相互拷贝（TR 账单新旧版）

- `gen_bill.py`（根目录）和 `TR账单自动生成/gen_bill.py`（目录内）是**两份相似但不同的代码**
- 根目录版：旧版 CLI，硬编码 5 月，无报关组功能
- `TR账单自动生成/` 版：新版 Web，支持报关组合并、动态月份
- **待重构**：统一为同一个文件，用参数/配置区分行为

### 5.7 SR 账单金额累加（2026-07-01 修复）

- **问题**：`read_system_bill()` 中 `wb_info['fees'][fee_type] = amount` 覆盖同费用类型
- **后果**：JTT202606000378 的 4 条税金只取了最后一条 1321.87，丢失 41442.69
- **修复**：改为 `wb_info['fees'][fee_type] = existing + amt`
- **注意**：运费（运费）也可能有同样问题（当前未出现重复记录）

### 5.8 提单 PDF 模板坐标

- `_sea_train_fields()` 和 `_truck_fields()` 中的 (x, y) 坐标是在**特定 PDF 模板**上手测的
- 模板用不同软件编辑后坐标可能偏移 → 需重新测量
- **对齐**：`_fit_text()` 自动缩小字号防止溢出，不能解决坐标偏移

### 5.9 字体跨平台回退

```python
FONT_PATHS = [
    '/System/Library/Fonts/Arial.ttf',          # macOS
    '/usr/share/fonts/truetype/msttcorefonts/Arial.ttf',  # Linux
]
```

- 各平台 Arial 路径不同；Linux 上需 `apt install ttf-mscorefonts-installer`
- 找不到 Arial 则回退到 Liberation Sans → DejaVu Sans → fitz 内置 helv

### 5.10 Docker 构建顺序

```dockerfile
COPY requirements.txt .
COPY TR账单自动生成/requirements.txt ./TR_requirements.txt
RUN pip install -r requirements.txt && pip install -r TR_requirements.txt
COPY . .
```

- 先装依赖再复制源码 → 利用 Docker 层缓存加速
- 但 `COPY . .` 包含 `uploads/`、`__pycache__/` 等 → 需 `.dockerignore`（目前没有）

---

## 六、运行 & 测试命令

### 6.1 Web 主应用

```bash
cd "/Users/admin/bb plan1/TR账单自动生成"
python3 app.py
# 浏览器: http://localhost:5000
```

### 6.2 发票转换独立应用

```bash
cd "/Users/admin/bb plan1/发票转换"
python3 app.py
# 浏览器: http://localhost:5001
```

### 6.3 CLI 命令

```bash
# TR 账单（新版 Web 版本）
cd "/Users/admin/bb plan1/TR账单自动生成"
python3 gen_bill.py <订单列表.xlsx> <拣货数据.xlsx> <应收价格.xlsx> [输出.xlsx]

# TR 账单（旧版 CLI，根目录）
cd "/Users/admin/bb plan1"
python3 gen_bill.py <订单列表.xlsx> <拣货数据.xlsx> <应收价格.xlsx> [输出.xlsx]

# SR 账单
cd "/Users/admin/bb plan1"
python3 gen_sr_bill.py <系统账单.xls> [输出.xlsx]

# 发票转换
cd "/Users/admin/bb plan1"
python3 发票转换/convert_invoice.py <TR发票.xlsx> --to 天图 [输出路径]
python3 发票转换/convert_invoice.py <TR发票.xlsx> --to 航乐-uk [输出路径]
python3 发票转换/convert_invoice.py <TR发票.xlsx> --to 航乐-eu [输出路径]

# 拣货数据导出
cd "/Users/admin/bb plan1/拣货数据"
python3 export_picking_data.py

# 投保区间拆分
cd "/Users/admin/bb plan1"
python3 投保区间拆分发票/split_insurance_v2.py <下单发票.xlsx> [--out-dir ./输出]

# 提单保函（模块调用）
cd "/Users/admin/bb plan1/TR账单自动生成"
python3 -c "from gen_bl_docs import generate_bl_docs; print(generate_bl_docs('data.xlsx'))"
```

### 6.4 测试文件

```bash
cd "/Users/admin/bb plan1/TR账单自动生成"
python3 test_price_fix.py     # 价格修复测试（TR 账单）
```

### 6.5 部署

```bash
git push origin main          # Railway 自动部署
# 或 Docker 手动构建：
docker build -t jtt-assistant .
docker run -p 5000:5000 jtt-assistant
```

---

## 七、对外 API / 数据结构

### 7.1 前端 → 后端 表单字段

见 3.7 节路由表。所有接口为 `POST multipart/form-data`，返回 `.xlsx` 文件或 redirect。

### 7.2 模块间共享数据结构

#### TR 账单 `build_rows` 输出

```python
rows = [
    {
        'so': str,          # 运单号
        'fba': str,         # FBA 箱号
        'date': float|str,  # Excel 序列号 or 字符串
        'service': str,     # 渠道名
        'wh': str,          # 仓库代码（截断 - 后缀）
        'boxes': int,       # 件数
        'length': float,
        'width': float,
        'height': float,
        'weight': int,      # ROUND 后的计费重
        'weight_raw': float,# 原始重量
        'unit_price': float,# 应收单价
    }
]
```

#### SR 账单 `read_system_bill` 输出

```python
waybills = OrderedDict{
    'JTT202606000xxx': {
        'date': '20260605',
        'country': 'DE',          # 发往国家
        'ext_no': '',             # 扩展单号
        'weight': 976.0,          # 收费重
        'has_fba': bool,
        'fees': {
            '运费': float,
            '保费': float,
            '出口(国内)报关费': float,
            '进口(海外)清关费': float,
            '超品名费': float,
            '税金': float,         # 2026-07-01 前只保留最后一条
        },
        'fee_unit_prices': {
            '运费': '4.40/KG',     # 原始字符串
        },
    }
}
```

#### 发票转换 `TRInvoice` 属性

```python
tr.header = {
    '发货人': 'xxx',
    '收件人': 'xxx',
    '服务': 'FBA',
    '客户订单号': 'ORD-xxx',
    # ...
}
tr.data_rows = [
    ['箱号', '产品中文名', '数量', '长', '宽', '高', '实重(KG)', '材重(KG)', ...],
    # ...
]
tr.images = {
    src_row_index: image_bytes,
    # ...
}
```

#### 拣货数据匹配结果

```python
output_rows = [
    {
        'so': str,           # 匹配到的 SO 号
        'fba_id': str,       # 前 12 位
        'service': str,
        'warehouse': str,
        'box_no': str,       # 原始箱号
        'cartons': int,      # 箱数
        'name': str,         # 品名
        'length': float, 'width': float, 'height': float, 'weight': float,
        'ref_length': float|None, 'ref_width': float|None,
        'ref_height': float|None, 'ref_weight': float|None,  # 历史匹配结果
        'unit_price': float|None,        # 应收单价
        'supplier_price': float|None,    # 应付单价
        'supplier_channel': str|None,    # 供应商渠道
    }
]
```

#### 投保区间拆分

```python
# parse_source 输出
src_data = {
    'currency': 'GBP',         # 申报币种
    'rate': 9.0,               # 汇率（英镑×9）
    'service': 'FBA',
    'header_rows': [           # Row 1-26 样式+值
        [{'value': ..., 'style': {...}}, ...],
    ],
    'col_headers': [...],      # Row 27 列标题
    'data_rows': [...],        # Row 28+ 数据
    'box_groups': {            # 按箱号归组
        '110-5598136-1': [data_rows],
        '110-5598136-2': [data_rows],
    },
    'ranges': {                # 每箱RMB
        '110-5598136-1': 4567.89,
    },
}
```

---

## 八、前端要点

### 8.1 文件

- **单一 HTML**：`TR账单自动生成/templates/index.html`（563 行）
- 6 个 Tab 用 CSS `display: none/block` 切换
- 文件选择用 `updateLabel()` 显示文件名

### 8.2 表单提交（重要）

- **大多数表单**：`form.addEventListener('submit', ...)` + `fetch()` + `Content-Disposition` 检测
- **SR 表单**：原生提交（无 fetch），浏览器处理下载
- **blDocs 表单**：fetch + 专门的结果展示 (`showBlDocsResult`)
- **失败回退**：`window.location.reload()` 显示 flash 错误信息

### 8.3 Tab 切换

```javascript
function switchTab(name) {
  // CSS class 切换 + history.replaceState 更新 URL
}
```

---

## 九、紧急恢复指南

### 场景 A：Flask 启动即崩溃

1. 检查依赖：`pip install -r requirements.txt`
2. 检查 xlrd：`python3 -c "import xlrd"` — 如果失败，`pip install xlrd`
3. 检查模板文件路径：`python3 -c "import os; print(os.path.exists('账单模板.xlsx'))"`
4. 日志：Railway 上查看 **Deployment Logs**（`railway logs`）

### 场景 B：SR 账单下载后无文件

1. 检查 `Content-Disposition` 头是否含 `attachment`
2. 浏览器 F12 → Network → 检查 `/generate_sr` 请求的响应头
3. 如果 `content-disposition` 不在响应头 → 服务端异常；检查 Railway 日志
4. 如果头正常但 JS 未触发下载 → 表单已改为原生提交，刷新页面重试

### 场景 C：Excel 生成的账单公式不计算

- 用 `openpyxl` 写入的是公式，打开 Excel 后自动计算
- WPS 可能需要手动点击「允许编辑」→ 公式才计算
- 如果 `data_only=True`，openpyxl 读取的是缓存值（0），不是公式计算结果

### 场景 D：恢复未推送的本地修改

```bash
cd "/Users/admin/bb plan1"
git log --oneline -5                  # 看最后提交
git stash list                        # 检查 stash
git stash pop                         # 恢复最近 stash
```
