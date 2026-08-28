"""
FBA 仓点库 + 覆盖渠道价查询模块
============================
- WAREHOUSE_POINTS：各国常用 FBA 仓点静态库（依据天图/每周报价/航乐文件分析整理），
  部署 Railway 无外部依赖也能用。
- extract_warehouse_coverage(fp)：从供应商报价文件提取 仓点码 → {渠道: 单价}。
- query_warehouses(data, country)：聚合所有已存报价的覆盖渠道价。

覆盖渠道数据随供应商文件上传后缓存进 price_supplier.load_prices() 的
`data[sup][country]['warehouses']`；JTT 每周报价的仓点覆盖用
weekly_quotation.match_warehouse 反向匹配。
"""

import os
import re

from openpyxl import load_workbook

from supplier_parser import (
    _skip_sheet, _looks_like_channel, _normalize_channel, _to_kg_price,
    _parse_sheet_rows, MIN_KG_PRICE, MAX_KG_PRICE, MAX_SHEET_ROWS, MAX_SHEET_COLS,
)

# JTT 每周报价伪供应商名（price_supplier 复用，避免重复定义）
WEEKLY_SUPPLIER_NAME = "JTT每周报价"
WEEKLY_COUNTRY = "每周"

# 仓点码：字母开头 + 数字（含 WM-LAX1 带连字符形式）。
# 不用 \b —— Python 把中文当单词字符，\b 在「国DTM2」间不成立，用字母/数字前后瞻替代。
FBA_CODE_RE = re.compile(r"(?<![A-Z0-9])((?:[A-Z]{2,4}-)?[A-Z]{2,4}\d{1,2})(?![A-Z0-9])")


# ── 各国常用仓点库（区域分组） ──

WAREHOUSE_POINTS = {
    "美国": {
        "西部": ["ONT8", "LGB8", "LAX9", "SBD1", "POC1", "POC2", "POC3",
                "IUSP", "IUSJ", "IUSQ", "IUTI", "WM-LAX1", "WM-LAX2",
                "GYR1", "PHX3", "PHX7", "LAS1", "SCK1", "SMF1"],
        "中部": ["FTW1", "DFW2", "HOU2", "MEM1", "STL4"],
        "东部": ["IND9", "CVG1", "CMH1", "BNA1", "MDW2", "ORD2", "RDU1",
                "CLT2", "MCO2", "TPA1", "JAX2", "PHL1", "ACY1", "BWI1",
                "ATL1", "MQJ1", "RIC2", "CHA1", "CHO1"],
    },
    "加拿大": {
        "多伦多": ["YYZ1", "YYZ2", "YYZ3", "YYZ4", "YYZ5", "YYZ6", "YYZ7", "YYZ8", "YYZ9"],
        "渥太华": ["YOW1", "YOW3"],
        "汉密尔顿": ["YHM1"],
        "温哥华": ["YVR1", "YVR2", "YVR3", "YVR4", "YVR5", "YVR6", "YVR7", "YVR8", "YVR9",
                  "YXX1", "YXX2"],
        "卡尔加里": ["YYC1", "YYC4", "YYC6"],
        "埃德蒙顿": ["YEG1", "YEG2"],
        "温尼伯": ["YWG1"],
    },
    "英国": {
        "英国": ["BHX4", "LBA4", "LBA5", "MAN1", "DSA1", "EMA1", "NCL1",
                "BRS1", "PIK1", "CWL1", "STN1"],
    },
    "欧洲": {
        "德国": ["DTM2", "HAJ1", "HAJ2", "KSF1", "LEJ1", "MUC3", "BER3",
                "DUS2", "XDU1", "XDP1", "XDX1"],
        "波兰": ["WRO5", "POZ1", "WAW1", "WAW2"],
        "法国": ["LYS1", "CDG1", "ORY1"],
    },
}


# ── 每周报价表模块（懒加载，拣货数据目录不在 sys.path 默认里） ──

def _load_weekly_module():
    import sys
    try:
        import weekly_quotation
        return weekly_quotation
    except ImportError:
        d = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), '拣货数据')
        if d not in sys.path:
            sys.path.insert(0, d)
        import weekly_quotation
        return weekly_quotation


def _weekly_covers(pattern, code):
    """JTT 每周报价仓点模式是否覆盖给定仓点码（严格版，不做 FBA 兜底）。

    只认显式仓点码/范围（ONT8、YYZ1-9、DTM2-44145）与国家模式（德国/波兰），
    邮编前缀（8,9 / 97.98.99开头）与 FBA 泛兜底无法映射到具体仓点码 → 不算覆盖。
    """
    p = (pattern or '').upper()
    code = (code or '').upper()
    if not p or not code:
        return False
    wq = _load_weekly_module()
    if any(name in p for name in wq.COUNTRY_NAMES):
        return wq.match_warehouse(pattern, code)
    expanded = set()
    for tok in wq._split_tokens(p):
        expanded |= wq._expand_token(tok)
    return code in expanded


# ── 覆盖渠道提取 ──

def _fba_code_cell(v):
    """单元格是否为纯仓点码（如 YYZ1 / WM-LAX1）。"""
    if isinstance(v, str):
        s = v.strip().upper()
        if FBA_CODE_RE.fullmatch(s):
            return s
    return None


def _set_cov(cov, code, ch, price):
    m = cov.setdefault(code, {})
    if ch not in m or (price is not None and (m[ch] is None or price > m[ch])):
        m[ch] = price


def _clean_channel_label(v):
    """「仓库代码」表头旁的渠道名：取首行、去尾部标点（如「洛杉矶普船28日达直送-」）。"""
    s = str(v or '').split('\n')[0].strip()
    return re.sub(r'[-（(、，,：: ·]+$', '', s)


def _parse_price_cell(v):
    """解析 KG 单价，兼容「6.5/KG」「￥6.5」等格式。"""
    p = _to_kg_price(v)
    if p is not None:
        return p
    if isinstance(v, str):
        m = re.search(r'(\d+(?:\.\d+)?)', v.replace(',', ''))
        if m:
            p = float(m.group(1))
            if MIN_KG_PRICE <= p <= MAX_KG_PRICE:
                return p
    return None


def _sheet_base_name(sn):
    """sheet 名去仓点码/标点 → 语境名（如「德国DTM2.WRO5.HAJ1前置仓」→「德国前置仓」）。"""
    s = FBA_CODE_RE.sub('', sn)
    return re.sub(r'[\s\-_./·]+', '', s) or sn


# 目的仓行渠道：短运输词（海运/卡航/铁路…）
_TRANSPORT_LABELS = {"海运", "空运", "卡航", "铁路", "快递", "专车", "陆运", "海派", "空派", "快递派"}


def extract_warehouse_coverage(fp):
    """解析单个供应商报价文件 → {仓点码: {渠道名: 单价}}。

    只用两类可靠结构（避免查询参考表跨国家仓点码污染）：
      A. 「仓库代码」列（天图加拿大/美西/美转加式）：某行某列标题为「仓库代码」，
         其下行为仓点码（YYZ1/ONT8/LGB8/…），渠道名在该行右侧，价格为行内首个
         KG 单价（华南列）。
      B. sheet 名含仓点码（航乐「德国DTM2.WRO5.HAJ1前置仓」式）：该 sheet 解析出的
         渠道覆盖这些仓点（用渠道解析价）。
    """
    cov = {}
    try:
        wb = load_workbook(fp, data_only=True, read_only=True)
    except Exception:
        return cov
    try:
        for sn in wb.sheetnames:
            if _skip_sheet(sn):
                continue
            ws = wb[sn]
            rows = list(ws.iter_rows(values_only=True, max_row=MAX_SHEET_ROWS, max_col=MAX_SHEET_COLS))
            if not rows:
                continue

            # A. 「仓库代码」列 → 渠道行 × 仓点行取价
            hdr_idx = hdr_col = None
            for i, row in enumerate(rows[:40]):
                for ci, v in enumerate(row[:8]):
                    if isinstance(v, str) and '仓库代码' in v:
                        hdr_idx, hdr_col = i, ci
                        break
                if hdr_idx is not None:
                    break
            if hdr_idx is not None:
                ch_label = ''
                if hdr_col + 1 < len(rows[hdr_idx]):
                    ch_label = _clean_channel_label(rows[hdr_idx][hdr_col + 1])
                if not ch_label and hdr_idx + 1 < len(rows) and hdr_col + 1 < len(rows[hdr_idx + 1]):
                    ch_label = _clean_channel_label(rows[hdr_idx + 1][hdr_col + 1])
                for row in rows[hdr_idx + 1:]:
                    if not row or hdr_col >= len(row):
                        continue
                    code = _fba_code_cell(row[hdr_col])
                    if not code:
                        continue
                    price = None
                    for pv in row[hdr_col + 1:]:
                        p = _to_kg_price(pv)
                        if p is not None:
                            price = p
                            break
                    if ch_label and price is not None:
                        _set_cov(cov, code, ch_label, price)

            # B. sheet 名含仓点码 → 该 sheet 渠道覆盖（不覆盖 A 轮专属单价）
            name_codes = set(FBA_CODE_RE.findall(sn.upper()))
            if name_codes:
                sheet_channels = _parse_sheet_rows(rows, sn) or {}
                if sheet_channels:
                    for code in name_codes:
                        for ch, price in sheet_channels.items():
                            if ch in cov.get(code, {}):
                                continue
                            if price is not None:
                                _set_cov(cov, code, ch, price)
                else:
                    # B2. 目的仓行（航乐前置仓式）：col0=「CODE-邮编-国家」、col1=运输词、col2=KG价
                    base = _sheet_base_name(sn)
                    for row in rows:
                        if not row or not row[0] or not row[1]:
                            continue
                        m = FBA_CODE_RE.search(str(row[0]).upper())
                        if not m:
                            continue
                        label = str(row[1]).strip()
                        if label not in _TRANSPORT_LABELS:
                            continue
                        price = _parse_price_cell(row[2] if len(row) > 2 else None)
                        if price is not None:
                            _set_cov(cov, m.group(1), f"{base}-{label}", price)
    finally:
        wb.close()
    return cov


# ── 仓点查询 ──

def query_warehouses(data, country=""):
    """聚合所有已存报价 → {country: {region: [{code, channels:[{supplier,name,price}]}]}}。

    data 来自 price_supplier.load_prices()：
      {supplier: {country: {update_date, channels, warehouses?, wh_entries?}}}
    warehouses 为 extract_warehouse_coverage 的产物；wh_entries 为 JTT 每周报价原始条目。
    只展示 WAREHOUSE_POINTS 内置仓点，覆盖渠道跨供应商合并去重。
    """
    coverage = {}   # code -> {(supplier, channel): price}
    weekly_entries = []
    for sup, countries in (data or {}).items():
        for cn, p in countries.items():
            for code, chmap in (p.get("warehouses") or {}).items():
                for ch, price in chmap.items():
                    d = coverage.setdefault(code, {})
                    key = (sup, ch)
                    if key not in d or (price is not None and (d[key] is None or price > d[key])):
                        d[key] = price
            if p.get("wh_entries"):
                weekly_entries.extend(p["wh_entries"])

    pick_price = None
    if weekly_entries:
        wq = _load_weekly_module()
        pick_price = wq.pick_tier_price

    out = {}
    for cn, regions in WAREHOUSE_POINTS.items():
        if country and cn != country:
            continue
        out[cn] = {}
        for region, codes in regions.items():
            pts = []
            for code in codes:
                channels = []
                for (sup, ch), price in coverage.get(code, {}).items():
                    channels.append({"supplier": sup, "name": ch, "price": price})
                if pick_price:
                    for e in weekly_entries:
                        if not _weekly_covers(e.get("wh_pattern"), code):
                            continue
                        p = pick_price(e.get("tiers") or {}, 10 ** 6)
                        pv = round(p, 2) if p is not None else None
                        channels.append({
                            "supplier": WEEKLY_SUPPLIER_NAME,
                            "name": e.get("channel_raw") or e.get("channel"),
                            "price": pv,
                        })
                ded = {}
                for c in channels:
                    k = (c["supplier"], c["name"])
                    if k not in ded or (c["price"] is not None
                                        and (ded[k]["price"] is None or c["price"] > ded[k]["price"])):
                        ded[k] = c
                channels = sorted(ded.values(),
                                  key=lambda c: -(c["price"] if c["price"] is not None else 0))
                pts.append({"code": code, "region": region, "channels": channels})
            out[cn][region] = pts
    return out
