"""
供应商报价持久化查询模块
========================
固定供应商 + 自定义供应商，按「供应商 × 国家」槽位保存报价文件（再次上传即覆盖更新），
解析渠道价格，按 供应商 → 国家 → 渠道 组织，支持按国家 / 供应商 / 渠道关键词搜索。

存储目录（云上挂载 Railway Volume 于 /data，本地开发兜底 price_data/）：
    <STORAGE_DIR>/suppliers/<供应商slug>/supplier.json      # 显示名记录
    <STORAGE_DIR>/suppliers/<供应商slug>/<国家>/<原文件名>.xlsx   # 每个槽位一个文件

文件保留原文件名（内含生效日期），再次上传同槽位先清旧文件再写新文件 = 覆盖更新。
不依赖 JTT 成本模板，纯供应商报价查询。
"""

import os
import re
import json
import shutil
import sys
from datetime import datetime

THIS_DIR = os.path.dirname(os.path.abspath(__file__))
# app.py 会把 /Users/admin/报价工具 插到 sys.path[0]（仅供 ym_cost），其下也有旧的
# supplier_parser.py，会遮蔽本目录的版本。这里把本目录重新置顶，保证用本地增强版。
if THIS_DIR not in sys.path:
    sys.path.insert(0, THIS_DIR)
elif sys.path[0] != THIS_DIR:
    sys.path.remove(THIS_DIR)
    sys.path.insert(0, THIS_DIR)

from supplier_parser import parse_supplier_files, _extract_date_from_filename, classify_transport
from warehouse_points import (WEEKLY_SUPPLIER_NAME, WEEKLY_COUNTRY,
                              extract_warehouse_coverage, _load_weekly_module)

FIXED_SUPPLIERS = ["天图/心一", "英美", "美琦", "凯鑫", "航乐", "皓鹏"]
COUNTRIES = ["美国", "加拿大", "欧洲", "英国"]

# 非渠道行关键词（地址、仓库代码、备注、噪声等）
_NON_CHANNEL_KW = [
    "含私人地址", "住宅费", "谷仓", "4PX", "希音", "FBA", "warehouse",
    "最长边", "围长", "尺寸", "CM", "KG", "邮编", "Zip",
    "LBA", "BHX", "DTM", "WRO", "HAJ", "YVR", "YOW", "YYZ", "YEG", "YYC",
    "OAKVILLE", "CDISCOUNT", "match",
    # 英文/西语街道地址噪声（美琦等文件含收货地址行）
    "calle", "carretera", "avenida", "street", "avenue", "boulevard",
    "road", "suite", "unit", "house", "drive", "lane", "km",
]

# 供应商显示名 → 目录 slug（去掉路径非法字符，保证「天图/心一」安全落盘）
def _slug(name):
    return re.sub(r'[\\/:*?"<>|]', '-', (name or '').strip())


# 解析缓存（大文件如天图/英美解析要十几秒，按槽位文件 mtime+size 失效）
_cache = {"sig": None, "data": None}


def _build_sig():
    root = os.path.join(get_storage_dir(), "suppliers")
    sig = []
    if os.path.isdir(root):
        for slug in sorted(os.listdir(root)):
            sdir = os.path.join(root, slug)
            if not os.path.isdir(sdir):
                continue
            for country in sorted(os.listdir(sdir)):
                cdir = os.path.join(sdir, country)
                if not os.path.isdir(cdir):
                    continue
                for fn in os.listdir(cdir):
                    if fn.lower().endswith((".xlsx", ".xls")) and not fn.startswith("~$"):
                        fp = os.path.join(cdir, fn)
                        st = os.stat(fp)
                        sig.append((os.path.join(slug, country, fn), st.st_mtime, st.st_size))
    return tuple(sorted(sig))


def get_storage_dir():
    """解析持久化根目录：STORAGE_DIR 环境变量 > /data(卷) > 本地 price_data/。"""
    d = os.environ.get("STORAGE_DIR", "").strip()
    if not d:
        candidate = "/data"
        if not (os.path.exists(candidate) and os.access(candidate, os.W_OK)):
            candidate = os.path.join(THIS_DIR, "price_data")
        d = candidate
    return d


def _sup_dir(supplier):
    return os.path.join(get_storage_dir(), "suppliers", _slug(supplier))


def _write_supplier_meta(supplier):
    """记录供应商显示名，便于从 slug 目录还原。"""
    d = _sup_dir(supplier)
    os.makedirs(d, exist_ok=True)
    try:
        with open(os.path.join(d, "supplier.json"), "w", encoding="utf-8") as f:
            json.dump({"supplier": supplier}, f, ensure_ascii=False)
    except Exception:
        pass


def _read_supplier_meta(slug):
    try:
        with open(os.path.join(get_storage_dir(), "suppliers", slug, "supplier.json"),
                  "r", encoding="utf-8") as f:
            return json.load(f).get("supplier", slug)
    except Exception:
        return slug


def _read_meta(supplier, country):
    mp = os.path.join(_sup_dir(supplier), country, "meta.json")
    if os.path.isfile(mp):
        try:
            with open(mp, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def _write_meta(supplier, country, meta):
    d = os.path.join(_sup_dir(supplier), country)
    os.makedirs(d, exist_ok=True)
    mp = os.path.join(d, "meta.json")
    try:
        with open(mp, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def _find_slot_file(supplier, country):
    """槽位内唯一报价文件（每个槽位只保留一份 .xlsx）。"""
    d = os.path.join(_sup_dir(supplier), country)
    if not os.path.isdir(d):
        return None
    for fn in os.listdir(d):
        if fn.lower().endswith((".xlsx", ".xls")) and not fn.startswith("~$"):
            return os.path.join(d, fn)
    return None


def _is_real_channel(name):
    """过滤地址 / 仓库代码 / 备注 / 噪声行，只留真实渠道。"""
    if not name:
        return False
    name_lower = name.lower()
    for kw in _NON_CHANNEL_KW:
        if kw.lower() in name_lower:
            return False
    if re.fullmatch(r"[A-Z0-9]{3,8}", name):
        return False
    if "," in name and any(c.isdigit() for c in name[:5]):
        return False
    if len(name) > 80 or name.startswith(("1、", "2、", "d)", "a)")):
        return False
    return True


def _parse_file(fp):
    """解析单个报价文件 → {update_date, channels:[{name, price, mode}], warehouses, wh_postals}，去噪声、按名排序。"""
    p = parse_supplier_files([fp], region_filter=False)
    if not p:
        return None
    ch_sheets = p.get("channel_sheets") or {}
    chs = []
    for name, price in p.get("channels", {}).items():
        if not _is_real_channel(name):
            continue
        try:
            pv = round(float(price), 2)
        except (ValueError, TypeError):
            pv = None
        chs.append({"name": name, "price": pv,
                    "mode": classify_transport(name, ch_sheets.get(name, ""))})
    chs.sort(key=lambda c: c["name"])
    update_date = p.get("update_date", "") or ""
    if not update_date:
        update_date = _extract_date_from_filename(fp) or ""
    cov, postals = extract_warehouse_coverage(fp)
    return {"update_date": update_date, "channels": chs,
            "warehouses": cov, "wh_postals": postals}


def _parse_weekly_slot(fp):
    """解析 JTT每周报价表（格式不同，走 weekly_quotation）→ {update_date, channels, wh_entries}。"""
    wq = _load_weekly_module()
    entries = wq.parse_weekly_quotation(fp)
    if not entries:
        return None
    channels = []
    seen = {}
    for e in entries:
        key = (e.get('section'), e.get('channel_raw') or e.get('channel'))
        price = wq.pick_tier_price(e.get("tiers") or {}, 10 ** 6)
        pv = round(price, 2) if price is not None else None
        if key in seen:
            if pv is not None and (seen[key]['price'] is None or pv > seen[key]['price']):
                seen[key]['price'] = pv
            continue
        name = e.get('channel_raw') or e.get('channel') or ''
        d = {"name": name, "price": pv,
             "section": e.get('section', ''),
             "mode": classify_transport(name, e.get('section', ''))}
        seen[key] = d
        channels.append(d)
    channels.sort(key=lambda c: c["name"])
    update_date = _extract_date_from_filename(fp) or ""
    return {"update_date": update_date, "channels": channels, "wh_entries": entries}


def list_custom_suppliers():
    """自定义供应商 = FIXED_SUPPLIERS 之外、目录实际存在的供应商（显示名）。"""
    root = os.path.join(get_storage_dir(), "suppliers")
    if not os.path.isdir(root):
        return []
    out = []
    for slug in sorted(os.listdir(root)):
        if os.path.isdir(os.path.join(root, slug)):
            display = _read_supplier_meta(slug)
            if display not in FIXED_SUPPLIERS and display != WEEKLY_SUPPLIER_NAME and display not in out:
                out.append(display)
    return sorted(out)


def list_slots():
    """全部已上传槽位（用于「已存报价清单」）。"""
    root = os.path.join(get_storage_dir(), "suppliers")
    slots = []
    if not os.path.isdir(root):
        return slots
    for slug in sorted(os.listdir(root)):
        sdir = os.path.join(root, slug)
        if not os.path.isdir(sdir):
            continue
        display = _read_supplier_meta(slug)
        for country in sorted(os.listdir(sdir)):
            is_weekly = display == WEEKLY_SUPPLIER_NAME and country == WEEKLY_COUNTRY
            if country in COUNTRIES or is_weekly:
                fp = _find_slot_file(display, country)
                if fp:
                    meta = _read_meta(display, country)
                    # 用 meta 直出，避免每个槽位重复解析大文件；仅旧槽位无 meta 时才回退解析
                    if not meta.get("channel_count") or not meta.get("update_date"):
                        parsed = (_parse_weekly_slot(fp) if is_weekly else _parse_file(fp)) \
                            or {"update_date": "", "channels": []}
                        meta = dict(meta, update_date=parsed["update_date"],
                                    channel_count=len(parsed["channels"]),
                                    uploaded_at=meta.get("uploaded_at", ""))
                        # 回写 meta：中断的上传（有文件无 meta）只解析一次，避免每次重复全量解析
                        _write_meta(display, country, meta)
                    slots.append({
                        "supplier": display,
                        "country": "每周" if is_weekly else country,
                        "source_filename": os.path.basename(fp),
                        "uploaded_at": meta.get(
                            "uploaded_at",
                            datetime.fromtimestamp(os.path.getmtime(fp)).strftime("%Y-%m-%d %H:%M:%S"),
                        ),
                        "update_date": meta.get("update_date", ""),
                        "channel_count": meta.get("channel_count", ""),
                        "custom": display not in FIXED_SUPPLIERS and not is_weekly,
                    })
    return slots


def save_upload(supplier, countries, file):
    """上传报价文件到 (supplier × 多国) 槽位；再次上传同槽位即覆盖。

    注意：FileStorage 流只能 save 一次，其余国家用 shutil.copy 复制。
    """
    supplier = (supplier or "").strip()
    is_weekly = supplier == WEEKLY_SUPPLIER_NAME
    if not is_weekly and supplier not in FIXED_SUPPLIERS and supplier not in list_custom_suppliers():
        return {"error": f"供应商「{supplier}」不存在，请先用「新增供应商」创建"}
    if not file or file.filename == "":
        return {"error": "请选择报价文件"}
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        return {"error": "文件需要是 .xlsx / .xls"}
    if not countries and not is_weekly:
        return {"error": "请至少选择一个国家"}

    _write_supplier_meta(supplier)
    source_filename = file.filename
    saved = []
    first_fp = None
    target_countries = [WEEKLY_COUNTRY] if is_weekly else [c for c in countries if c in COUNTRIES]
    for country in target_countries:
        d = os.path.join(_sup_dir(supplier), country)
        os.makedirs(d, exist_ok=True)
        for fn in os.listdir(d):
            if fn.lower().endswith((".xlsx", ".xls")) and not fn.startswith("~$"):
                try:
                    os.remove(os.path.join(d, fn))
                except OSError:
                    pass
        fp = os.path.join(d, source_filename)
        if first_fp is None:
            file.save(fp)
            first_fp = fp
        else:
            shutil.copy(first_fp, fp)
        saved.append(country)

    # 只解析一次（多国是同一文件），共享结果写各槽位 meta
    parsed = (_parse_weekly_slot(first_fp) if is_weekly else _parse_file(first_fp)) \
        or {"update_date": "", "channels": []}
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for country in saved:
        meta = {
            "source_filename": source_filename,
            "uploaded_at": now,
            "update_date": parsed["update_date"],
            "channel_count": len(parsed["channels"]),
        }
        _write_meta(supplier, country, meta)
    _cache["sig"] = None  # 失效查询缓存

    return {"success": True, "supplier": supplier, "countries": saved,
            "slots": list_slots()}


def add_custom_supplier(name, countries, file):
    """新增自定义供应商并上传其首个报价文件。"""
    name = (name or "").strip()
    if not name:
        return {"error": "请填写供应商名称"}
    if name in FIXED_SUPPLIERS:
        return {"error": f"「{name}」是固定供应商，直接在其上传位上传即可"}
    if not countries:
        return {"error": "请至少选择一个国家"}
    _write_supplier_meta(name)
    return save_upload(name, countries, file)


def delete_slot(supplier, country):
    """删除单个槽位；若自定义供应商目录已空则一并移除。"""
    supplier = (supplier or "").strip()
    d = os.path.join(_sup_dir(supplier), country)
    if not os.path.isdir(d):
        return {"error": "槽位不存在"}
    shutil.rmtree(d, ignore_errors=True)
    sdir = _sup_dir(supplier)
    if os.path.isdir(sdir) and not os.listdir(sdir):
        shutil.rmtree(sdir, ignore_errors=True)
    _cache["sig"] = None  # 失效查询缓存
    return {"success": True, "slots": list_slots()}


def load_prices():
    """{supplier(显示名): {country: {update_date, channels:[{name,price}]}}}

    按槽位文件 mtime+size 做缓存，上传/删除后自动失效，避免大文件反复解析。
    """
    sig = _build_sig()
    if _cache["sig"] == sig and _cache["data"] is not None:
        return _cache["data"]

    root = os.path.join(get_storage_dir(), "suppliers")
    data = {}
    if os.path.isdir(root):
        for slug in sorted(os.listdir(root)):
            sdir = os.path.join(root, slug)
            if not os.path.isdir(sdir):
                continue
            display = _read_supplier_meta(slug)
            sub = {}
            for country in sorted(os.listdir(sdir)):
                is_weekly = display == WEEKLY_SUPPLIER_NAME and country == WEEKLY_COUNTRY
                if country not in COUNTRIES and not is_weekly:
                    continue
                fp = _find_slot_file(display, country)
                if fp:
                    p = _parse_weekly_slot(fp) if is_weekly else _parse_file(fp)
                    if p:
                        sub[country] = p
            if sub:
                data[display] = sub

    _cache["sig"] = sig
    _cache["data"] = data
    return data


def query(keyword="", country="", supplier="", mode=""):
    """供应商报价查询：按 供应商 → 国家 → 渠道 组织，支持国家/供应商/关键词/运输类别过滤。"""
    kw = (keyword or "").strip().lower()
    md = (mode or "").strip()
    data = load_prices()
    result = {}
    countries_used = set()
    total_channels = 0
    mode_stats = {}
    for sup in data:
        if supplier and sup != supplier:
            continue
        sub = {}
        for cn, p in data[sup].items():
            if country and cn != country:
                continue
            chs = p["channels"]
            if kw:
                chs = [c for c in chs if kw in c["name"].lower()]
            if md:
                chs = [c for c in chs if (c.get("mode") or "未知") == md]
            # JTT 每周报价按分区（美国渠道/加拿大渠道/欧线渠道/英国渠道）过滤国家
            if country and sup == WEEKLY_SUPPLIER_NAME:
                chs = [c for c in chs if country in (c.get("section") or "")]
            if not chs:
                continue
            countries_used.add(cn)
            total_channels += len(chs)
            for c in chs:
                m = c.get("mode") or "未知"
                mode_stats[m] = mode_stats.get(m, 0) + 1
            sub[cn] = {"update_date": p["update_date"], "channels": chs}
        if sub:
            result[sup] = sub
    return {
        "suppliers": list(result.keys()),
        "countries": sorted(countries_used),
        "fixed_suppliers": FIXED_SUPPLIERS,
        "custom_suppliers": list_custom_suppliers(),
        "total_channels": total_channels,
        "mode_stats": mode_stats,
        "data": result,
    }
