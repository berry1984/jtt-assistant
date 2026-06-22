#!/usr/bin/env python3
"""
生成提单PDF和电放保函xlsx
根据 TR 退税资料明细.xlsx 的"5月提单信息"页面，对应生成提单和电放保函

规则：
- 引用模板：By sea → 提单By sea.pdf, By train → 提单By train.pdf, By truck → 提单By truck.pdf
- 跳过 Place of receipt = "查验" 的货件
- 文件名格式：JTT号+渠道+箱数+提单/电放保函.后缀
"""

import os, sys
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import fitz  # PyMuPDF

# ── 路径 ──
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, '模版', '模版')
DATA_FILE = os.path.join(BASE_DIR, 'TR 退税资料明细.xlsx')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output_5月')

BL_SEA_TEMPLATE = os.path.join(TEMPLATE_DIR, '提单By sea.pdf')
BL_TRUCK_TEMPLATE = os.path.join(TEMPLATE_DIR, '提单By truck.pdf')
BL_TRAIN_TEMPLATE = os.path.join(TEMPLATE_DIR, '提单By train.pdf')
TELEX_TEMPLATE = os.path.join(TEMPLATE_DIR, '电放保函.xlsx')

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ── 固定信息 ──
SHIPPER = ("Guangzhou Tuorui Technology Co.,Ltd\n"
           "Room 411,No.101 Dexing Road Wanggang Jiahe Street\n"
           "Baiyun District, Guangzhou")

CONSIGNEE = ("Hong Kong Lixiang Trading Company Limited\n"
             "FLAT/RM 3B 3/F BANK TOWER NOS.351&353 KING'S\n"
             "ROAD NORTH POINT HK")


# ── 工具函数 ──
def _safe_str(v):
    if v is None:
        return ''
    if isinstance(v, (int, float)):
        return str(int(v)) if v == int(v) else str(v)
    return str(v)


def fmt_date(d):
    if isinstance(d, datetime):
        return d
    if isinstance(d, (int, float)):
        return datetime(1899, 12, 30) + timedelta(days=d)
    return d


def sanitize_filename(s):
    s = str(s).replace('/', '_').replace('\\', '_').replace(':', '_')
    s = s.replace('*', '_').replace('?', '_').replace('"', '_')
    s = s.replace('<', '_').replace('>', '_').replace('|', '_')
    s = s.replace(' ', '_')
    return s


def load_shipments():
    wb = openpyxl.load_workbook(DATA_FILE, data_only=True)
    ws = wb['5月提单信息']
    headers = []
    for c in list(ws.iter_rows(min_row=1, max_row=1))[0]:
        headers.append(c.value)
    shipments = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        d = dict(zip(headers, row))
        jtt_no = _safe_str(d.get('JTT no.', ''))
        template = _safe_str(d.get('引用模板', ''))
        if not jtt_no or not template:
            continue
        if _safe_str(d.get('Place of receipt', '')).strip() == '查验':
            print(f'  ⏭ 跳过查验: {jtt_no}')
            continue
        bl_no = _safe_str(d.get('B/L No.', '')).strip()
        if not bl_no:
            print(f'  ⏭ 跳过无提单号: {jtt_no}')
            continue
        shipments.append(d)
    wb.close()
    return shipments


def _get_template_path(template_type):
    mapping = {
        'by sea': BL_SEA_TEMPLATE,
        'by truck': BL_TRUCK_TEMPLATE,
        'by train': BL_TRAIN_TEMPLATE,
    }
    return mapping.get(template_type.lower().strip())


# ============================================================
# 1. 电放保函 (Telex Release) — .xlsx
# ============================================================
def generate_telex(shipment):
    if not os.path.exists(TELEX_TEMPLATE):
        print(f'  ❌ 找不到电放保函模板')
        return None
    jtt_no = _safe_str(shipment.get('JTT no.', ''))
    channel = _safe_str(shipment.get('渠道', ''))
    cartons = shipment.get('cartons', 0) or 0
    bl_no = _safe_str(shipment.get('B/L No.', ''))
    vessel = _safe_str(shipment.get('Ocean Vessel', ''))
    voy = _safe_str(shipment.get('Voy.No', ''))
    container = _safe_str(shipment.get('Container no.', ''))
    collect_date = fmt_date(shipment.get('collect', ''))
    # 公司名只取第一行，不要地址
    shipper_name = _safe_str(shipment.get('Shipper', '')).split('\n')[0].strip()
    consignee_name = _safe_str(shipment.get('Consignee', '')).split('\n')[0].strip()

    fname = f'{jtt_no}{sanitize_filename(channel)}{cartons}电放保函.xlsx'
    out_path = os.path.join(OUTPUT_DIR, fname)

    wb = openpyxl.load_workbook(TELEX_TEMPLATE)
    ws = wb['sheet1']

    # 只写值，不改任何格式（字体/对齐/边框保留模板原样）
    ws['E10'] = bl_no
    vessel_str = f'{vessel}/{voy}' if voy else vessel
    ws['E12'] = vessel_str
    ws['E14'] = container

    # Shipper/Consignee — 保留模板原有标签前缀，只换公司名
    ws['A16'] = f'Shipper （发货人）                   :      {shipper_name}'
    ws['A18'] = f'Consignee （收货人）              :    {consignee_name}'

    # 日期
    if isinstance(collect_date, datetime):
        m, d, y = collect_date.month, collect_date.day, collect_date.year
    else:
        m, d, y = '5', '30', '2026'
    ws['C31'] = m
    ws['D31'] = d
    ws['E31'] = y

    wb.save(out_path)
    wb.close()
    print(f'  ✅ 电放保函: {fname}')
    return out_path
    print(f'  ✅ 电放保函: {fname}')
    return out_path


# ============================================================
# 2. 提单 (Bill of Lading) — .pdf
# ============================================================
#
# 两阶段方法：
#   阶段1：用清除矩形覆盖旧数据区域 → 红划删除（白色填充）
#   阶段2：在指定插入点写入新文本
#
# 数据提取函数（接收 shipment 字典，返回文本）
# -----------------------------------------------------------

def _data_fns():
    """返回数据提取函数字典"""
    return {
        'bl_no':    lambda s: _safe_str(s.get('B/L No.', '')),
        'vessel':   lambda s: (_safe_str(s.get('Ocean Vessel', '')) + '/' +
                               _safe_str(s.get('Voy.No', ''))) if s.get('Voy.No') else _safe_str(s.get('Ocean Vessel', '')),
        'place_rcpt': lambda s: _safe_str(s.get('Place of receipt', '')),
        'port_load':  lambda s: _safe_str(s.get('Port of loading', '')),
        'port_disc':  lambda s: _safe_str(s.get('Port of discharge', '')),
        'place_delv': lambda s: _safe_str(s.get('Place of delivery', '')),
        'container':  lambda s: _safe_str(s.get('Container no.', '')),
        'marks':    lambda s: _safe_str(s.get('Marks & No.', '')) or 'N/M',
        'cartons':  lambda s: str(int(s.get('cartons', 0) or 0)) + ' CARTONS',
        'desc':     lambda s: _safe_str(s.get('Description of goods', '')),
        'kgs':      lambda s: (_safe_str(s.get('KGS', 0) or 0) + ' KGS') if s.get('KGS') else '',
        'cbm':      lambda s: (_safe_str(s.get('CBM', 0) or 0) + ' CBM') if s.get('CBM') else '',
        'pdi_date': lambda s: (fmt_date(s.get('Place and date of issue', '')).strftime('%Y/%m/%d')
                               if isinstance(fmt_date(s.get('Place and date of issue', '')), datetime)
                               else _safe_str(s.get('Place and date of issue', ''))),
        'ob_date':  lambda s: (fmt_date(s.get('on board date', '')).strftime('%Y/%m/%d')
                               if isinstance(fmt_date(s.get('on board date', '')), datetime)
                               else _safe_str(s.get('on board date', ''))),
    }


def _redact_and_insert(page, clear_rects, text_inserts, shipment):
    """
    clear_rects: [(x0, y0, x1, y1), ...]  — 红划删除区域
    text_inserts: [((x, y), field_key, fontsize), ...] — 插入点 + 字段名
    """
    F = _data_fns()
    # 阶段1：删除旧数据
    for rect in clear_rects:
        if not rect:
            continue
        page.add_redact_annot(fitz.Rect(*rect), fill=(1, 1, 1))
    page.apply_redactions()
    # 阶段2：写入新数据
    for pt, field_key, fontsize in text_inserts:
        text = F[field_key](shipment)
        if not text:
            continue
        page.insert_text(fitz.Point(*pt), text, fontsize=fontsize,
                         fontname='helv', color=(0, 0, 0))


def sea_train_fields(is_train=False):
    """
    By sea / By train 模板的清除矩形 + 文本插入点

    蓝色边框线位置（By sea, 左栏 x=50-251）:
      水平线: y=272, 286, 300, 313, 327, 341, 354
      单元格按 y 区间分：
        y=272-286: 标签 "Pre-carriage by"
        y=286-300: ⭐ Place of receipt 数据 (161,287)-(238,298)
        y=300-313: 标签 "Ocean vessel"
        y=313-327: ⭐ Vessel 数据 (52,314)-(150,322)
        y=327-341: 标签 "Port of discharge"
        y=341-354: ⭐ Port discharge 数据 (52,342)-(153,352)
      右栏 x=251-513: 整个区域 y=272-354 为一个单元格

    货物表格（x=50-513, y=354-547）：
      水平线: y=354(表头底), y=382(数据行顶), y=533/546(底)
      垂直线: x=119, 210, 384, 457
    """
    # ── 清除矩形 — 精确到每个单元格，距蓝色边框 2px ──
    clear_rects = [
        # B/L No（蓝线 y=54，旧数据 y=54-64）
        (421, 55, 509, 66),
        # Place of receipt（单元格 y=286-300，旧数据 y=287-298）
        (156, 288, 249, 299),
        # Vessel（单元格 y=313-327，旧数据 y=314-322）
        (52, 315, 154, 326),
        # Port loading（同单元格 y=313-327，旧数据 y=314-325）
        (156, 315, 249, 326),
        # Port discharge（单元格 y=341-354，旧数据 y=342-352）
        (52, 343, 158, 353),
        # Place delivery（同单元格，旧数据 y=342-352）
        (156, 343, 268, 353),
        # Container No（表格下方 y=533-546）
        (52, 535, 118, 545),
        # ── 货物表逐列清除（蓝线 y=382 和 y=533 之间）──
        (52, 384, 117, 531),    # 列1 Marks
        (121, 384, 208, 531),   # 列2 Quantity
        (212, 384, 382, 531),   # 列3 Description
        (386, 384, 455, 531),   # 列4 KGS
        (459, 384, 511, 531),   # 列5 CBM
        # 底部日期（避开蓝色文字和蓝线）
        #   SHIPPED ON BOARD 标签 y=605-616, 蓝线 y=630, 旧日期 y=618-630
        (426, 618, 473, 629),
        #   Place and Date of Issue 蓝色标签 y=635-644, 蓝线 y=656
        (395, 645, 455, 654),
        #   On board date 蓝色标签 y=686-696, 底部蓝线 y=709
        (164, 697, 213, 708),
    ]
    # Train 有额外 "CHONGQING" 文字在 Vessel 区域
    if is_train:
        clear_rects.append((52, 314, 249, 326))

    # ── 文字写入位置 — baseline 与模板原始数据对齐 ──
    # 模板原始数据 bbox:
    #   B/L No "CHN3261876P5"  bbox y=54-64   → baseline y=64
    #   Place of receipt       bbox y=287-298 → baseline y=298
    #   Vessel                 bbox y=314-322 → baseline y=322
    #   Port loading           bbox y=314-325 → baseline y=325
    #   Port discharge         bbox y=342-352 → baseline y=352
    #   Place delivery         bbox y=342-352 → baseline y=352
    #   Container No           bbox y=533-544 → baseline y=544
    #   表格第一行             bbox y=383-395 → baseline y=395
    text_inserts = [
        ((428, 67),   'bl_no',    9),
        ((161, 298),  'place_rcpt', 11),
        ((52, 322),   'vessel',   8),
        ((161, 325),  'port_load', 11),
        ((52, 352),   'port_disc', 11),
        ((161, 352),  'place_delv', 11),
        ((56, 544),   'container', 10),
        # 货物表 — baseline 对齐模板 y=395
        ((77, 395),   'marks',    9),
        ((130, 395),  'cartons',  9),
        ((218, 395),  'desc',     9),
        ((382, 395),  'kgs',      9),
        ((459, 395),  'cbm',      9),
        # 底部日期
        ((428, 628),  'ob_date',  9),
        ((397, 654),  'pdi_date', 9),
        ((166, 706),  'ob_date',  9),
    ]

    return clear_rects, text_inserts


def truck_fields():
    """By truck 模板的清除矩形 + 文本插入点"""
    clear_rects = [
        (480, 14, 560, 32),        # B/L No
        (35, 288, 310, 310),       # Vessel + Place loading
        (35, 310, 310, 340),       # Place discharge + delivery
        (35, 355, 590, 455),       # 货物表整区
        (375, 665, 460, 695),      # Place/Date of Issue
        (180, 718, 240, 750),      # Date of departure
    ]
    text_inserts = [
        ((488, 18),    'bl_no',    10),
        ((37, 292),    'vessel',   9),
        ((202, 292),   'port_load', 9),
        ((37, 316),    'port_disc', 9),
        ((202, 316),   'place_delv', 9),
        ((42, 360),    'container', 8),
        ((46, 362),    'marks',    9),
        ((298, 362),   'desc',     8),
        ((435, 362),   'kgs',      8),
        ((435, 372),   'cbm',      8),
        ((380, 683),   'pdi_date', 9),
        ((182, 739),   'ob_date',  9),
    ]
    return clear_rects, text_inserts


def generate_bl(shipment):
    template_type = _safe_str(shipment.get('引用模板', ''))
    template_path = _get_template_path(template_type)
    if not template_path or not os.path.exists(template_path):
        print(f'  ⚠ 找不到提单模板: {template_type}')
        return None

    jtt_no = _safe_str(shipment.get('JTT no.', ''))
    channel = _safe_str(shipment.get('渠道', ''))
    cartons = shipment.get('cartons', 0) or 0
    fname = f'{jtt_no}{sanitize_filename(channel)}{cartons}提单.pdf'
    out_path = os.path.join(OUTPUT_DIR, fname)

    try:
        doc = fitz.open(template_path)
        page = doc[0]
        tt = template_type.lower().strip()

        if tt in ('by sea', 'by train'):
            is_train = (tt == 'by train')
            clear_rects, inserts = sea_train_fields(is_train)
        elif tt == 'by truck':
            clear_rects, inserts = truck_fields()
        else:
            print(f'  ⚠ 未知模板类型: {template_type}')
            doc.close()
            return None

        _redact_and_insert(page, clear_rects, inserts, shipment)
        doc.save(out_path, garbage=4, deflate=True)
        doc.close()
        print(f'  ✅ 提单: {fname}')
        return out_path
    except Exception as e:
        import traceback
        print(f'  ❌ 提单生成失败 {jtt_no}: {e}')
        traceback.print_exc()
        return None


# ============================================================
# Main
# ============================================================
def main():
    print(f'📂 数据文件: {DATA_FILE}')
    print(f'📂 模板目录: {TEMPLATE_DIR}')
    print(f'📂 输出目录: {OUTPUT_DIR}\n')

    shipments = load_shipments()
    print(f'📋 共 {len(shipments)} 票待生成\n')

    telex_ok = bl_ok = 0
    for i, s in enumerate(shipments, 1):
        jtt_no = _safe_str(s.get('JTT no.', ''))
        channel = _safe_str(s.get('渠道', ''))
        template = _safe_str(s.get('引用模板', ''))
        cartons = s.get('cartons', 0) or 0
        print(f'[{i:02d}] {jtt_no} | {channel} | {template} | {cartons}箱')

        if generate_telex(s):
            telex_ok += 1
        if generate_bl(s):
            bl_ok += 1
        print()

    print(f'={"=" * 50}')
    print(f'📊 完成统计:')
    print(f'   ✅ 电放保函: {telex_ok}/{len(shipments)}')
    print(f'   ✅ 提单:     {bl_ok}/{len(shipments)}')
    print(f'   📁 输出目录: {OUTPUT_DIR}')


if __name__ == '__main__':
    main()
