#!/usr/bin/env python3
"""
提单及电放保函生成模块 — 供 JTT电商AI助手 Flask 应用调用

用法（在 app.py 中）:
    from gen_bl_docs import generate_bl_docs
    zip_path = generate_bl_docs(uploaded_excel_path, output_dir)

输入要求:
    Excel 文件含 sheet "5月提单信息"，列名与 TR 退税资料明细 一致
    (JTT no. / 渠道 / 引用模板 / B/L No. / Ocean Vessel / ...)

输出:
    返回 ZIP 文件路径，内含 提单PDF + 电放保函xlsx
"""

import os, sys, io, zipfile, tempfile, shutil
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import fitz  # PyMuPDF

# ── 模板路径（相对于本模块） ──
MODULE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(MODULE_DIR, 'templates_bl')
BL_SEA = os.path.join(TEMPLATES_DIR, '提单By sea.pdf')
BL_TRAIN = os.path.join(TEMPLATES_DIR, '提单By train.pdf')
BL_TRUCK = os.path.join(TEMPLATES_DIR, '提单By truck.pdf')
TELEX = os.path.join(TEMPLATES_DIR, '电放保函.xlsx')


# ═══════════════════════════════════════════════
#  工具函数
# ═══════════════════════════════════════════════

def _safe_str(v):
    if v is None:
        return ''
    if isinstance(v, (int, float)):
        return str(int(v)) if v == int(v) else str(v)
    return str(v)


def _fmt_date(d):
    if isinstance(d, datetime):
        return d
    if isinstance(d, (int, float)):
        return datetime(1899, 12, 30) + timedelta(days=d)
    return d


def _sanitize(s):
    s = str(s).replace('/', '_').replace('\\', '_').replace(':', '_')
    s = s.replace('*', '_').replace('?', '_').replace('"', '_')
    s = s.replace('<', '_').replace('>', '_').replace('|', '_')
    s = s.replace(' ', '_')
    return s


# ═══════════════════════════════════════════════
#  数据提取
# ═══════════════════════════════════════════════

def _load_shipments(excel_path):
    """从 Excel 的 '5月提单信息' sheet 加载数据"""
    wb = openpyxl.load_workbook(excel_path, data_only=True)

    # 优先找含"提单"的 sheet，其次含"5月"+"提单"的
    sheet_name = None
    for sn in wb.sheetnames:
        if '提单' in sn:
            sheet_name = sn
            break
    if not sheet_name:
        for sn in wb.sheetnames:
            if '5月' in sn or '提单' in sn:
                sheet_name = sn
                break
    if not sheet_name:
        sheet_name = wb.sheetnames[0]  # 默认第一个

    ws = wb[sheet_name]

    headers = []
    for c in list(ws.iter_rows(min_row=1, max_row=1))[0]:
        headers.append(c.value)

    shipments = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        d = dict(zip(headers, row))
        jtt_no = _safe_str(d.get('JTT no.', '')).strip()
        template = _safe_str(d.get('引用模板', '')).strip()
        if not jtt_no or not template:
            continue
        # 跳过查验
        if _safe_str(d.get('Place of receipt', '')).strip() == '查验':
            continue
        bl_no = _safe_str(d.get('B/L No.', '')).strip()
        if not bl_no:
            continue
        shipments.append(d)

    wb.close()
    return shipments


def _data_fns():
    """数据提取函数字典"""
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
        'pdi_date': lambda s: (_fmt_date(s.get('Place and date of issue', '')).strftime('%Y/%m/%d')
                               if isinstance(_fmt_date(s.get('Place and date of issue', '')), datetime)
                               else _safe_str(s.get('Place and date of issue', ''))),
        'ob_date':  lambda s: (_fmt_date(s.get('on board date', '')).strftime('%Y/%m/%d')
                               if isinstance(_fmt_date(s.get('on board date', '')), datetime)
                               else _safe_str(s.get('on board date', ''))),
    }


# ═══════════════════════════════════════════════
#  电放保函生成
# ═══════════════════════════════════════════════

def _gen_telex(shipment, out_dir):
    if not os.path.exists(TELEX):
        return None

    jtt_no = _safe_str(shipment.get('JTT no.', ''))
    channel = _safe_str(shipment.get('渠道', ''))
    cartons = shipment.get('cartons', 0) or 0
    bl_no = _safe_str(shipment.get('B/L No.', ''))
    vessel = _safe_str(shipment.get('Ocean Vessel', ''))
    voy = _safe_str(shipment.get('Voy.No', ''))
    container = _safe_str(shipment.get('Container no.', ''))
    collect_date = _fmt_date(shipment.get('collect', ''))
    shipper_name = _safe_str(shipment.get('Shipper', '')).split('\n')[0].strip()
    consignee_name = _safe_str(shipment.get('Consignee', '')).split('\n')[0].strip()

    fname = f'{jtt_no}{_sanitize(channel)}{cartons}电放保函.xlsx'
    out_path = os.path.join(out_dir, fname)

    wb = openpyxl.load_workbook(TELEX)
    ws = wb['sheet1']

    ws['E10'] = bl_no
    ws['E12'] = f'{vessel}/{voy}' if voy else vessel
    ws['E14'] = container
    ws['A16'] = f'Shipper （发货人）                   :      {shipper_name}'
    ws['A18'] = f'Consignee （收货人）              :    {consignee_name}'

    if isinstance(collect_date, datetime):
        m, d, y = collect_date.month, collect_date.day, collect_date.year
    else:
        m, d, y = '5', '30', '2026'
    ws['C31'] = m
    ws['D31'] = d
    ws['E31'] = y

    wb.save(out_path)
    wb.close()
    return out_path


# ═══════════════════════════════════════════════
#  提单 PDF 生成
# ═══════════════════════════════════════════════

def _get_bl_template(template_type):
    mapping = {
        'by sea':  BL_SEA,
        'by train': BL_TRAIN,
        'by truck': BL_TRUCK,
    }
    return mapping.get(template_type.lower().strip())


def _sea_train_fields(is_train=False):
    """By sea / By train 模板的清除矩形 + 文本插入点"""
    clear_rects = [
        (421, 55, 509, 66),
        (156, 288, 249, 299),
        (52, 315, 154, 326),
        (156, 315, 249, 326),
        (52, 343, 158, 353),
        (156, 343, 268, 353),
        (52, 535, 118, 545),
        (52, 384, 117, 531),
        (121, 384, 208, 531),
        (212, 384, 382, 531),
        (386, 384, 455, 531),
        (459, 384, 511, 531),
        (426, 618, 473, 629),
        (395, 645, 455, 654),
        (164, 697, 213, 708),
    ]
    if is_train:
        clear_rects.append((52, 314, 249, 326))

    text_inserts = [
        ((428, 67),   'bl_no',    9),
        ((161, 298),  'place_rcpt', 11),
        ((52, 322),   'vessel',   8),
        ((161, 325),  'port_load', 11),
        ((52, 352),   'port_disc', 11),
        ((161, 352),  'place_delv', 11),
        ((56, 544),   'container', 10),
        ((77, 395),   'marks',    9),
        ((130, 395),  'cartons',  9),
        ((218, 395),  'desc',     9),
        ((382, 395),  'kgs',      9),
        ((459, 395),  'cbm',      9),
        ((428, 628),  'ob_date',  9),
        ((397, 654),  'pdi_date', 9),
        ((166, 706),  'ob_date',  9),
    ]
    return clear_rects, text_inserts


def _truck_fields():
    """By truck 模板的清除矩形 + 文本插入点"""
    clear_rects = [
        (480, 14, 560, 32),
        (35, 288, 310, 310),
        (35, 310, 310, 340),
        (35, 355, 590, 455),
        (375, 665, 460, 695),
        (180, 718, 240, 750),
    ]
    text_inserts = [
        ((488, 18),   'bl_no',    10),
        ((37, 292),   'vessel',   9),
        ((202, 292),  'port_load', 9),
        ((37, 316),   'port_disc', 9),
        ((202, 316),  'place_delv', 9),
        ((42, 360),   'container', 8),
        ((46, 362),   'marks',    9),
        ((298, 362),  'desc',     8),
        ((435, 362),  'kgs',      8),
        ((435, 372),  'cbm',      8),
        ((380, 683),  'pdi_date', 9),
        ((182, 739),  'ob_date',  9),
    ]
    return clear_rects, text_inserts


def _gen_bl(shipment, out_dir):
    template_type = _safe_str(shipment.get('引用模板', ''))
    template_path = _get_bl_template(template_type)
    if not template_path or not os.path.exists(template_path):
        return None

    jtt_no = _safe_str(shipment.get('JTT no.', ''))
    channel = _safe_str(shipment.get('渠道', ''))
    cartons = shipment.get('cartons', 0) or 0
    fname = f'{jtt_no}{_sanitize(channel)}{cartons}提单.pdf'
    out_path = os.path.join(out_dir, fname)

    F = _data_fns()
    tt = template_type.lower().strip()

    if tt in ('by sea', 'by train'):
        clear_rects, inserts = _sea_train_fields(tt == 'by train')
    elif tt == 'by truck':
        clear_rects, inserts = _truck_fields()
    else:
        return None

    try:
        doc = fitz.open(template_path)
        page = doc[0]

        # 阶段1：清除旧数据
        for rect in clear_rects:
            page.add_redact_annot(fitz.Rect(*rect), fill=(1, 1, 1))
        page.apply_redactions()

        # 阶段2：写入新数据
        for pt, field_key, fontsize in inserts:
            text = F[field_key](shipment)
            if text:
                page.insert_text(fitz.Point(*pt), text, fontsize=fontsize,
                                 fontname='helv', color=(0, 0, 0))

        doc.save(out_path, garbage=4, deflate=True)
        doc.close()
        return out_path
    except Exception:
        return None


# ═══════════════════════════════════════════════
#  主入口
# ═══════════════════════════════════════════════

def generate_bl_docs(excel_path, output_dir=None):
    """
    生成提单 + 电放保函 ZIP

    参数:
        excel_path: 上传的 Excel 文件路径（含 '5月提单信息' sheet）
        output_dir: 输出目录，None 则自动创建临时目录

    返回:
        zip_path: 生成的 ZIP 文件路径（内含所有 提单PDF + 电放保函xlsx）
    """
    if output_dir is None:
        output_dir = tempfile.mkdtemp(prefix='bl_docs_')
    else:
        os.makedirs(output_dir, exist_ok=True)

    # 加载数据
    shipments = _load_shipments(excel_path)

    if not shipments:
        raise ValueError("Excel 中未找到有效的提单数据（需要 '5月提单信息' sheet，且含有效的 B/L No.）")

    # 生成所有文件到临时目录
    telex_ok = bl_ok = 0
    for s in shipments:
        if _gen_telex(s, output_dir):
            telex_ok += 1
        if _gen_bl(s, output_dir):
            bl_ok += 1

    if telex_ok == 0 and bl_ok == 0:
        raise ValueError("未能生成任何文件，请检查数据格式")

    # 打包为 ZIP
    zip_name = f'提单电放保函_{datetime.now().strftime("%Y%m%d_%H%M%S")}.zip'
    zip_path = os.path.join(output_dir, zip_name)

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for fname in sorted(os.listdir(output_dir)):
            fpath = os.path.join(output_dir, fname)
            if fname.endswith('.pdf') or fname.endswith('.xlsx'):
                if fname != zip_name:
                    zf.write(fpath, fname)

    return zip_path, telex_ok, bl_ok
