#!/usr/bin/env python3
"""
投保区间拆分工具 v2

将「下单 Excel」按每箱申报价值(RMB)拆分为多个独立 Excel 文件，
每个文件对应一个投保区间，用于分别投保。

规则详见：投保区间拆分规则说明_v2.md

用法：
  python3 split_insurance_v2.py <下单发票.xlsx>
  python3 split_insurance_v2.py <下单发票.xlsx> --out-dir ./输出目录
"""

import sys
import os
import re
import argparse
import zipfile
import io
from copy import copy
from xml.etree import ElementTree as ET

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XlImage


# ═══════════════════════════════════════════════════════════════
#  常量
# ═══════════════════════════════════════════════════════════════

CURRENCY_RATES = {
    '美金': 7,
    '美元': 7,
    'USD': 7,
    '欧元': 8,
    'EUR': 8,
    '英镑': 9,
    'GBP': 9,
}

RANGES = [
    ('不足5000RMB',    0,  5000),
    ('5000-10000RMB',  5000, 10000),
    ('10000-20000RMB', 10000, 20000),
    ('20000-30000RMB', 20000, 30000),
    ('30000-40000RMB', 30000, 40000),
]

HEADER_END = 26       # Row 1-26 = 头部
COL_HEADER_ROW = 27   # Row 27 = 列名行
DATA_START_ROW = 28   # Row 28+ = 数据区
NUM_DATA_COLS = 19    # A 到 S 列（源文件数据列）
TOTAL_COLS = 20       # 输出文件含 T 列（辅助列）
IMG_COL = 16          # P 列 = 产品图片

AUX_SHEETS = ['FBA地址库编码表', '服务名称', '换算']

# 标准样式
THIN_SIDE = Side(style='thin')
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE,
                     top=THIN_SIDE, bottom=THIN_SIDE)

# WPS cellimages 命名空间
NS_PIC = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types'


# ═══════════════════════════════════════════════════════════════
#  样式辅助函数
# ═══════════════════════════════════════════════════════════════

def _copy_style(src_cell):
    """复制单元格的样式属性"""
    return {
        'font': copy(src_cell.font) if src_cell.font is not None else Font(),
        'alignment': (copy(src_cell.alignment)
                      if src_cell.alignment is not None else Alignment()),
        'border': copy(src_cell.border) if src_cell.border is not None else Border(),
        'fill': copy(src_cell.fill) if src_cell.fill is not None else PatternFill(),
        'number_format': src_cell.number_format,
    }


def _apply_style(dst_cell, style):
    """将样式应用到目标单元格"""
    dst_cell.font = copy(style['font'])
    dst_cell.alignment = copy(style['alignment'])
    dst_cell.border = copy(style['border'])
    dst_cell.fill = copy(style['fill'])
    dst_cell.number_format = style['number_format']


# ═══════════════════════════════════════════════════════════════
#  图片提取（WPS cellimages.xml 格式）
# ═══════════════════════════════════════════════════════════════

def _extract_images(src_path):
    """
    从 WPS cellimages.xml 提取嵌入图片。

    返回: {dispimg_id: image_bytes, ...}
    """
    images = {}
    try:
        with zipfile.ZipFile(src_path, 'r') as z:
            if 'xl/cellimages.xml' not in z.namelist():
                return images
            if 'xl/_rels/cellimages.xml.rels' not in z.namelist():
                return images

            # 关系映射: rId → media 图文件
            rels_root = ET.fromstring(z.read('xl/_rels/cellimages.xml.rels'))
            rid_to_target = {}
            for rel in rels_root:
                rid = rel.get('Id', '')
                target = rel.get('Target', '')
                if rid and target:
                    rid_to_target[rid] = target

            # 解析 cellimages.xml
            ci_root = ET.fromstring(z.read('xl/cellimages.xml'))
            for ci in ci_root.iter(f'{{{NS_PIC}}}pic'):
                nvPicPr = ci.find(f'{{{NS_PIC}}}nvPicPr')
                if nvPicPr is None:
                    continue
                cNvPr = nvPicPr.find(f'{{{NS_PIC}}}cNvPr')
                if cNvPr is None:
                    continue
                img_name = cNvPr.get('name', '')

                blipFill = ci.find(f'{{{NS_PIC}}}blipFill')
                if blipFill is None:
                    continue
                blip = blipFill.find(f'{{{NS_A}}}blip')
                if blip is None:
                    continue
                embed = blip.get(f'{{{NS_R}}}embed', '')
                if not embed:
                    continue

                img_target = rid_to_target.get(embed)
                if not img_target:
                    continue

                img_path = (f'xl/{img_target}'
                            if not img_target.startswith('xl/')
                            else img_target)
                try:
                    images[img_name] = z.read(img_path)
                except KeyError:
                    continue
    except Exception as e:
        print(f'  ⚠️ 提取图片失败: {e}')

    return images


def _map_images_to_rows(data_rows, data_styles, wps_images):
    """
    将 DISPIMG ID 映射到数据行号。

    返回: {src_row: img_bytes}
    """
    img_map = {}
    for r in data_rows:
        cell_info = data_styles.get((r, IMG_COL), {})
        val = cell_info.get('value', '') or ''
        m = re.search(r'DISPIMG\("([^"]+)"', str(val))
        if m:
            dispimg_id = m.group(1)
            if dispimg_id in wps_images:
                img_map[r] = wps_images[dispimg_id]
    return img_map


# ═══════════════════════════════════════════════════════════════
#  源文件解析
# ═══════════════════════════════════════════════════════════════

def parse_source(src_path):
    """
    解析源文件，提取所有需要的数据。

    返回 dict，包含：
      - currency, rate, service_name
      - header_styles: {(row, col): style_dict}
      - col_header_styles: {col: style_dict}
      - data_styles: {(row, col): style_dict}
      - box_groups: [{box_no, rows:[], total_price, rmb}]
      - img_map: {src_row: img_bytes}
      - merged_cells: [str]
      - col_widths: {col: width}
      - row_heights: {row: height}
      - aux_sheets: {name: [[cells]]}
    """
    wb = load_workbook(src_path, data_only=True)

    if '发票' not in wb.sheetnames:
        print(f'❌ 源文件缺少「发票」工作表，找到: {wb.sheetnames}')
        sys.exit(1)

    ws = wb['发票']
    max_row = ws.max_row

    # ── 1. 提取元信息 ──
    currency = str(ws.cell(24, 2).value or '').strip()
    rate = CURRENCY_RATES.get(currency, 9)
    service_name = str(ws.cell(1, 2).value or '').strip()
    print(f'  币种: {currency} → 汇率: {rate}')
    print(f'  投保拆分公式: 每箱RMB = 单箱子货值 × 汇率({rate}) × 1.1')
    print(f'  服务: {service_name}')

    # ── 2. 头部 Row 1-26 样式 ──
    header_styles = {}
    for r in range(1, HEADER_END + 1):
        for c in range(1, TOTAL_COLS + 1):
            cell = ws.cell(r, c)
            header_styles[(r, c)] = {'value': cell.value, **_copy_style(cell)}

    # ── 3. 列头 Row 27 样式 ──
    col_header_styles = {}
    for c in range(1, TOTAL_COLS + 1):
        cell = ws.cell(COL_HEADER_ROW, c)
        col_header_styles[c] = {'value': cell.value, **_copy_style(cell)}

    # ── 4. 数据行读取 ──
    data_styles = {}
    data_rows = []

    for r in range(DATA_START_ROW, max_row + 1):
        box_no = ws.cell(r, 1).value
        if box_no is None or str(box_no).strip() == '':
            break
        data_rows.append(r)
        for c in range(1, NUM_DATA_COLS + 1):
            cell = ws.cell(r, c)
            data_styles[(r, c)] = {'value': cell.value, **_copy_style(cell)}

    # ── 5. 箱组分组 ──
    box_groups = []
    cur_no = None
    cur_group = None

    for r in data_rows:
        bno = str(ws.cell(r, 1).value).strip()
        qty = ws.cell(r, 4).value or 0
        uprice = ws.cell(r, 5).value or 0

        if bno != cur_no:
            cur_no = bno
            cur_group = {
                'box_no': bno,
                'rows': [r],
                'total_price': qty * uprice,
            }
            box_groups.append(cur_group)
        else:
            cur_group['rows'].append(r)
            cur_group['total_price'] += qty * uprice

    # 计算每箱 RMB（每箱RMB = 单箱子货值 × 汇率 × 1.1）
    for bg in box_groups:
        bg['rmb'] = round(bg['total_price'] * rate * 1.1, 2)

    # ── 6. 提取图片 ──
    wps_images_all = _extract_images(src_path)
    img_map = _map_images_to_rows(data_rows, data_styles, wps_images_all)
    if img_map:
        print(f'  提取图片: {len(img_map)} 张')

    # ── 7. 合并单元格（仅头部区域） ──
    merged_cells = []
    for mc in ws.merged_cells.ranges:
        if mc.min_row < DATA_START_ROW:
            merged_cells.append(str(mc))

    # ── 8. 列宽 ──
    col_widths = {}
    for c in range(1, TOTAL_COLS + 1):
        letter = get_column_letter(c)
        if letter in ws.column_dimensions:
            w = ws.column_dimensions[letter].width
            if w:
                col_widths[c] = w

    # ── 9. 行高（头部） ──
    row_heights = {}
    for r in range(1, DATA_START_ROW):
        if r in ws.row_dimensions:
            h = ws.row_dimensions[r].height
            if h:
                row_heights[r] = h

    # ── 10. 辅助 Sheet ──
    aux_sheets = {}
    for sn in AUX_SHEETS:
        if sn in wb.sheetnames:
            aws = wb[sn]
            rows = []
            for row in aws.iter_rows(min_row=1, max_row=aws.max_row,
                                     max_col=aws.max_column):
                rows.append([cell.value for cell in row])
            aux_sheets[sn] = rows

    wb.close()

    return {
        'currency': currency,
        'rate': rate,
        'service_name': service_name,
        'header_styles': header_styles,
        'col_header_styles': col_header_styles,
        'data_styles': data_styles,
        'data_rows': data_rows,
        'box_groups': box_groups,
        'img_map': img_map,
        'merged_cells': merged_cells,
        'col_widths': col_widths,
        'row_heights': row_heights,
        'aux_sheets': aux_sheets,
    }


# ═══════════════════════════════════════════════════════════════
#  区间分配
# ═══════════════════════════════════════════════════════════════

def assign_ranges(box_groups):
    """按每箱 RMB 将箱组分配到各投保区间"""
    assigned = {r[0]: [] for r in RANGES}
    for bg in box_groups:
        placed = False
        for r_name, r_low, r_high in RANGES:
            if r_low <= bg['rmb'] < r_high:
                assigned[r_name].append(bg)
                placed = True
                break
        if not placed:
            assigned[RANGES[-1][0]].append(bg)
    return assigned


# ═══════════════════════════════════════════════════════════════
#  输出文件生成
# ═══════════════════════════════════════════════════════════════

def create_range_output(src_data, range_name, box_groups, output_path):
    """
    为一个投保区间生成独立 .xlsx 文件。
    """
    wb = Workbook()
    ws = wb.active
    ws.title = '发票'

    total_boxes = len(box_groups)

    # ══════════════════════════════════════════════════════
    #  1. 头部 Row 1-26
    # ══════════════════════════════════════════════════════
    for r in range(1, HEADER_END + 1):
        for c in range(1, TOTAL_COLS + 1):
            cell = ws.cell(r, c)
            info = src_data['header_styles'].get((r, c), {})
            cell.value = info.get('value')
            if 'font' in info:
                _apply_style(cell, info)

    for mc_str in src_data['merged_cells']:
        ws.merge_cells(mc_str)

    ws.cell(1, 1).value = (
        f'{src_data["service_name"]} - {range_name} ({total_boxes}箱)'
    )
    ws.cell(26, 2).value = total_boxes

    # ══════════════════════════════════════════════════════
    #  2. 列头 Row 27
    # ══════════════════════════════════════════════════════
    for c in range(1, TOTAL_COLS + 1):
        cell = ws.cell(COL_HEADER_ROW, c)
        info = src_data['col_header_styles'].get(c, {})
        cell.value = info.get('value')
        if 'font' in info:
            _apply_style(cell, info)

    # T 列列头
    t_hdr = ws.cell(COL_HEADER_ROW, 20)
    t_hdr.value = '每箱RMB'
    t_hdr.font = Font(name='微软雅黑', size=10, bold=True)
    t_hdr.alignment = Alignment(horizontal='center', vertical='center')
    t_hdr.border = THIN_BORDER
    t_hdr.fill = PatternFill(start_color='DAEEF3', end_color='DAEEF3',
                             fill_type='solid')

    # ══════════════════════════════════════════════════════
    #  3. 数据行 Row 28+
    # ══════════════════════════════════════════════════════
    current_row = DATA_START_ROW
    img_map = src_data.get('img_map', {})

    for bg in box_groups:
        for src_row in bg['rows']:
            # 统一行高 = 16
            ws.row_dimensions[current_row].height = 16

            # A-S 列
            for c in range(1, NUM_DATA_COLS + 1):
                cell = ws.cell(current_row, c)
                info = src_data['data_styles'].get((src_row, c), {})
                # P列(IMG_COL=16)：去掉 DISPIMG 公式，改为纯文本标注（图片通过 openpyxl Image 嵌入显示）
                if c == IMG_COL:
                    cell.value = '[产品图片]'
                else:
                    cell.value = info.get('value')
                if 'font' in info:
                    _apply_style(cell, info)

            # T 列：每箱 RMB
            t_cell = ws.cell(current_row, 20)
            t_cell.value = bg['rmb']
            t_cell.font = Font(name='微软雅黑', size=10)
            t_cell.alignment = Alignment(horizontal='center', vertical='center')
            t_cell.number_format = '#,##0.00'
            t_cell.border = THIN_BORDER

            # 嵌入产品图片（P 列）
            if src_row in img_map:
                try:
                    img_bytes = img_map[src_row]
                    img = XlImage(io.BytesIO(img_bytes))
                    # 缩放到合适尺寸（宽80像素，高80像素）
                    img.width = 80
                    img.height = 80
                    img.anchor = f'P{current_row}'
                    ws.add_image(img)
                except Exception as e:
                    print(f'  ⚠️ 嵌入图片 Row {src_row} 失败: {e}')

            current_row += 1

    # ══════════════════════════════════════════════════════
    #  4. 列宽 & 行高（头部）
    # ══════════════════════════════════════════════════════
    for c, w in src_data['col_widths'].items():
        letter = get_column_letter(c)
        ws.column_dimensions[letter].width = w
    ws.column_dimensions['T'].width = 15

    for r, h in src_data['row_heights'].items():
        if r <= COL_HEADER_ROW:
            ws.row_dimensions[r].height = h

    # ══════════════════════════════════════════════════════
    #  5. 辅助 Sheet
    # ══════════════════════════════════════════════════════
    for sn, rows in src_data['aux_sheets'].items():
        ws_aux = wb.create_sheet(title=sn)
        for r_idx, row_vals in enumerate(rows, 1):
            for c_idx, val in enumerate(row_vals, 1):
                ws_aux.cell(r_idx, c_idx).value = val

    # ══════════════════════════════════════════════════════
    #  6. 保存
    # ══════════════════════════════════════════════════════
    wb.save(output_path)
    wb.close()


# ═══════════════════════════════════════════════════════════════
#  Web 集成接口（供 Flask 应用调用）
# ═══════════════════════════════════════════════════════════════

def split_invoice_to_ranges(src_path, output_dir=None):
    """
    供 JTT电商AI助手 Web 应用调用的高层接口。

    返回: dict {range_name: output_file_path, ...}
    """
    if output_dir is None:
        output_dir = os.path.join(os.path.dirname(src_path), 'output')
    os.makedirs(output_dir, exist_ok=True)

    src_data = parse_source(src_path)
    range_boxes = assign_ranges(src_data['box_groups'])

    out_files = {}
    for r_name, _, _ in RANGES:
        boxes = range_boxes[r_name]
        total = len(boxes)
        out_name = f'{r_name} ({total}箱).xlsx'
        out_path = os.path.join(output_dir, out_name)
        create_range_output(src_data, r_name, boxes, out_path)
        out_files[r_name] = out_path

    return out_files


# ═══════════════════════════════════════════════════════════════
#  主入口
# ═══════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description='下单发票 → 按投保区间拆分 v2',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python3 split_insurance_v2.py 下单发票.xlsx
  python3 split_insurance_v2.py 下单发票.xlsx -o ./输出目录
        """,
    )
    parser.add_argument('input', help='下单发票 .xlsx 文件路径')
    parser.add_argument('--out-dir', '-o',
                        help='输出目录（默认: 源文件所在目录下的 output 目录）')

    args = parser.parse_args()

    src_path = os.path.abspath(args.input)
    if not os.path.exists(src_path):
        print(f'❌ 文件不存在: {src_path}')
        sys.exit(1)

    if args.out_dir:
        output_dir = os.path.abspath(args.out_dir)
    else:
        output_dir = os.path.join(os.path.dirname(src_path), 'output')
    os.makedirs(output_dir, exist_ok=True)

    print(f'📂 读取源文件: {os.path.basename(src_path)}')
    src_data = parse_source(src_path)

    print(f'  总箱数: {len(src_data["box_groups"])}')
    for bg in src_data['box_groups']:
        print(f'  📦 {bg["box_no"]}: {len(bg["rows"])} 品名行, '
              f'原币 {bg["total_price"]:.2f} → ¥{bg["rmb"]:.2f}')

    # 分配区间
    range_boxes = assign_ranges(src_data['box_groups'])

    print(f'\n📊 按区间拆分:')
    out_files = []
    for r_name, _, _ in RANGES:
        boxes = range_boxes[r_name]
        total = len(boxes)
        out_name = f'{r_name} ({total}箱).xlsx'
        out_path = os.path.join(output_dir, out_name)
        create_range_output(src_data, r_name, boxes, out_path)
        print(f'  ✅ {out_name}')
        out_files.append(out_path)

    print(f'\n✅ 拆分完成! 共 {len(out_files)} 个文件')
    print(f'   保存目录: {output_dir}')


if __name__ == '__main__':
    main()
