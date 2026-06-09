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
import shutil
import tempfile
from copy import copy
from xml.etree import ElementTree as ET

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter


# ═══════════════════════════════════════════════════════════════
#  常量
# ═══════════════════════════════════════════════════════════════

CURRENCY_RATES = {
    '美金': 7, '美元': 7, 'USD': 7,
    '欧元': 8, 'EUR': 8,
    '英镑': 9, 'GBP': 9,
}

RANGES = [
    ('不足5000RMB',    0,  5000),
    ('5000-10000RMB',  5000, 10000),
    ('10000-20000RMB', 10000, 20000),
    ('20000-30000RMB', 20000, 30000),
    ('30000-40000RMB', 30000, 40000),
]

HEADER_END = 26
COL_HEADER_ROW = 27
DATA_START_ROW = 28
NUM_DATA_COLS = 19    # A-S
TOTAL_COLS = 20       # +T辅助列
IMG_COL = 16          # P列 = 产品图片

AUX_SHEETS = ['FBA地址库编码表', '服务名称', '换算']

THIN_SIDE = Side(style='thin')
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE,
                     top=THIN_SIDE, bottom=THIN_SIDE)


# ═══════════════════════════════════════════════════════════════
#  辅助函数
# ═══════════════════════════════════════════════════════════════

def _copy_style(src_cell):
    return {
        'font': copy(src_cell.font) if src_cell.font is not None else Font(),
        'alignment': (copy(src_cell.alignment)
                      if src_cell.alignment is not None else Alignment()),
        'border': copy(src_cell.border) if src_cell.border is not None else Border(),
        'fill': copy(src_cell.fill) if src_cell.fill is not None else PatternFill(),
        'number_format': src_cell.number_format,
    }


def _apply_style(dst_cell, style):
    dst_cell.font = copy(style['font'])
    dst_cell.alignment = copy(style['alignment'])
    dst_cell.border = copy(style['border'])
    dst_cell.fill = copy(style['fill'])
    dst_cell.number_format = style['number_format']


def _normalize_box_no(box_no):
    """
    货箱编号统一格式：确保末尾数字前有 U0000 前缀。

    如果箱号中已包含 "U0000"（如 FBA15...U000001），则原样保留。
    如果箱号是纯数字（如 "1"），则转为 "U00001"。
    """
    s = str(box_no).strip()
    # 已包含 U0000 → 无需处理
    if 'U0000' in s:
        return s
    m = re.search(r'(\d+)$', s)
    if not m:
        return f'U0000{s}'
    prefix = s[:m.start()]
    digits = m.group(1)
    return f'{prefix}U0000{digits}'


def _calc_total_boxes(box_groups):
    """
    箱数 = 末尾数字最大值 - 末尾数字最小值 + 1

    从所有箱号中提取尾部数字序号，计算涵盖的箱数范围。
    """
    nums = []
    for bg in box_groups:
        m = re.search(r'(\d+)$', str(bg['box_no']).strip())
        if m:
            nums.append(int(m.group(1)))
    if not nums:
        return len(box_groups)
    return max(nums) - min(nums) + 1


# ═══════════════════════════════════════════════════════════════
#  图片提取 + 保留（WPS cellimages.xml 格式）
# ═══════════════════════════════════════════════════════════════

def _get_cellimages_zipdata(src_path):
    """
    从源文件提取 WPS cellimages 相关数据，供输出文件后处理使用。

    返回 dict: {'cellimages.xml': bytes, 'rels.xml': bytes, 'media': [(name, bytes), ...]}
    """
    data = {}
    try:
        with zipfile.ZipFile(src_path, 'r') as z:
            if 'xl/cellimages.xml' in z.namelist():
                data['cellimages.xml'] = z.read('xl/cellimages.xml')
            if 'xl/_rels/cellimages.xml.rels' in z.namelist():
                data['cellimages.xml.rels'] = z.read('xl/_rels/cellimages.xml.rels')
            # 收集 xl/media/ 下的所有图片文件
            media_files = []
            for name in z.namelist():
                if name.startswith('xl/media/') and not name.endswith('/'):
                    media_files.append((name, z.read(name)))
            data['media'] = media_files
            # 记录 Content_Types 中已有的 png/jpeg 扩展名
            ct = z.read('[Content_Types].xml')
            data['content_types'] = ct
    except Exception:
        pass
    return data


def _embed_cellimages_postprocess(output_path, cellimages_data):
    """
    后处理输出 xlsx：将 WPS cellimages.xml 及其图片文件复制到输出文件中，
    使 DISPIMG 公式能在 WPS 中正常渲染图片。
    """
    if not cellimages_data or 'cellimages.xml' not in cellimages_data:
        return

    tmp_dir = tempfile.mkdtemp()
    try:
        # 解压输出文件
        with zipfile.ZipFile(output_path, 'r') as z:
            z.extractall(tmp_dir)

        # 复制 cellimages.xml
        ci_path = os.path.join(tmp_dir, 'xl', 'cellimages.xml')
        with open(ci_path, 'wb') as f:
            f.write(cellimages_data['cellimages.xml'])

        # 复制 cellimages.xml.rels
        if 'cellimages.xml.rels' in cellimages_data:
            rels_dir = os.path.join(tmp_dir, 'xl', '_rels')
            os.makedirs(rels_dir, exist_ok=True)
            rels_path = os.path.join(rels_dir, 'cellimages.xml.rels')
            with open(rels_path, 'wb') as f:
                f.write(cellimages_data['cellimages.xml.rels'])

        # 复制 xl/media/ 图片文件
        if 'media' in cellimages_data:
            media_dir = os.path.join(tmp_dir, 'xl', 'media')
            os.makedirs(media_dir, exist_ok=True)
            for name, data in cellimages_data['media']:
                dst = os.path.join(tmp_dir, name)
                os.makedirs(os.path.dirname(dst), exist_ok=True)
                with open(dst, 'wb') as f:
                    f.write(data)

        # 补充 Content_Types 中缺少的图片类型
        if 'content_types' in cellimages_data:
            NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
            ct_path = os.path.join(tmp_dir, '[Content_Types].xml')
            ct_tree = ET.parse(ct_path)
            ct_root = ct_tree.getroot()

            # 检查已注册的扩展名
            existing_exts = set()
            for child in ct_root.findall(f'{{{NS_CT}}}Default'):
                ext = child.get('Extension', '')
                if ext:
                    existing_exts.add(ext.lower())

            needed = {'png': 'image/png', 'jpeg': 'image/jpeg', 'jpg': 'image/jpeg'}
            for ext, ctype in needed.items():
                if ext not in existing_exts:
                    el = ET.SubElement(ct_root, f'{{{NS_CT}}}Default')
                    el.set('Extension', ext)
                    el.set('ContentType', ctype)

            ct_tree.write(ct_path, xml_declaration=True, encoding='UTF-8')

        # 重新打包
        tmp_out = output_path + '.tmp'
        with zipfile.ZipFile(tmp_out, 'w', zipfile.ZIP_DEFLATED) as zout:
            for dirpath, _, filenames in os.walk(tmp_dir):
                for fn in filenames:
                    full = os.path.join(dirpath, fn)
                    arcname = os.path.relpath(full, tmp_dir)
                    zout.write(full, arcname)

        shutil.move(tmp_out, output_path)

    except Exception as e:
        print(f'  ⚠️ 图片后处理失败: {e}')
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ═══════════════════════════════════════════════════════════════
#  源文件解析
# ═══════════════════════════════════════════════════════════════

def parse_source(src_path):
    """
    解析源文件，提取所有需要的数据。

    返回 dict:
      - currency, rate, service_name
      - header_styles, col_header_styles, data_styles
      - box_groups: [{box_no, rows[], total_price, rmb}]
      - cellimages_data: 供后处理使用
      - merged_cells, col_widths, row_heights, aux_sheets
    """
    wb = load_workbook(src_path, data_only=True)

    if '发票' not in wb.sheetnames:
        print(f'❌ 源文件缺少「发票」工作表，找到: {wb.sheetnames}')
        sys.exit(1)

    ws = wb['发票']
    max_row = ws.max_row

    # ── 元信息 ──
    currency = str(ws.cell(24, 2).value or '').strip()
    rate = CURRENCY_RATES.get(currency, 9)
    service_name = str(ws.cell(1, 2).value or '').strip()
    print(f'  币种: {currency} → 汇率: {rate}')
    print(f'  公式: 每箱RMB = 单箱子货值 × 汇率({rate}) × 1.1')
    print(f'  服务: {service_name}')

    # ── 头部 Row 1-26 ──
    header_styles = {}
    for r in range(1, HEADER_END + 1):
        for c in range(1, TOTAL_COLS + 1):
            cell = ws.cell(r, c)
            header_styles[(r, c)] = {'value': cell.value, **_copy_style(cell)}

    # ── 列头 Row 27 ──
    col_header_styles = {}
    for c in range(1, TOTAL_COLS + 1):
        cell = ws.cell(COL_HEADER_ROW, c)
        col_header_styles[c] = {'value': cell.value, **_copy_style(cell)}

    # ── 数据行 ──
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

    # ── 箱组分组（按货箱编号） ──
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

    # 每箱 RMB = 货值 × 汇率 × 1.1
    for bg in box_groups:
        bg['rmb'] = round(bg['total_price'] * rate * 1.1, 2)

    # ── 提取 cellimages 数据（供后处理） ──
    cellimages_data = _get_cellimages_zipdata(src_path)
    if cellimages_data:
        print(f'  提取图片: {len(cellimages_data.get("media", []))} 张')

    # ── 合并单元格（头部） ──
    merged_cells = []
    for mc in ws.merged_cells.ranges:
        if mc.min_row < DATA_START_ROW:
            merged_cells.append(str(mc))

    # ── 列宽 ──
    col_widths = {}
    for c in range(1, TOTAL_COLS + 1):
        letter = get_column_letter(c)
        if letter in ws.column_dimensions:
            w = ws.column_dimensions[letter].width
            if w:
                col_widths[c] = w

    # ── 行高（头部） ──
    row_heights = {}
    for r in range(1, DATA_START_ROW):
        if r in ws.row_dimensions:
            h = ws.row_dimensions[r].height
            if h:
                row_heights[r] = h

    # ── 辅助 Sheet ──
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
        'cellimages_data': cellimages_data,
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
    保存后会自动后处理嵌入 WPS cellimages 图片。
    """
    wb = Workbook()
    ws = wb.active
    ws.title = '发票'

    # 箱数 = 末尾数字 max-min+1
    total_boxes = _calc_total_boxes(box_groups)

    # ═══════════════════════════════════════════════
    #  1. 头部 Row 1-26
    # ═══════════════════════════════════════════════
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

    # ═══════════════════════════════════════════════
    #  2. 列头 Row 27
    # ═══════════════════════════════════════════════
    for c in range(1, TOTAL_COLS + 1):
        cell = ws.cell(COL_HEADER_ROW, c)
        info = src_data['col_header_styles'].get(c, {})
        cell.value = info.get('value')
        if 'font' in info:
            _apply_style(cell, info)

    t_hdr = ws.cell(COL_HEADER_ROW, 20)
    t_hdr.value = '每箱RMB'
    t_hdr.font = Font(name='微软雅黑', size=10, bold=True)
    t_hdr.alignment = Alignment(horizontal='center', vertical='center')
    t_hdr.border = THIN_BORDER
    t_hdr.fill = PatternFill(start_color='DAEEF3', end_color='DAEEF3',
                             fill_type='solid')

    # ═══════════════════════════════════════════════
    #  3. 数据行 Row 28+
    # ═══════════════════════════════════════════════
    current_row = DATA_START_ROW

    for bg in box_groups:
        for src_row in bg['rows']:
            # 统一行高 16
            ws.row_dimensions[current_row].height = 16

            for c in range(1, NUM_DATA_COLS + 1):
                cell = ws.cell(current_row, c)
                info = src_data['data_styles'].get((src_row, c), {})

                if c == 1:
                    # A列：货箱编号 → 统一加 U0000 前缀
                    raw = info.get('value', '')
                    cell.value = _normalize_box_no(raw)
                elif c == IMG_COL:
                    # P列：保留原始 DISPIMG 公式（WPS 通过 cellimages.xml 渲染）
                    cell.value = info.get('value')
                else:
                    cell.value = info.get('value')

                if 'font' in info:
                    _apply_style(cell, info)

            # T列：每箱 RMB
            t_cell = ws.cell(current_row, 20)
            t_cell.value = bg['rmb']
            t_cell.font = Font(name='微软雅黑', size=10)
            t_cell.alignment = Alignment(horizontal='center', vertical='center')
            t_cell.number_format = '#,##0.00'
            t_cell.border = THIN_BORDER

            current_row += 1

    # ═══════════════════════════════════════════════
    #  4. 列宽 & 头部行高
    # ═══════════════════════════════════════════════
    for c, w in src_data['col_widths'].items():
        letter = get_column_letter(c)
        ws.column_dimensions[letter].width = w
    ws.column_dimensions['T'].width = 15

    for r, h in src_data['row_heights'].items():
        if r <= COL_HEADER_ROW:
            ws.row_dimensions[r].height = h

    # ═══════════════════════════════════════════════
    #  5. 辅助 Sheet
    # ═══════════════════════════════════════════════
    for sn, rows in src_data['aux_sheets'].items():
        ws_aux = wb.create_sheet(title=sn)
        for r_idx, row_vals in enumerate(rows, 1):
            for c_idx, val in enumerate(row_vals, 1):
                ws_aux.cell(r_idx, c_idx).value = val

    # ═══════════════════════════════════════════════
    #  6. 保存
    # ═══════════════════════════════════════════════
    wb.save(output_path)
    wb.close()

    # ═══════════════════════════════════════════════
    #  7. 后处理：嵌入 WPS cellimages 图片
    # ═══════════════════════════════════════════════
    cellimages_data = src_data.get('cellimages_data', {})
    if cellimages_data:
        _embed_cellimages_postprocess(output_path, cellimages_data)


# ═══════════════════════════════════════════════════════════════
#  Web 集成接口
# ═══════════════════════════════════════════════════════════════

def split_invoice_to_ranges(src_path, output_dir=None):
    """
    供 JTT电商AI助手 Web 应用调用的高层接口。
    返回: dict {range_name: output_file_path}
    """
    if output_dir is None:
        output_dir = os.path.join(os.path.dirname(src_path), 'output')
    os.makedirs(output_dir, exist_ok=True)

    src_data = parse_source(src_path)
    range_boxes = assign_ranges(src_data['box_groups'])

    out_files = {}
    for r_name, _, _ in RANGES:
        boxes = range_boxes[r_name]
        total = _calc_total_boxes(boxes)
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

    print(f'  总箱数: {_calc_total_boxes(src_data["box_groups"])}')
    for bg in src_data['box_groups']:
        print(f'  📦 {_normalize_box_no(bg["box_no"])}: {len(bg["rows"])} 品名行, '
              f'原币 {bg["total_price"]:.2f} → ¥{bg["rmb"]:.2f}')

    range_boxes = assign_ranges(src_data['box_groups'])

    print(f'\n📊 按区间拆分:')
    out_files = []
    for r_name, _, _ in RANGES:
        boxes = range_boxes[r_name]
        total = _calc_total_boxes(boxes)
        out_name = f'{r_name} ({total}箱).xlsx'
        out_path = os.path.join(output_dir, out_name)
        create_range_output(src_data, r_name, boxes, out_path)
        print(f'  ✅ {out_name}')
        out_files.append(out_path)

    print(f'\n✅ 拆分完成! 共 {len(out_files)} 个文件')
    print(f'   保存目录: {output_dir}')


if __name__ == '__main__':
    main()
