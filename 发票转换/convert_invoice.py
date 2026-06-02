#!/usr/bin/env python3
"""
TR发票 → 供应商发票 转换工具

功能：将TR/思锐/赛诺吉下单发票(.xlsx)转换为以下供应商格式：
  1. 天图 (--to 天图)
  2. 航乐英国 (--to 航乐-uk)
  3. 航乐欧洲 (--to 航乐-eu)

用法：
  python3 convert_invoice.py <TR发票.xlsx> --to 天图 [输出路径]
  python3 convert_invoice.py <TR发票.xlsx> --to 航乐-uk [输出路径]
  python3 convert_invoice.py <TR发票.xlsx> --to 航乐-eu [输出路径]

源文件兼容：
  - TR系统下单发票  ✅
  - 思锐客户下单发票 ✅
  - 赛诺吉发票      ✅

TR源文件结构：Sheet=Page1, 头部Row 1-16, 数据Row 18+
"""

import sys
import os
import re
import shutil
import argparse
from copy import copy
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import (
    Font, Alignment, Border, Side, PatternFill
)
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as OpenpyxlImage
from io import BytesIO


# ═══════════════════════════════════════════════════════════════
#  常量
# ═══════════════════════════════════════════════════════════════

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

TIANTU_TEMPLATE = os.path.join(SCRIPT_DIR, '天图单票专用模板20260601.xlsx')
HANGLE_UK_TEMPLATE = os.path.join(SCRIPT_DIR, '航乐-客户名称 客户单号 英国发票模板9.9更新.xls')
HANGLE_EU_TEMPLATE = os.path.join(SCRIPT_DIR, '航乐-客户单号- 欧州发票模板2.26更新.xls')

# TR发票列映射 (Row 17+ header)
TR_COL = {
    'A': '箱号', 'B': '重量_KG', 'C': '长度_CM', 'D': '宽度_CM', 'E': '高度_CM',
    'F': '英文品名', 'G': '中文品名', 'H': '单价_USD', 'I': '数量',
    'J': '材质', 'K': '海关编码', 'L': '用途', 'M': '品牌', 'N': '型号',
    'O': '链接', 'P': '销售价格', 'Q': '产品图片', 'R': '产品重量_KG',
    'S': 'ASIN', 'T': 'FNSKU', 'U': 'SKU', 'V': 'PO_Number',
}

# ═══════════════════════════════════════════════════════════════
#  1. TR发票读取
# ═══════════════════════════════════════════════════════════════

class TRInvoice:
    """解析TR/思锐/赛诺吉下单发票"""

    def __init__(self, path):
        self.path = path
        self.wb = load_workbook(path, data_only=True)
        self.ws = self.wb['Page1']
        self.header = {}       # 头部字段 dict
        self.data_rows = []    # 数据行列表
        self.images = {}       # {行号: image_bytes} 提取的嵌入图片
        self._parse()

    def _cell_str(self, row, col):
        v = self.ws.cell(row=row, column=col).value
        if v is None:
            return ''
        return str(v).strip()

    def _cell_val(self, row, col):
        return self.ws.cell(row=row, column=col).value

    def _parse_header_pair(self, row, label_col, val_col, label_key, val_key):
        """解析一对 标签列:值列，如 A:B, E:F"""
        label = self._cell_str(row, label_col)
        val = self._cell_val(row, val_col)
        if label:
            self.header[label.rstrip('*:')] = val
        return val

    def _parse(self):
        ws = self.ws

        # ── 头部 Row 1-16 ──
        # 左侧 A/B 列
        for r in range(1, 17):
            a = self._cell_str(r, 1)
            b = self._cell_val(r, 2)
            if a:
                self.header[a.rstrip('*:')] = b

        # 右侧 E/F 列 (覆盖写入，优先级高于左侧)
        for r in range(1, 17):
            e = self._cell_str(r, 5)
            f = self._cell_val(r, 6)
            if e:
                self.header[e.rstrip('*:')] = f

        # 右侧 I/J 列
        for r in range(1, 17):
            i = self._cell_str(r, 9)
            j = self._cell_val(r, 10)
            if i:
                self.header[i.rstrip('*:')] = j

        # ── 数据行 Row 18+ ──
        max_row = ws.max_row
        for r in range(18, max_row + 1):
            box_no = self._cell_val(r, 1)
            if box_no is None or str(box_no).strip() == '':
                continue
            row_data = {
                'A': box_no,                     # 箱号
                'B': self._cell_val(r, 2),       # 重量
                'C': self._cell_val(r, 3),       # 长度
                'D': self._cell_val(r, 4),       # 宽度
                'E': self._cell_val(r, 5),       # 高度
                'F': self._cell_val(r, 6),       # 英文品名
                'G': self._cell_val(r, 7),       # 中文品名
                'H': self._cell_val(r, 8),       # 单价
                'I': self._cell_val(r, 9),       # 数量
                'J': self._cell_val(r, 10),      # 材质
                'K': self._cell_val(r, 11),      # 海关编码
                'L': self._cell_val(r, 12),      # 用途
                'M': self._cell_val(r, 13),      # 品牌
                'N': self._cell_val(r, 14),      # 型号
                'O': self._cell_val(r, 15),      # 链接
                'P': self._cell_val(r, 16),      # 销售价格
                'Q': self._cell_val(r, 17),      # 产品图片
                'R': self._cell_val(r, 18),      # 产品重量
                'S': self._cell_val(r, 19),      # ASIN
                'T': self._cell_val(r, 20),      # FNSKU
                'U': self._cell_val(r, 21),      # SKU
                'V': self._cell_val(r, 22),      # PO Number
                '_row': r,                       # 源文件行号（图片映射用）
            }
            self.data_rows.append(row_data)

        # ── 提取嵌入图片 (来自Q列/col 17) ──
        if hasattr(ws, '_images') and ws._images:
            for img in ws._images:
                try:
                    anchor = img.anchor
                    if hasattr(anchor, '_from'):
                        img_col = anchor._from.col + 1
                        img_row = anchor._from.row + 1
                        # Q列=17，数据行范围18+
                        if img_col == 17 and img_row >= 18:
                            # 获取图片字节数据
                            img_data = None
                            ref = getattr(img, 'ref', None)
                            if isinstance(ref, BytesIO):
                                img_data = ref.getvalue()
                            elif isinstance(ref, str):
                                # 通过关系ID从workbook获取
                                if hasattr(self.wb, '_images') and ref in self.wb._images:
                                    part = self.wb._images[ref]
                                    if hasattr(part, '_blob'):
                                        img_data = part._blob
                            if img_data is None:
                                data = img._data()
                                if isinstance(data, bytes):
                                    img_data = data
                                elif isinstance(data, BytesIO):
                                    img_data = data.getvalue()
                            if img_data:
                                self.images[img_row] = img_data
                except Exception:
                    pass

    def get(self, key, default=None):
        """获取头部字段值"""
        return self.header.get(key, default)

    def __repr__(self):
        info = f'TRInvoice({os.path.basename(self.path)})\n'
        info += f'  订单号: {self.get("客户订单号")}\n'
        info += f'  服务: {self.get("服务")}\n'
        info += f'  收件人: {self.get("收件人姓名")}\n'
        info += f'  箱数: {self.get("箱数")}\n'
        info += f'  数据行: {len(self.data_rows)} 行\n'
        return info


# ═══════════════════════════════════════════════════════════════
#  2. 天图 转换
# ═══════════════════════════════════════════════════════════════

def convert_to_tiantu(tr, output_path):
    """
    TR发票 → 天图格式
    规则详见 TR转天图发票_转换规则说明.md
    """
    print(f'📄 天图模板: {TIANTU_TEMPLATE}')
    if not os.path.exists(TIANTU_TEMPLATE):
        print('❌ ERROR: 天图模板文件不存在!')
        return False

    # 输出模板版本日志
    tpl_size = os.path.getsize(TIANTU_TEMPLATE)
    print(f'  模板大小: {tpl_size} bytes')
    # 检查新模板特征：Sheet2行数
    try:
        tmp_wb = load_workbook(TIANTU_TEMPLATE, data_only=True)
        s2_rows = tmp_wb['Sheet2'].max_row
        print(f'  模板Sheet2服务数: {s2_rows} 条')
        tmp_wb.close()
    except:
        pass

    # 复制模板
    shutil.copy(TIANTU_TEMPLATE, output_path)
    wb = load_workbook(output_path)
    ws = wb['Sheet1']

    # ── 样式定义 ──
    thin_side = Side(style='thin')
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    font_label_big = Font(name='微软雅黑', size=11, bold=True)
    font_label_small = Font(name='微软雅黑', size=10, bold=True)
    font_value = Font(name='微软雅黑', size=10)
    font_data = Font(name='微软雅黑', size=10)
    font_header = Font(name='微软雅黑', size=11, bold=True)
    font_total = Font(name='微软雅黑', size=11, bold=True)
    font_red = Font(name='微软雅黑', size=12, bold=True, color='FF0000')

    fill_header = PatternFill(start_color='DAEEF3', end_color='DAEEF3', fill_type='solid')

    center_align = Alignment(horizontal='center', vertical='center')

    # ── 辅助函数 ──
    def set_cell(col, row, value, font=None, align=None, border=None, number_format=None):
        cell = ws[f'{col}{row}']
        cell.value = value
        if font:
            cell.font = font
        if align:
            cell.alignment = align
        if border:
            cell.border = border
        if number_format:
            cell.number_format = number_format
        return cell

    # ── 头部 Row 1-28 ──

    # B1: 服务 = TR 服务字段，同时确保该服务在 Sheet2 下拉列表中
    service_name = tr.get('服务', '')
    if service_name:
        # 检查 Sheet2 中是否已有该服务，没有则追加
        ws2 = wb['Sheet2']
        found = False
        for r in range(1, ws2.max_row + 1):
            if ws2.cell(row=r, column=1).value == service_name:
                found = True
                break
        if not found:
            next_row = ws2.max_row + 1
            ws2.cell(row=next_row, column=1).value = service_name
    set_cell('B', 1, service_name, font_value)

    # B2: 仓库代码 = TR B4 (收件人姓名，如 YVR4/FTW1/POZ1/RFD2)
    wh_code = tr.get('收件人姓名', '')
    set_cell('B', 2, wh_code, font_value)

    # B3: 收件人姓名 = TR B4
    set_cell('B', 3, wh_code, font_value)

    # B4: 收件人公司 = TR B5
    set_cell('B', 4, tr.get('收件人公司', ''), font_value)

    # B5: 收件人地址一 = TR B6
    set_cell('B', 5, tr.get('收件人地址一', ''), font_value)

    # B6: 收件人地址二 → 留空
    set_cell('B', 6, '', font_value)

    # B7: 收件人地址三 → 留空
    set_cell('B', 7, '', font_value)

    # B8: 收件人城市 = TR B9
    set_cell('B', 8, tr.get('收件人城市', ''), font_value)

    # B9: 收件人省份/州 = TR B10
    set_cell('B', 9, tr.get('收件人省份/州', ''), font_value)

    # B10: 收件人邮编 = TR B11
    set_cell('B', 10, tr.get('收件人邮编', ''), font_value)

    # B11: 收件人国家代码 = TR B12, 空则默认US
    country = tr.get('收件人国家代码(二字代码)', '') or tr.get('收件人国家代码', '') or 'US'
    set_cell('B', 11, country, font_value)

    # B12: 收件人电话 = TR B13
    set_cell('B', 12, tr.get('收件人电话', ''), font_value)

    # B13: 收件人邮箱 = TR B14
    set_cell('B', 13, tr.get('收件人邮箱', ''), font_value)

    # B14: 客户订单号 = TR B1
    order_no = tr.get('客户订单号', '')
    set_cell('B', 14, order_no, font_value)

    # B15: Amazon Reference ID → 留空
    set_cell('B', 15, '', font_value)

    # B16: 带电 → "带电"/"不带电"
    has_battery = tr.get('带电', '否')
    set_cell('B', 16, '带电' if has_battery == '是' else '不带电', font_value)

    # B17: 带磁 → "带磁"/"不带磁"
    has_magnet = tr.get('带磁', '否')
    set_cell('B', 17, '带磁' if has_magnet == '是' else '不带磁', font_value)

    # B18: 交税方式 = TR F8
    set_cell('B', 18, tr.get('交税方式', ''), font_value)

    # B19: 报关方式 = TR F6
    set_cell('B', 19, tr.get('报关方式', ''), font_value)

    # B20: 清关方式 = TR F7
    set_cell('B', 20, tr.get('清关方式', ''), font_value)

    # B21: VAT号 = TR E10
    set_cell('B', 21, tr.get('VAT号', ''), font_value)

    # B22: 总箱数 = TR B16
    set_cell('B', 22, tr.get('箱数', ''), font_value)

    # B23: 备注 → 留空
    set_cell('B', 23, '', font_value)

    # B24: 国内报关抬头信息补充 (A24:B24 合并, 保留原标题)
    # keep original

    # B25: 公司名称 (TR中没有，留空)
    set_cell('B', 25, '', font_value)

    # B26: 公司税号 (TR中没有，留空)
    set_cell('B', 26, '', font_value)

    # B27: 真实出口货值 (TR中没有，留空)
    set_cell('B', 27, '', font_value)

    # B28: 币种 = TR J14 (申报币种) 或 TR E14 (币种)
    currency = tr.get('申报币种', '') or tr.get('币种', '')
    if not currency:
        # 根据国家推断
        if country in ('GB', 'UK'):
            currency = '英镑 GBP'
        elif country == 'US':
            currency = '美元 USD'
        elif country in ('DE', 'FR', 'IT', 'ES', 'NL', 'BE', 'PL', 'CZ'):
            currency = '欧元 EUR'
        else:
            currency = '美元 USD'
    set_cell('B', 28, currency, font_value)

    # ── 数据行 Row 30+ ← TR Row 18+ ──
    # 先清除模板中已有的示例数据行
    for r in range(30, 100):
        for c in range(1, 19):  # A-R
            cell = ws.cell(row=r, column=c)
            cell.value = None
            cell.border = Border()

    # 写入数据
    data_start_row = 30
    for i, dr in enumerate(tr.data_rows):
        r = data_start_row + i
        ws.row_dimensions[r].height = 80  # 产品图片需要足够高度

        # A: 货箱编号
        set_cell('A', r, dr['A'], font_data, center_align, thin_border)

        # B: PO Number (优先行级V列，其次头部PO)
        po = dr.get('V') or tr.get('PO Number', '')
        set_cell('B', r, po, font_data, center_align, thin_border)

        # C: 产品英文品名
        set_cell('C', r, dr['F'], font_data, center_align, thin_border)

        # D: 产品中文品名
        set_cell('D', r, dr['G'], font_data, center_align, thin_border)

        # E: 产品申报单价(USD)
        set_cell('E', r, dr['H'] if dr['H'] is not None else '',
                 font_data, center_align, thin_border, '#,##0.00')

        # F: 产品单箱申报数量
        set_cell('F', r, dr['I'] if dr['I'] is not None else '',
                 font_data, center_align, thin_border, '0')

        # G: 产品单箱申报总价(USD) = E × F (计算值)
        e_val = dr['H'] or 0
        f_val = dr['I'] or 0
        g_val = e_val * f_val
        set_cell('G', r, g_val, font_data, center_align, thin_border, '#,##0.00')

        # H: 产品材质
        set_cell('H', r, dr['J'], font_data, center_align, thin_border)

        # I: 产品海关编码
        set_cell('I', r, dr['K'], font_data, center_align, thin_border)

        # J: 产品用途
        set_cell('J', r, dr['L'], font_data, center_align, thin_border)

        # K: 产品品牌
        set_cell('K', r, dr['M'], font_data, center_align, thin_border)

        # L: 产品型号
        set_cell('L', r, dr['N'], font_data, center_align, thin_border)

        # M: 产品图片 (优先嵌入图片，其次放链接文本)
        src_row = dr.get('_row')
        img_bytes = tr.images.get(src_row) if src_row else None
        if img_bytes:
            try:
                img = OpenpyxlImage(BytesIO(img_bytes))
                img.width = 70
                img.height = 70
                ws.add_image(img, f'M{r}')
            except:
                pass
        # 也保留产品图片链接文本（如果有的话）
        q_val = dr.get('Q', '') or ''
        if q_val and not img_bytes:
            set_cell('M', r, q_val, font_data, center_align, thin_border)

        # N: 产品销售链接
        set_cell('N', r, dr['O'] if dr['O'] else '',
                 font_data, center_align, thin_border)

        # O: 货箱重量(KG)
        set_cell('O', r, dr['B'] if dr['B'] is not None else '',
                 font_data, center_align, thin_border, '0.0')

        # P: 货箱长度(CM)
        set_cell('P', r, dr['C'] if dr['C'] is not None else '',
                 font_data, center_align, thin_border, '0.0')

        # Q: 货箱宽度(CM)
        set_cell('Q', r, dr['D'] if dr['D'] is not None else '',
                 font_data, center_align, thin_border, '0.0')

        # R: 货箱高度(CM)
        set_cell('R', r, dr['E'] if dr['E'] is not None else '',
                 font_data, center_align, thin_border, '0.0')

    num_data_rows = len(tr.data_rows)

    # ── 合计行 ──
    total_row = data_start_row + num_data_rows + 1  # 空一行 + 合计行

    # Row 29是表头，保持不动

    set_cell('A', total_row, '合计', font_total, center_align, thin_border)

    # F: 总数量
    total_qty = sum(dr['I'] or 0 for dr in tr.data_rows)
    set_cell('F', total_row, total_qty, font_total, center_align, thin_border, '0')

    # G: 总价
    total_value = sum((dr['H'] or 0) * (dr['I'] or 0) for dr in tr.data_rows)
    set_cell('G', total_row, total_value, font_total, center_align, thin_border, '#,##0.00')

    # 其他列加边框
    for col_l in ['B', 'C', 'D', 'E', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R']:
        cell = ws[f'{col_l}{total_row}']
        cell.border = thin_border
        cell.font = font_total
        cell.alignment = center_align

    # ── 清理模板中的VLOOKUP公式 (B3-B5, B8-B10, B12) ──
    # 已经在前面被覆盖了

    # ── 保存 ──
    wb.save(output_path)
    print(f'✅ 天图发票已生成: {os.path.basename(output_path)}')
    print(f'   数据: {num_data_rows} 行, 合计: {total_qty} 件, ${total_value:.2f}')
    return True


# ═══════════════════════════════════════════════════════════════
#  3. 航乐 转换 (通用)
# ═══════════════════════════════════════════════════════════════

def convert_to_hangle(tr, output_path, region='uk'):
    """
    TR发票 → 航乐格式 (UK 或 EU)

    输出 .xlsx 格式，布局和数据规则参照航乐模板。
    """
    if region == 'uk':
        currency_label = 'GBP'
        currency_name = '英镑'
    else:
        currency_label = 'EUR'
        currency_name = '欧元'

    # ── 获取TR数据 ──
    country = tr.get('收件人国家代码(二字代码)', '') or tr.get('收件人国家代码', '') or 'US'
    country_map = {
        'GB': 'UK', 'UK': 'UK', 'US': 'US', 'DE': 'DE', 'FR': 'FR',
        'IT': 'IT', 'ES': 'ES', 'NL': 'NL', 'BE': 'BE', 'PL': 'PL',
        'CZ': 'CZ', 'CA': 'CA',
    }
    country_code = country_map.get(country.upper(), country.upper())
    has_battery = tr.get('带电', '否')
    has_magnet = tr.get('带磁', '否')

    # ── 创建 Workbook ──
    wb = load_workbook(TIANTU_TEMPLATE)
    for sn in wb.sheetnames:
        del wb[sn]
    ws = wb.create_sheet('Packing list装箱单发票')

    # ── 样式定义 ──
    thin_side = Side(style='thin')
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    normal_font = Font(name='微软雅黑', size=10)
    bold_font = Font(name='微软雅黑', size=10, bold=True)
    title_font = Font(name='微软雅黑', size=11, bold=True, color='FF0000')
    total_font = Font(name='微软雅黑', size=10, bold=True)
    note_font = Font(name='微软雅黑', size=9, color='808080')
    header_fill = PatternFill(start_color='DAEEF3', end_color='DAEEF3', fill_type='solid')
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # ── 列宽 ──
    col_widths = {'A':18,'B':16,'C':22,'D':12,'E':12,'F':16,'G':14,'H':14,'I':16,
                  'J':14,'K':14,'L':14,'M':14,'N':12,'O':12,'P':18,'Q':8,'R':8,
                  'S':10,'T':12,'U':14,'V':12,'W':20,'X':30}
    for cl, w in col_widths.items():
        ws.column_dimensions[cl].width = w

    # ── 辅助：写单元格 ──
    def wcell(r, c, val, font=None, align=None, border=None, nf=None):
        cell = ws.cell(row=r, column=c)
        cell.value = val
        if font: cell.font = font
        if align: cell.alignment = align
        if border: cell.border = border
        if nf: cell.number_format = nf
        return cell

    # ──────────────────────────────────────────────
    #  标题与说明行 (Row 1-2)
    # ──────────────────────────────────────────────
    ws.merge_cells('A1:M1')
    title_text = 'Commerical Invoice（商业发票）&Packing list（装箱单）红色字体及标颜色部分必填'
    wcell(1, 1, title_text, title_font)

    ws.merge_cells('A2:M2')
    wcell(2, 1, '贵司需如实根据实际装箱货物填写清单，如遇海关查验，发现装箱单与实际货物不符，出现瞒报现象，产生责任由贵司承担，并承担罚款费用',
          note_font, Alignment(horizontal='left', vertical='center', wrap_text=True))

    # ──────────────────────────────────────────────
    #  头部信息 (Row 3-9)
    # ──────────────────────────────────────────────
    label_font = bold_font
    val_font = normal_font

    # 行结构: A=标签 B=值 C=标签 D=值 E=标签 F=值 G=标签 H=值 I=标签 J=值 ...
    # Row 3: ADD / COUNTRY / SHIPMENT ID / AMAZON REF
    header_fields_r3 = [
        (1, 'ADD(收件地址)*：', 2, tr.get('收件人地址一', '')),
        (5, 'COUNTRY 国家', 6, country_code),
        (7, 'SHIPMENT ID（FBA货物编码）：', 8, tr.get('客户订单号', '')),
        (10, 'AMAZON REFERENCE ID（亚马逊内部编码）：', 11, ''),
    ]
    for lab_col, lab, val_col, val in header_fields_r3:
        wcell(3, lab_col, lab, label_font)
        wcell(3, val_col, val, val_font)

    # Row 4: ZIP / COMPANY / TEL / 报关类型 / 交货仓库
    header_fields_r4 = [
        (1, 'ZIP CODE（邮编）*：', 2, tr.get('收件人邮编', '')),
        (3, 'COMPANY(收件公司)* ：', 4, tr.get('收件人公司', '')),
        (5, 'TEL(收件人电话) ：', 6, tr.get('收件人电话', '')),
        (8, '报关类型（委托/单证）：必填', 9, tr.get('报关方式', '委托')),
        (10, '交货仓库', 11, tr.get('收件人姓名', '')),
    ]
    for lab_col, lab, val_col, val in header_fields_r4:
        wcell(4, lab_col, lab, label_font)
        wcell(4, val_col, val, val_font)

    # Row 5: CITY / ATTN / EMAIL
    header_fields_r5 = [
        (1, 'CITY*(城市名)：', 2, tr.get('收件人城市', '')),
        (3, 'ATTN(收件人)  ：', 4, tr.get('收件人姓名', '')),
        (5, 'EMAIL*（邮箱）：', 6, tr.get('收件人邮箱', '')),
    ]
    for lab_col, lab, val_col, val in header_fields_r5:
        wcell(5, lab_col, lab, label_font)
        wcell(5, val_col, val, val_font)

    # Row 6-9: VAT & 渠道信息 (use pairs, no full-row merges)
    vat_label = 'VAT公司名称*' if region == 'uk' else '公司名称*'
    wcell(6, 1, vat_label, label_font)
    wcell(6, 2, tr.get('VAT公司英文名', ''), val_font)
    wcell(7, 1, 'VAT号*', label_font)
    wcell(7, 2, tr.get('VAT号', ''), val_font)
    wcell(8, 1, 'EORI号*', label_font)
    wcell(8, 2, tr.get('EORI号', ''), val_font)
    wcell(9, 1, 'VAT公司注册地址*', label_font)
    wcell(9, 2, tr.get('VAT注册地址', ''), val_font)

    # Row 6 额外: 渠道 / 是否包税 / 物品属性
    attrs = []
    if tr.get('带电', '') == '是': attrs.append('带电')
    if tr.get('带磁', '') == '是': attrs.append('带磁')
    if tr.get('液体', '') == '是': attrs.append('液体')
    if tr.get('粉末', '') == '是': attrs.append('粉末')
    if tr.get('危险品', '') == '是': attrs.append('危险品')
    attr_str = ','.join(attrs) if attrs else '普货'

    wcell(6, 8, '渠道：', label_font)
    wcell(6, 9, tr.get('服务', ''), val_font)
    wcell(6, 11, '是否包税：', label_font)
    wcell(6, 12, tr.get('交税方式', ''), val_font)
    wcell(6, 13, '物品属性', label_font)
    wcell(6, 14, attr_str, val_font)

    # ──────────────────────────────────────────────
    #  列标题 Row 11
    # ──────────────────────────────────────────────
    col_headers = [
        (1, 'Box No.\nFBA箱号'), (2, '品名（中文）'), (3, '品名（英文)'),
        (4, '材质\n（中文）'), (5, '材质\n（英文）'),
        (6, '海关编码\n（十位数）'), (7, '型号'), (8, '品牌'),
        (9, '产品用途\n（英文）'), (10, '*Quantity (pcs)\n单箱数量'),
        (11, '*Quantity (pcs)\n总数量'), (12, 'Net Weight\n单个产品净重'),
        (13, f'Gross weight(kg)\n实重'),
        (14, f'单价\n({currency_label})'),
        (15, f'总价\n({currency_label})'),
        (16, 'Size (cm)\n尺寸（长宽高）'),
        (19, '材重'), (20, 'CBM(M3)\n方数'),
        (21, '是否带电\n（锂电or干电池）*'), (22, '是否带磁*'),
        (23, '产品图片'), (24, '产品销售链接'),
    ]
    for c, label in col_headers:
        wcell(11, c, label, Font(name='微软雅黑', size=10, bold=True),
              center_align, thin_border)
        ws.cell(row=11, column=c).fill = header_fill

    # ──────────────────────────────────────────────
    #  数据行 (Row 12-23)
    # ──────────────────────────────────────────────
    num_rows = min(len(tr.data_rows), 12)
    for i, dr in enumerate(tr.data_rows[:num_rows]):
        r = 12 + i
        ws.row_dimensions[r].height = 30
        qty = dr['I'] or 0
        box_wt = dr['B'] or 0
        unit_price = dr['H'] or 0

        data = {
            1: dr['A'],                                          # A: Box No
            2: dr['G'],                                          # B: 品名中文
            3: dr['F'],                                          # C: 品名英文
            4: dr['J'] or '',                                    # D: 材质中文
            5: dr['J'] or '',                                    # E: 材质英文
            6: str(dr['K']) if dr['K'] is not None else '',      # F: 海关编码
            7: str(dr['N']) if dr['N'] else '',                  # G: 型号
            8: dr['M'] or '',                                    # H: 品牌
            9: dr['L'] or '',                                    # I: 用途
            10: qty,                                             # J: 单箱数量
            11: qty,                                             # K: 总数量
            12: round(box_wt / qty, 3) if qty > 0 else '',       # L: 净重
            13: box_wt if box_wt else '',                        # M: 实重
            14: unit_price,                                      # N: 单价
            15: round(unit_price * qty, 2),                      # O: 总价
            16: f'{dr["C"] or 0}*{dr["D"] or 0}*{dr["E"] or 0}',# P: 尺寸
            19: round(dr['C']*dr['D']*dr['E']/5000, 1) if dr['C'] and dr['D'] and dr['E'] else '',  # S: 材重
            20: round(dr['C']*dr['D']*dr['E']/1000000, 4) if dr['C'] and dr['D'] and dr['E'] else '',  # T: CBM
            21: '是' if has_battery == '是' else '否',           # U: 带电
            22: '是' if has_magnet == '是' else '否',            # V: 带磁
            23: dr.get('Q', '') or '',                           # W: 图片
            24: dr['O'] if dr['O'] else '',                      # X: 链接
        }

        nf_map = {10: '0', 11: '0', 12: '0.000', 13: '0.0', 14: '0.00', 15: '0.00',
                  19: '0.0', 20: '0.0000'}

        for c, val in data.items():
            cell = wcell(r, c, val, normal_font, center_align, thin_border,
                         nf_map.get(c, None))

    # 空数据行保留边框
    for r in range(12, 24):
        for c in range(1, 25):
            cell = ws.cell(row=r, column=c)
            if not cell.border or cell.border == Border():
                cell.border = thin_border

    # ──────────────────────────────────────────────
    #  合计行 Row 24
    # ──────────────────────────────────────────────
    total_weight = sum(dr['B'] or 0 for dr in tr.data_rows)
    total_qty = sum(dr['I'] or 0 for dr in tr.data_rows)
    total_price = sum((dr['H'] or 0) * (dr['I'] or 0) for dr in tr.data_rows)
    total_vol_wt = round(sum(
        (dr.get('C') or 0) * (dr.get('D') or 0) * (dr.get('E') or 0) / 5000
        for dr in tr.data_rows if dr.get('C') and dr.get('D') and dr.get('E')
    ), 1) if any(dr.get('C') and dr.get('D') and dr.get('E') for dr in tr.data_rows) else 0
    total_cbm = round(sum(
        (dr.get('C') or 0) * (dr.get('D') or 0) * (dr.get('E') or 0) / 1000000
        for dr in tr.data_rows if dr.get('C') and dr.get('D') and dr.get('E')
    ), 4) if any(dr.get('C') and dr.get('D') and dr.get('E') for dr in tr.data_rows) else 0

    wcell(24, 1, '', total_font, center_align, thin_border)
    ws.merge_cells('C24:E24')
    wcell(24, 3, 'Total No.of Boxes and weight', total_font, center_align, thin_border)
    wcell(24, 11, total_qty, total_font, center_align, thin_border, '0')
    wcell(24, 13, total_weight, total_font, center_align, thin_border, '0.0')
    wcell(24, 15, total_price, total_font, center_align, thin_border, '0.00')
    wcell(24, 19, total_vol_wt, total_font, center_align, thin_border, '0.0')
    wcell(24, 20, total_cbm, total_font, center_align, thin_border, '0.0000')
    for c in range(1, 25):
        cell = ws.cell(row=24, column=c)
        cell.font = total_font
        cell.alignment = center_align
        cell.border = thin_border

    # ──────────────────────────────────────────────
    #  渠道参考列表 Row 27+ (来自原模板)
    # ──────────────────────────────────────────────
    channel_list = []
    if region == 'uk':
        channel_list = [
            '英国卡航包税（限时达）', '英国卡航不包税（限时达）', '英国卡航包税', '英国卡航不包税',
            '英国海运包税', '英国海运不包税', '英国海运不包税（卡派）',
            '欧洲卡航包税', '欧洲海运包税', '欧洲海运不包税（递延）', '欧洲铁路包税',
            '英国空运经济线包税', '英国空运经济线不包税',
            '英国空运快线包税', '英国空运快线不包税',
            '英国空运（限时达）包税', '英国空运（限时达）不包税',
            '英国空运包税（带电/化妆品）', '英国空运不包税（带电/化妆品）',
            '欧洲空运包税（普货）', '欧洲空运包税（带电/化妆品）',
        ]
    else:
        channel_list = [
            'FBA欧洲卡航DPD', '私人地址欧洲卡航DPD', 'FBA欧洲卡航UPS', '私人地址欧洲卡航UPS',
            '欧洲卡航FBA专仓卡派（kg）', '欧洲卡航万邑通/谷仓/4PX专仓卡派（kg）', '欧洲卡航卡派', '欧卡海外仓服务',
            'FBA欧洲海运DPD', '私人地址欧洲海运DPD', 'FBA欧洲海运UPS', '私人地址欧洲海运UPS',
            '欧洲海运FBA专仓卡派（kg）', '欧洲海运万邑通/谷仓/4PX专仓卡派（kg）', '欧洲海运卡派', '欧海海外仓服务',
            'FBA欧洲铁路DPD', '私人地址欧洲铁路DPD', 'FBA欧洲铁路UPS', '私人地址欧洲铁路UPS',
            '欧洲铁路FBA专仓卡派（kg）', '欧洲铁路万邑通/谷仓/4PX专仓卡派（kg）', '欧洲铁路卡派', '欧铁海外仓服务',
        ]
    wcell(27, 1, '【渠道参考列表】', bold_font)
    for i, ch in enumerate(channel_list):
        wcell(27 + i, 2, ch, normal_font)

    # ── 保存 ──
    wb.save(output_path)
    region_label = '英国' if region == 'uk' else '欧洲'
    print(f'✅ 航乐{region_label}发票已生成: {os.path.basename(output_path)}')
    print(f'   数据: {num_rows} 箱, 总价: {total_price:.2f} {currency_label}')
    return True


# ═══════════════════════════════════════════════════════════════
#  4. 批处理 — 目录中所有 TR 发票
# ═══════════════════════════════════════════════════════════════

def batch_convert(input_dir, output_dir, target='天图'):
    """转换目录中所有 TR 发票"""
    if not os.path.isdir(input_dir):
        print(f'❌ 输入目录不存在: {input_dir}')
        return

    os.makedirs(output_dir, exist_ok=True)

    # 找到所有TR发票文件
    files = [f for f in os.listdir(input_dir)
             if f.endswith('.xlsx') and '订单' not in f
             and '模板' not in f and '天图' not in f
             and '航乐' not in f]
    files.sort()

    if not files:
        print(f'⚠️  未找到 .xlsx 文件 (排除模板和已转换的)')
        return

    success = 0
    for fn in files:
        in_path = os.path.join(input_dir, fn)
        ext = ''
        if target == '天图':
            ext = f'-{target}.xlsx'
        elif target in ('航乐-uk', '航乐-eu'):
            region = 'uk' if target == '航乐-uk' else 'eu'
            ext = f'-{target}.xlsx'
        else:
            ext = f'-{target}.xlsx'

        base = os.path.splitext(fn)[0]
        out_name = f'{base}{ext}'
        out_path = os.path.join(output_dir, out_name)

        try:
            tr = TRInvoice(in_path)
            if target == '天图':
                ok = convert_to_tiantu(tr, out_path)
            elif target == '航乐-uk':
                ok = convert_to_hangle(tr, out_path, region='uk')
            elif target == '航乐-eu':
                ok = convert_to_hangle(tr, out_path, region='eu')
            else:
                print(f'❌ 未知目标格式: {target}')
                return
            if ok:
                success += 1
        except Exception as e:
            print(f'❌ 转换失败 {fn}: {e}')

    print(f'\n📊 完成: {success}/{len(files)} 文件已转换')


# ═══════════════════════════════════════════════════════════════
#  5. 主入口
# ═══════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description='TR发票 → 供应商发票 转换工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python3 convert_invoice.py TR发票.xlsx --to 天图
  python3 convert_invoice.py TR发票.xlsx --to 航乐-uk 输出文件.xlsx
  python3 convert_invoice.py TR发票.xlsx --to 航乐-eu
  python3 convert_invoice.py --batch ./发票 --to 天图 --out ./输出
        """
    )
    parser.add_argument('input', nargs='?',
                        help='TR发票 .xlsx 文件路径')
    parser.add_argument('--to', '-t', default='天图',
                        choices=['天图', '航乐-uk', '航乐-eu'],
                        help='目标供应商格式 (默认: 天图)')
    parser.add_argument('output', nargs='?',
                        help='输出文件路径 (可选，默认自动生成)')
    parser.add_argument('--batch', '-b', action='store_true',
                        help='批量模式: 转换目录中所有TR发票')
    parser.add_argument('--in-dir', default=SCRIPT_DIR,
                        help='批量模式的输入目录')
    parser.add_argument('--out-dir', default=os.path.join(SCRIPT_DIR, 'output'),
                        help='批量模式的输出目录 (默认: ./output)')

    args = parser.parse_args()

    # ── 批量模式 ──
    if args.batch:
        batch_convert(args.in_dir, args.out_dir, args.to)
        return

    # ── 单文件模式 ──
    if not args.input:
        parser.print_help()
        print('\n❌ 请指定TR发票文件路径')
        sys.exit(1)

    if not os.path.exists(args.input):
        print(f'❌ 文件不存在: {args.input}')
        sys.exit(1)

    # 读取TR发票
    print(f'📂 读取: {args.input}')
    tr = TRInvoice(args.input)
    print(tr)

    # 自动生成输出文件名
    if not args.output:
        base_name = os.path.splitext(os.path.basename(args.input))[0]
        if args.to == '天图':
            args.output = os.path.join(os.path.dirname(args.input),
                                       f'{base_name}-天图.xlsx')
        elif args.to == '航乐-uk':
            args.output = os.path.join(os.path.dirname(args.input),
                                       f'{base_name}-航乐-UK.xlsx')
        elif args.to == '航乐-eu':
            args.output = os.path.join(os.path.dirname(args.input),
                                       f'{base_name}-航乐-EU.xlsx')

    # 执行转换
    if args.to == '天图':
        convert_to_tiantu(tr, args.output)
    elif args.to == '航乐-uk':
        convert_to_hangle(tr, args.output, region='uk')
    elif args.to == '航乐-eu':
        convert_to_hangle(tr, args.output, region='eu')


if __name__ == '__main__':
    main()
