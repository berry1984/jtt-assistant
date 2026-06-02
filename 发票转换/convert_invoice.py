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
import zipfile
import tempfile
import xml.etree.ElementTree as ET
from copy import copy
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import (
    Font, Alignment, Border, Side, PatternFill
)
from openpyxl.utils import get_column_letter
from io import BytesIO


# ═══════════════════════════════════════════════════════════════
#  常量
# ═══════════════════════════════════════════════════════════════

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

TIANTU_TEMPLATE = os.path.join(SCRIPT_DIR, '天图单票专用模板20260601.xlsx')
HANGLE_UK_TEMPLATE = os.path.join(SCRIPT_DIR, '航乐-客户名称 客户单号 英国发票模板9.9更新.xls')
HANGLE_EU_TEMPLATE = os.path.join(SCRIPT_DIR, '航乐-客户单号- 欧州发票模板2.26更新.xls')
HANGLE_UK_TEMPLATE_XLSX = os.path.join(SCRIPT_DIR, '航乐-客户名称 客户单号 英国发票模板9.9更新.xlsx')
HANGLE_EU_TEMPLATE_XLSX = os.path.join(SCRIPT_DIR, '航乐-客户单号- 欧州发票模板2.26更新.xlsx')

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

def convert_to_tiantu(tr, output_path, image_url_base=None):
    """
    TR发票 → 天图格式
    规则详见 TR转天图发票_转换规则说明.md

    image_url_base: 如果提供，则为 IMAGE 公式的 URL 前缀（如 http://host/temp/xxx），
                    实现 Excel 365 "放置在单元格中"。
                    如果不提供，回退到文本链接。
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

    # 收集待嵌入的图片 (cell_ref → image_bytes)
    pending_images = {}

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
            # 收集图片，保存后会通过 IMAGE() 公式嵌入单元格
            pending_images[f'M{r}'] = img_bytes
            # 占位文本（Excel 365 会以 IMAGE 公式结果覆盖显示）
            set_cell('M', r, '[图片]', font_data, center_align, thin_border)
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

    # 后处理：将图片嵌入为单元格 IMAGE() 公式（Excel 365 "放置在单元格中"）
    if pending_images:
        _embed_images_as_cell_images(output_path, pending_images, image_url_base)

    print(f'✅ 天图发票已生成: {os.path.basename(output_path)}')
    print(f'   数据: {num_data_rows} 行, 合计: {total_qty} 件, ${total_value:.2f}')
    if pending_images:
        print(f'   嵌入图片: {len(pending_images)} 张')
    return True


# ═══════════════════════════════════════════════════════════════
#  3. 航乐 转换 (通用)
# ═══════════════════════════════════════════════════════════════

def convert_to_hangle(tr, output_path, region='uk'):
    """
    TR发票 → 航乐格式 (UK 或 EU)

    基于实际模板 (.xlsx 转换版)，复制后填充数据。
    保留模板的所有格式、合并单元格、渠道参考列表等。
    """
    # ── 选择模板 ──
    if region == 'uk':
        template_path = HANGLE_UK_TEMPLATE_XLSX
        currency_label = 'GBP'
        region_label = '英国'
    else:
        template_path = HANGLE_EU_TEMPLATE_XLSX
        currency_label = 'EUR'
        region_label = '欧洲'

    if not os.path.exists(template_path):
        # 回退到旧版 .xls 路径提示
        print(f'❌ 航乐{region_label}模板 (.xlsx) 不存在: {template_path}')
        print(f'   请先运行: python3 convert_template_xls_to_xlsx.py')
        return False

    print(f'📄 航乐{region_label}模板: {template_path}')

    # ── 获取 TR 数据 ──
    country = tr.get('收件人国家代码(二字代码)', '') or tr.get('收件人国家代码', '') or 'US'
    country_map = {
        'GB': 'UK', 'UK': 'UK', 'US': 'US', 'DE': 'DE', 'FR': 'FR',
        'IT': 'IT', 'ES': 'ES', 'NL': 'NL', 'BE': 'BE', 'PL': 'PL',
        'CZ': 'CZ', 'CA': 'CA',
    }
    country_code = country_map.get(country.upper(), country.upper())
    has_battery = tr.get('带电', '否')
    has_magnet = tr.get('带磁', '否')

    # ── 解析箱号中的箱数（如 "FBA15LV645QTU000001-5" → 5 箱）──
    import re
    def parse_box_count(box_no):
        if not box_no:
            return 1
        m = re.search(r'-(\d+)$', str(box_no))
        return int(m.group(1)) if m else 1

    # ── 复制模板 ──
    shutil.copy(template_path, output_path)
    wb = load_workbook(output_path)
    ws = wb['Packing list装箱单发票']

    # ── 样式定义（用于填入的数据）──
    thin_side = Side(style='thin')
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    data_font = Font(name='微软雅黑', size=10)
    bold_font = Font(name='微软雅黑', size=10, bold=True)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    def wcell(r, c, val, font=None, align=None, border=None, nf=None):
        cell = ws.cell(row=r, column=c)
        cell.value = val
        if font: cell.font = font
        if align: cell.alignment = align
        if border: cell.border = border
        if nf: cell.number_format = nf
        return cell

    # ══════════════════════════════════════════════
    #  头部数据填充——模板已含标签，只需填值
    # ══════════════════════════════════════════════

    # ── Row 3: ADD / COUNTRY / SHIPMENT ID / AMAZON REF ──
    # B3:E3 (merged) = ADD(收件地址) 值 — correct版本用 收件人姓名(仓库代码)
    ws['B3'] = tr.get('收件人姓名', '')
    # G3 = COUNTRY 值
    ws['G3'] = country_code
    # H3:J3 (merged) = SHIPMENT ID 标签（保留模板标签，值不填或后续手动）
    # K3:M3 (merged) = AMAZON REF 标签（保留模板标签）

    # ── Row 4: ZIP / COMPANY / TEL / 报关类型 / 交货仓库 ──
    # B4 = ZIP CODE
    ws['B4'] = tr.get('收件人邮编', '')
    # D4:E4 (merged) = COMPANY
    ws['D4'] = tr.get('收件人公司', '')
    # G4:H4 (merged) = TEL
    ws['G4'] = tr.get('收件人电话', '')
    # I4:J4 (merged) = 报关类型标签（保留模板标签，值写到 Row5 I5）
    # K4:M4 (merged) = 交货仓库标签（保留模板标签，值写到 Row5 K5）

    # ── Row 5: CITY / ATTN / EMAIL / 报关退税 / 交货仓库 ──
    ws['B5'] = tr.get('收件人城市', '')
    # D5:E5 (merged) = ATTN
    ws['D5'] = tr.get('收件人姓名', '')
    # G5:H5 (merged) = EMAIL
    ws['G5'] = tr.get('收件人邮箱', '')
    # I5:J5 (merged) = 报关类型值
    ws['I5'] = tr.get('报关方式', '')
    # K5:M5 (merged) = 交货仓库值
    ws['K5'] = ''  # 交货仓库 — 按需填写

    # ── Row 7-10: VAT / 渠道 / 包税 / 物品属性 ──
    # B7:H7 = VAT公司名称/公司名称
    ws['B7'] = tr.get('VAT公司英文名', '')
    # B8:H8 = VAT号
    ws['B8'] = tr.get('VAT号', '')
    # B9:H9 = EORI号
    ws['B9'] = tr.get('EORI号', '')
    # B10:H10 = VAT注册地址
    ws['B10'] = tr.get('VAT注册地址', '')
    # I7:K10 (merged) = 渠道（服务名）
    ws['I7'] = tr.get('服务', '')
    # L7:L10 (merged) = 是否包税
    ws['L7'] = tr.get('交税方式', '')
    # M7:M10 (merged) = 物品属性（从中文品名推断）
    category = _guess_product_category(tr)
    ws['M7'] = category

    # ══════════════════════════════════════════════
    #  数据行 (Row 12+)
    # ══════════════════════════════════════════════

    # 清空模板数据区域 (row 12-23)
    for r in range(12, 50):
        for c in range(1, 30):
            cell = ws.cell(row=r, column=c)
            # 保留公式提示文字（合计行标签）
            if r == 24 and c == 3 and str(cell.value or '').strip() == 'Total No.of Boxes and weight':
                cell.value = None
            else:
                cell.value = None

    # 添加 PO Number 表头（模板不含，正确版本有此列）
    ws['Y11'] = 'PO Number*\n（Reference ID）'

    data_start = 12
    for i, dr in enumerate(tr.data_rows):
        r = data_start + i
        ws.row_dimensions[r].height = 30

        qty = dr['I'] or 0
        box_wt = dr['B'] or 0
        unit_price = dr['H'] or 0
        box_count = parse_box_count(dr['A'])
        total_qty = qty * box_count

        # 尺寸 — 取整（匹配正确版本做法）
        length_cm = int(dr['C']) if dr['C'] else 0
        width_cm = int(dr['D']) if dr['D'] else 0
        height_cm = int(dr['E']) if dr['E'] else 0

        # 材重除数：欧洲 6000，英国 5000
        vol_div = 6000 if region == 'eu' else 5000
        raw_vol_wt = (length_cm * width_cm * height_cm / vol_div) if length_cm and width_cm and height_cm else None
        raw_cbm = (length_cm * width_cm * height_cm / 1000000) if length_cm and width_cm and height_cm else None
        # 净重 = 实重 / 单箱数量（不是总数量）
        net_wt = round(box_wt / qty, 4) if box_wt and qty else None

        data = {
            1: str(dr['A']) if dr['A'] is not None else '',                      # A: Box No
            2: str(dr['G']) if dr['G'] else '',                                  # B: 品名中文
            3: str(dr['F']) if dr['F'] else '',                                  # C: 品名英文
            4: str(dr['J']) if dr['J'] else '',                                  # D: 材质中文
            5: str(dr['J']) if dr['J'] else '',                                  # E: 材质英文
            6: str(int(dr['K'])) if dr['K'] is not None else '',                 # F: 海关编码
            7: str(dr['N']) if dr['N'] else '无',                                # G: 型号
            8: str(dr['M']) if dr['M'] else '无品牌',                            # H: 品牌
            9: str(dr['L']) if dr['L'] else '',                                  # I: 用途
            10: qty,                                                             # J: 单箱数量
            11: total_qty,                                                       # K: 总数量
            12: net_wt,                                                          # L: 净重
            13: box_wt,                                                          # M: 实重
            14: unit_price,                                                      # N: 单价
            15: round(unit_price * total_qty, 2),                                # O: 总价
            16: length_cm,                                                       # P: 长
            17: width_cm,                                                        # Q: 宽
            18: height_cm,                                                       # R: 高
            19: raw_vol_wt,                                                       # S: 材重
            20: raw_cbm,                                                          # T: CBM
            21: '是' if has_battery == '是' else '否',                            # U: 是否带电
            22: '是' if has_magnet == '是' else '否',                             # V: 是否带磁
            23: str(dr.get('Q', '') or ''),                                      # W: 产品图片
            24: str(dr['O']) if dr['O'] else '',                                 # X: 链接
            25: str(dr.get('V', '') or ''),                                      # Y: PO Number
        }

        nf_map = {10: '0', 11: '0', 12: '0.0000', 13: '0.0',
                  14: '0.00', 15: '0.00', 19: '0.000', 20: '0.000000'}

        for col_idx, val in data.items():
            if val is None or val == '':
                continue
            wcell(r, col_idx, val, data_font, center_align, thin_border,
                  nf_map.get(col_idx, None))

    num_data = len(tr.data_rows)

    # ══════════════════════════════════════════════
    #  合计行
    # ══════════════════════════════════════════════
    total_row = data_start + num_data + 3  # 空3行间隙

    # 计算合计
    total_qty = sum((dr['I'] or 0) * parse_box_count(dr['A']) for dr in tr.data_rows)
    total_weight = sum(dr['B'] or 0 for dr in tr.data_rows)
    total_price = sum((dr['H'] or 0) * (dr['I'] or 0) * parse_box_count(dr['A']) for dr in tr.data_rows)

    vol_div = 6000 if region == 'eu' else 5000
    vol_wts = []
    cbms = []
    for dr in tr.data_rows:
        l = int(dr['C']) if dr['C'] else 0
        w = int(dr['D']) if dr['D'] else 0
        h = int(dr['E']) if dr['E'] else 0
        if l and w and h:
            vol_wts.append(l * w * h / vol_div)
            cbms.append(l * w * h / 1000000)
    total_vol_wt = round(sum(vol_wts), 3) if vol_wts else 0
    total_cbm = round(sum(cbms), 6) if cbms else 0

    # 写入合计
    ws.merge_cells(start_row=total_row, start_column=3, end_row=total_row, end_column=5)
    ws.cell(row=total_row, column=3).value = 'Total No.of Boxes and weight'

    total_fields = {
        11: (total_qty, '0'),
        13: (total_weight, '0.0'),
        15: (total_price, '0.00'),
        19: (total_vol_wt, '0.000'),
        20: (total_cbm, '0.000000'),
    }
    for col_idx, (val, nf) in total_fields.items():
        cell = ws.cell(row=total_row, column=col_idx)
        cell.value = val
        cell.font = bold_font
        cell.alignment = center_align
        cell.number_format = nf

    # 合计行加粗+边框
    for c in range(1, 26):
        cell = ws.cell(row=total_row, column=c)
        if cell.font and cell.font.name:
            cell.font = Font(name=cell.font.name, size=10, bold=True)
        else:
            cell.font = bold_font
        cell.alignment = center_align

    # ── 保存 ──
    wb.save(output_path)
    print(f'✅ 航乐{region_label}发票已生成: {os.path.basename(output_path)}')
    print(f'   数据: {num_data} 行, 总件数: {total_qty}, 总价: {total_price:.2f} {currency_label}')
    return True


def _guess_product_category(tr):
    """
    从 TR 发票的多个产品中文品名推断物品属性。
    例如所有产品含"灯" → "灯类"，否则合并或不填。
    """
    import re
    all_names = [dr.get('G', '') or '' for dr in tr.data_rows if dr.get('G')]
    if not all_names:
        return ''

    # 常见分类关键词
    keywords = {
        '灯类': ['灯', '照明'],
        '服装': ['衣', '服', '裤', '裙', '袜', '帽', '鞋'],
        '电子': ['电子', '电', '机', '器', '充电', '电池', '线'],
        '家居': ['家具', '桌', '椅', '凳', '架', '柜', '收纳'],
        '玩具': ['玩具', '玩'],
        '箱包': ['包', '箱', '袋', '背包'],
        '美容': ['美容', '化妆', '护肤', '梳'],
    }

    scores = {}
    for cat, kws in keywords.items():
        score = 0
        for name in all_names:
            for kw in kws:
                if kw in name:
                    score += 1
                    break
        if score > 0:
            scores[cat] = score

    if scores:
        best = max(scores, key=scores.get)
        return best

    return '普货'


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
#  图片嵌入后处理 — 将图片嵌入为单元格 IMAGE() 公式
#  (Excel 365 "Place in Cell" / "放置在单元格中")
# ═══════════════════════════════════════════════════════════════

def _embed_images_as_cell_images(xlsx_path, cell_image_map, image_url_base=None):
    """
    后处理 xlsx 文件，将图片嵌入为单元格值。

    始终做两件事：
      1. 图片写入 xl/media/（离线后备，文件自带图片）
      2. 单元格写入 _xlfn.IMAGE() 公式（Excel 365 "放置在单元格中"）

    image_url_base: 如果提供，公式引用服务器 URL（在线时显示图片）；
                    如果不提供，用 "0#" 内部引用（部分 Excel 版本可能不可用）。
    """
    if not cell_image_map:
        return

    ET.register_namespace('', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    NS_S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types'

    tmp_dir = tempfile.mkdtemp()
    try:
        # ── 1. 解压 xlsx ──
        with zipfile.ZipFile(xlsx_path, 'r') as z:
            z.extractall(tmp_dir)

        sorted_refs = list(cell_image_map.items())

        # ── 2. 写入图片到 xl/media/（离线后备）──
        media_dir = os.path.join(tmp_dir, 'xl', 'media')
        os.makedirs(media_dir, exist_ok=True)
        img_filenames = {}
        for i, (cell_ref, img_bytes) in enumerate(sorted_refs):
            img_filename = f'image_tiantu_{i+1}.png'
            with open(os.path.join(media_dir, img_filename), 'wb') as f:
                f.write(img_bytes)
            img_filenames[cell_ref] = img_filename

        # ── 3. 修改 sheet1.xml：写入 IMAGE 公式 ──
        sheet_path = os.path.join(tmp_dir, 'xl', 'worksheets', 'sheet1.xml')
        if not os.path.exists(sheet_path):
            print(f'  ⚠️  sheet1.xml 不存在，跳过图片嵌入')
            return

        tree = ET.parse(sheet_path)
        root = tree.getroot()

        for cell_ref, img_bytes in sorted_refs:
            img_filename = img_filenames[cell_ref]
            cell = root.find(f'.//{{{NS_S}}}c[@r="{cell_ref}"]')
            if cell is not None:
                for child in list(cell):
                    cell.remove(child)
                f_elem = ET.SubElement(cell, f'{{{NS_S}}}f')
                if image_url_base:
                    f_elem.text = f'_xlfn.IMAGE("{image_url_base}/{img_filename}", "", 0)'
                else:
                    f_elem.text = f'_xlfn.IMAGE("0#{img_filename}", "", 0)'
                if 't' in cell.attrib:
                    del cell.attrib['t']
            else:
                print(f'  ⚠️  未找到单元格 {cell_ref}')

        tree.write(sheet_path, xml_declaration=True, encoding='UTF-8')

        # ── 4. 补充 Content_Types ──
        ct_path = os.path.join(tmp_dir, '[Content_Types].xml')
        if os.path.exists(ct_path):
            ct_tree = ET.parse(ct_path)
            ct_root = ct_tree.getroot()
            exists = False
            for default in ct_root.findall(f'{{{NS_CT}}}Default'):
                if default.get('Extension') == 'png':
                    exists = True
                    break
            if not exists:
                png_default = ET.SubElement(ct_root, f'{{{NS_CT}}}Default')
                png_default.set('Extension', 'png')
                png_default.set('ContentType', 'image/png')
            ct_tree.write(ct_path, xml_declaration=True, encoding='UTF-8')

        # ── 5. 添加 sheet → media 关系 ──
        sheet_rels_dir = os.path.join(tmp_dir, 'xl', 'worksheets', '_rels')
        os.makedirs(sheet_rels_dir, exist_ok=True)
        rels_path = os.path.join(sheet_rels_dir, 'sheet1.xml.rels')

        next_rId = 1
        existing_rels = {}
        if os.path.exists(rels_path):
            rels_tree = ET.parse(rels_path)
            rels_root = rels_tree.getroot()
            for rel_elem in rels_root:
                rid = rel_elem.get('Id', '')
                if rid.startswith('rId'):
                    try:
                        num = int(rid[3:])
                        next_rId = max(next_rId, num + 1)
                    except:
                        pass
                existing_rels[rid] = rel_elem
        else:
            rels_root = ET.Element('Relationships')
            rels_root.set('xmlns', NS_R)

        existing_targets = set()
        for rid, elem in existing_rels.items():
            tgt = elem.get('Target', '')
            existing_targets.add(tgt)

        for cell_ref, img_bytes in sorted_refs:
            img_filename = img_filenames[cell_ref]
            rel_target = f'../media/{img_filename}'
            if rel_target in existing_targets:
                continue
            rel_elem = ET.SubElement(rels_root, 'Relationship')
            rel_elem.set('Id', f'rId{next_rId}')
            rel_elem.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
            rel_elem.set('Target', rel_target)
            existing_targets.add(rel_target)
            next_rId += 1

        rels_tree = ET.ElementTree(rels_root)
        rels_tree.write(rels_path, xml_declaration=True, encoding='UTF-8')

        # ── 6. 重新打包 ──
        final_path = xlsx_path + '.tmp'
        with zipfile.ZipFile(final_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for dirpath, dirnames, filenames in os.walk(tmp_dir):
                for fn in filenames:
                    full_path = os.path.join(dirpath, fn)
                    arcname = os.path.relpath(full_path, tmp_dir)
                    zout.write(full_path, arcname)

        shutil.move(final_path, xlsx_path)

        mode = f'IMAGE 公式 + 服务器 URL + xl/media/ 后备' if image_url_base else '0# 内部引用 + xl/media/'
        print(f'  📷 已嵌入 {len(cell_image_map)} 张图片（{mode}）')

    except Exception as e:
        print(f'  ⚠️  图片嵌入后处理出错: {e}')
        import traceback
        traceback.print_exc()
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


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
