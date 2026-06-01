#!/usr/bin/env python3
"""
番茄钟账单生成器
用途：根据 订单列表、拣货数据、应收价格 三个Excel文件，生成标准格式账单

用法：
  python3 gen_bill.py <订单列表.xlsx> <拣货数据.xlsx> <应收价格.xlsx> [输出文件名.xlsx]

规则：
  - J列(计费重) = ROUND(拣货收费重) 并调整最大项使合计匹配订单列表
  - S列(报关费) = 350/1.06, 每个唯一走货渠道收取一次（合并显示）
  - 保留完整浮点精度（使用公式）
  - 排序：按渠道分组 → SO → FBA
"""

import sys, re, os, shutil
from datetime import datetime, timedelta
from collections import defaultdict, OrderedDict
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy

TEMPLATE = None  # Will look for 账单模板.xlsx next to script or in cwd

# ── Column mapping ──
COL = {chr(65+i): i+1 for i in range(26)}  # A=1, B=2, ...
COL.update({'AA': 27, 'AB': 28, 'AC': 29})

def find_template():
    """Find template file"""
    candidates = [
        os.path.join(os.path.dirname(__file__), '账单模板.xlsx'),
        os.path.join(os.getcwd(), '账单模板.xlsx'),
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return None

def load_data(order_path, pick_path, price_path):
    """Load and parse three input files"""
    
    # ── Order list ──
    wb = load_workbook(order_path)
    ws = wb.active
    h = [c.value for c in list(ws.iter_rows(min_row=1, max_row=1))[0]]
    orders = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        d = dict(zip(h, r))
        orders[d['运单号']] = d
    
    # ── Picking data ──
    wb = load_workbook(pick_path)
    ws = wb.active
    picks = list(ws.iter_rows(min_row=2, values_only=True))
    
    # ── Prices ──
    wb = load_workbook(price_path)
    ws = wb.active
    prices = {}
    price_rows_raw = []  # keep raw rows (with formulas) for 报价表A
    for r in ws.iter_rows(min_row=2, values_only=False):
        vals = [c.value for c in r]
        ch, wh, price, customs = vals[0], vals[1], vals[2], vals[3] if len(vals) > 3 else None
        if wh:
            prices[wh] = {'channel': ch, 'price': price, 'customs_fee': customs or 0}
            price_rows_raw.append(vals)

    return orders, picks, prices, price_rows_raw

def lookup_price(prices, wh_code):
    """Look up price by warehouse code, with suffix matching"""
    if not wh_code:
        return {}
    p = prices.get(wh_code, {})
    if not p:
        prefix = wh_code.split('-')[0]
        p = prices.get(prefix, {})
    if not p:
        for key in prices:
            if wh_code.startswith(key):
                p = prices[key]
                break
    return p

def build_rows(orders, picks, prices):
    """Build bill rows from source data"""

    pick_groups = defaultdict(lambda: defaultdict(list))
    for r in picks:
        so = r[1]
        m = re.match(r'(.+?)U\d+$', str(r[2]))
        fba = m.group(1) if m else str(r[2])
        pick_groups[so][fba].append(r)

    def to_num(v, default=0):
        """Safely convert Excel value to float"""
        if v is None:
            return float(default)
        if isinstance(v, (int, float)):
            return float(v)
        try:
            return float(str(v).strip())
        except (ValueError, TypeError):
            return float(default)

    def to_int(v, default=0):
        return int(to_num(v, default))

    def clean_wh(code):
        """Clean warehouse code: strip suffixes like -Amazon"""
        return code.split('-')[0] if code else (code or '')

    all_rows = []

    for so, o in sorted(orders.items()):
        service = o['服务']
        date_val = o.get('发货日期', o.get('工作日期', ''))
        wh_code = clean_wh(o.get('收件人', ''))
        total_weight_order = to_num(o.get('收费重', 0))
        fba_ext = o['扩展单号'].split(',') if o['扩展单号'] else []
        so_picks = pick_groups.get(so, {})

        if so_picks:
            fba_rows = []
            for fba, rows in so_picks.items():
                total_w = sum(to_num(r[8]) for r in rows)
                fba_rows.append({
                    'fba': fba,
                    'boxes': len(rows),
                    'length': to_num(rows[0][5]),
                    'width': to_num(rows[0][4]),
                    'height': to_num(rows[0][3]),
                    'weight_raw': total_w,
                })

            for r in fba_rows:
                r['weight'] = round(to_num(r['weight_raw']))

            rounded_sum = sum(r['weight'] for r in fba_rows)
            diff = total_weight_order - rounded_sum

            if diff != 0 and fba_rows:
                fba_rows_sorted = sorted(enumerate(fba_rows),
                    key=lambda x: (-x[1]['weight_raw'], x[1]['fba']))
                remaining, sign = abs(diff), 1 if diff > 0 else -1
                for idx, item in fba_rows_sorted:
                    if remaining <= 0: break
                    item['weight'] += sign * 1
                    remaining -= 1

            for r in fba_rows:
                p = lookup_price(prices, wh_code)
                r.update({'so': so, 'service': service, 'date': date_val,
                          'wh': wh_code, 'unit_price': to_num(p.get('price', 0)) if p else 0})
            all_rows.extend(fba_rows)
        else:
            p = lookup_price(prices, wh_code)
            all_rows.append({
                'so': so, 'fba': fba_ext[0] if fba_ext else '',
                'date': date_val, 'service': service, 'wh': wh_code,
                'boxes': to_int(o.get('件数', 0)),
                'length': 0, 'width': 0, 'height': 0,
                'weight': total_weight_order, 'weight_raw': total_weight_order,
                'unit_price': to_num(p.get('price', 0)) if p else 0,
            })

    # Filter out rows with zero weight (no actual cargo)
    all_rows = [r for r in all_rows if r['weight'] > 0]

    return all_rows

def sort_rows(rows):
    """Sort rows: by channel group, then SO, then FBA"""
    # Define channel group order
    channel_order = OrderedDict()
    seen = []
    for r in rows:
        ch = r['service']
        if ch not in seen:
            seen.append(ch)
    
    def sort_key(r):
        ch_idx = seen.index(r['service']) if r['service'] in seen else 999
        return (ch_idx, r['so'], r['fba'])
    
    return sorted(rows, key=sort_key)

def generate_bill(rows, output_path, template_path=None, title_str=None, date_range_str=None, price_rows_raw=None, year=None):
    """Generate bill Excel file"""
    
    if template_path is None:
        template_path = find_template()
    if not template_path:
        print("ERROR: 账单模板.xlsx not found!")
        sys.exit(1)
    
    # Copy template
    shutil.copy(template_path, output_path)
    wb = load_workbook(output_path)
    ws = wb['5月人民币账单（已调格式）']

    # Capture template column fills from row 4 (first data row) for style preservation
    from openpyxl.styles import PatternFill
    template_fills = {}
    for c in range(1, 29):
        src = ws.cell(row=4, column=c)
        f = src.fill
        if f.patternType and f.patternType != 'none':
            try:
                template_fills[c] = PatternFill(patternType=f.patternType,
                    fgColor=src.fill.fgColor.rgb if f.fgColor else None)
            except:
                template_fills[c] = PatternFill(patternType='solid', fgColor='FFFFFFFF')

    # ── Unmerge old cells in data area ──
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row >= 4:
            ws.unmerge_cells(str(mr))
    for row in range(4, 30):
        for col in range(1, 29):
            try:
                ws.cell(row=row, column=col).value = None
            except AttributeError:
                pass

    # ── Styles ──
    thin_border = Border(
        left=Side(style='hair'), right=Side(style='hair'),
        top=Side(style='hair'), bottom=Side(style='hair'))
    data_font = Font(name='微软雅黑', size=9)
    bold_font = Font(name='微软雅黑', size=9, bold=True)
    center = Alignment(horizontal='center', vertical='center')
    
    n = len(rows)
    
    # ── Write data rows ──
    for i, r in enumerate(rows):
        row_num = 4 + i
        ws.row_dimensions[row_num].height = 20
        
        # Format date: serial → "5月23日"
        if isinstance(r['date'], (int, float)):
            dt = datetime(1899, 12, 30) + timedelta(days=r['date'])
            date_str = f"{dt.month}月{dt.day}日"
        else:
            date_str = str(r['date'])
        vals = [
            ('A', date_str, '@'),
            ('B', r['so'], '@'),
            ('C', r['fba'], '@'),
            ('D', r['service'], '@'),
            ('E', r['wh'], '@'),
            ('F', r['boxes'], '0'),
            ('G', r['length'] if r['length'] else '', '0'),
            ('H', r['width'] if r['width'] else '', '0'),
            ('I', r['height'] if r['height'] else '', '0'),
            ('J', r['weight'], '0'),
            ('K', r['unit_price'], '0.00'),
        ]
        
        for col_l, val, nf in vals:
            cell = ws[f'{col_l}{row_num}']
            cell.value = val
            cell.font = data_font
            cell.alignment = center
            cell.border = thin_border
            if nf != '@':
                cell.number_format = nf
            # Fill applied in post-processing below
        
        # L = K*0.07/1.06, M = K*0.35, N = K*0.58 (formulas)
        for col_l, formula in [('L', f'=K{row_num}*0.07/1.06'),
                                ('M', f'=K{row_num}*0.35'),
                                ('N', f'=K{row_num}*0.58')]:
            cell = ws[f'{col_l}{row_num}']
            cell.value = formula
            cell.font = data_font
            cell.alignment = center
            cell.border = thin_border
            cell.number_format = '0.00'
        
        # O = L*J, P = O*0.06, Q = M*J, R = N*J (formulas)
        for col_l, formula in [('O', f'=L{row_num}*J{row_num}'),
                                ('P', f'=O{row_num}*0.06'),
                                ('Q', f'=M{row_num}*J{row_num}'),
                                ('R', f'=N{row_num}*J{row_num}')]:
            cell = ws[f'{col_l}{row_num}']
            cell.value = formula
            cell.font = data_font
            cell.alignment = center
            cell.border = thin_border
            cell.number_format = '#,##0.00'
        
        # S, T (customs fee - filled per channel group later)
        # U-Y = 0
        for col_l in ['U', 'V', 'W', 'X', 'Y']:
            cell = ws[f'{col_l}{row_num}']
            cell.value = 0
            cell.font = data_font
            cell.alignment = center
            cell.border = thin_border
            cell.number_format = '#,##0.00'
        
        # Z = RMB
        ws[f'Z{row_num}'].value = 'RMB'
        ws[f'Z{row_num}'].font = data_font
        ws[f'Z{row_num}'].alignment = center
        ws[f'Z{row_num}'].border = thin_border
        
        # AA = SUM(O:Y) (formula)
        ws[f'AA{row_num}'].value = f'=SUM(O{row_num}:Y{row_num})'
        ws[f'AA{row_num}'].font = data_font
        ws[f'AA{row_num}'].alignment = center
        ws[f'AA{row_num}'].border = thin_border
        ws[f'AA{row_num}'].number_format = '#,##0.00'
        
        # AB blank
        ws[f'AB{row_num}'].border = thin_border

    # Apply template column fills to all data rows
    for rn in range(4, 4 + n):
        # A column border
        ws[f'A{row_num}'].border = thin_border
    
    # ── Customs fee: one per unique channel ──
    # Find channel groups and their row ranges
    channel_groups = []
    current_ch = None
    start_row = None
    
    for i, r in enumerate(rows):
        row_num = 4 + i
        ch = r['service']
        if ch != current_ch:
            if current_ch is not None:
                channel_groups.append((current_ch, start_row, row_num - 1))
            current_ch = ch
            start_row = row_num
    if current_ch is not None:
        channel_groups.append((current_ch, start_row, 4 + n - 1))
    
    # Apply customs fee (350/1.06) to first row of each channel group, merge S column
    for ch, s_row, e_row in channel_groups:
        # First row gets the formula
        cell_s = ws[f'S{s_row}']
        cell_s.value = '=350/1.06'
        cell_s.font = data_font
        cell_s.alignment = center
        cell_s.border = thin_border
        cell_s.number_format = '#,##0.00'
        if 19 in template_fills:
            cell_s.fill = template_fills[19]

        cell_t = ws[f'T{s_row}']
        cell_t.value = f'=S{s_row}*0.06'
        cell_t.font = data_font
        cell_t.alignment = center
        cell_t.border = thin_border
        cell_t.number_format = '#,##0.00'
        if 20 in template_fills:
            cell_t.fill = template_fills[20]
        
        # Merge S and T columns across all rows of this channel group
        if e_row > s_row:
            ws.merge_cells(f'S{s_row}:S{e_row}')
            ws.merge_cells(f'T{s_row}:T{e_row}')
        
        # Add borders to merged cells
        for rn in range(s_row, e_row + 1):
            for cl in ['S', 'T']:
                cell = ws[f'{cl}{rn}']
                cell.border = thin_border

    # Apply template column fills to ALL data rows (after customs section)
    # ── Summary row (with blank separator before it, like correct bill) ──
    sr = 4 + n + 1  # blank row at 4+n, 合计 at sr
    info_row = sr + 1

    ws.merge_cells(f'B{sr}:E{sr}')
    ws[f'B{sr}'].value = '合计'
    ws[f'B{sr}'].font = bold_font
    ws[f'B{sr}'].alignment = center

    sum_cols = ['F', 'J', 'O', 'P', 'Q', 'R', 'S', 'T', 'W', 'AA']
    for cl in sum_cols:
        ws[f'{cl}{sr}'].value = f'=SUM({cl}4:{cl}{sr-2})'  # data up to last data row (sr-2), blank sr-1 excluded
        ws[f'{cl}{sr}'].font = bold_font
        ws[f'{cl}{sr}'].alignment = center
        ws[f'{cl}{sr}'].border = thin_border
        ws[f'{cl}{sr}'].number_format = '#,##0.00'

    for cl in ['A','B','C','D','E','G','H','I','K','L','M','N','U','V','X','Y','Z','AB']:
        ws[f'{cl}{sr}'].border = thin_border
        ws[f'{cl}{sr}'].font = bold_font
        ws[f'{cl}{sr}'].alignment = center
    
    # ── Bank info ──
    ws.merge_cells(f'A{info_row}:AA{info_row}')
    ws[f'A{info_row}'].value = '以下账户信息为我司唯一合法收款账户，如付款至其他账户我司概不负责!'
    ws[f'A{info_row}'].font = Font(name='微软雅黑', size=9, color='FF0000')
    ws[f'A{info_row}'].alignment = Alignment(horizontal='left', vertical='center')
    
    acct_row = info_row + 1
    bank_text = ('对公账户：\n'
        '账户名：赛诺吉(深圳)国际货运代理有限公司\n'
        '公司地址：深圳市福田区沙头街道沙嘴社区沙嘴路8号红树华府A栋35层3504、3505、3506\n'
        '联系电话：0755-82720817\n'
        '账  号：41009000040015688\n'
        '开户行：中国农业银行深圳福田保税区支行')
    ws.merge_cells(f'A{acct_row}:AA{acct_row}')
    ws[f'A{acct_row}'].value = bank_text
    ws[f'A{acct_row}'].font = Font(name='微软雅黑', size=9)
    ws[f'A{acct_row}'].alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    ws.row_dimensions[acct_row].height = 110
    
    note_row = acct_row + 1
    ws.merge_cells(f'A{note_row}:AA{note_row}')
    ws[f'A{note_row}'].value = '请勿用私人账户付到我司公账，谢谢配合。'
    ws[f'A{note_row}'].font = Font(name='微软雅黑', size=9, color='FF0000')
    ws[f'A{note_row}'].alignment = Alignment(horizontal='left', vertical='center')
    
    # ── Update title ──
    if title_str:
        ws['A2'].value = title_str
    
    # ── 开票金额 sheet ──
    ws_inv = wb['开票金额']
    for r in range(2, 6):
        for c in range(1, 6):
            ws_inv.cell(row=r, column=c).value = None
    
    # Write formulas referencing the bill sheet
    sheet_name = '5月人民币账单（已调格式）'
    ref_row = sr  # subtotal row
    
    inv_data = [
        ('国际货物运输代理服务', f"='{sheet_name}'!Q{ref_row}+'{sheet_name}'!R{ref_row}+'{sheet_name}'!U{ref_row}+'{sheet_name}'!W{ref_row}", '免税', 0),
        ('国内货物运输代理服务', f"='{sheet_name}'!O{ref_row}", 0.06, f"='{sheet_name}'!P{ref_row}"),
        ('经纪代理服务-报关费', f"='{sheet_name}'!S{ref_row}", 0.06, f"='{sheet_name}'!T{ref_row}"),
    ]
    
    ib = Border(left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin'))
    
    for i, (name, amt_formula, rate, tax_formula) in enumerate(inv_data):
        r = 2 + i
        ws_inv.cell(row=r, column=1, value=name).font = data_font
        ws_inv.cell(row=r, column=1).alignment = center
        ws_inv.cell(row=r, column=1).border = ib
        
        ws_inv.cell(row=r, column=2, value=amt_formula if isinstance(amt_formula, str) else amt_formula).font = data_font
        ws_inv.cell(row=r, column=2).alignment = center
        ws_inv.cell(row=r, column=2).number_format = '#,##0.00'
        ws_inv.cell(row=r, column=2).border = ib
        
        ws_inv.cell(row=r, column=3, value=rate if rate == '免税' else rate).font = data_font
        ws_inv.cell(row=r, column=3).alignment = center
        ws_inv.cell(row=r, column=3).border = ib
        if rate != '免税':
            ws_inv.cell(row=r, column=3).number_format = '0%'
        
        ws_inv.cell(row=r, column=4, value=tax_formula if isinstance(tax_formula, str) else tax_formula).font = data_font
        ws_inv.cell(row=r, column=4).alignment = center
        ws_inv.cell(row=r, column=4).number_format = '#,##0.00'
        ws_inv.cell(row=r, column=4).border = ib
        
        # 合计 = B + D
        ws_inv.cell(row=r, column=5, value=f'=B{r}+D{r}').font = data_font
        ws_inv.cell(row=r, column=5).alignment = center
        ws_inv.cell(row=r, column=5).number_format = '#,##0.00'
        ws_inv.cell(row=r, column=5).border = ib
    
    # Total row
    ws_inv.cell(row=5, column=1, value='总额').font = bold_font
    ws_inv.cell(row=5, column=1).alignment = center
    ws_inv.cell(row=5, column=1).border = ib
    ws_inv.cell(row=5, column=2, value='=B2+B3+B4').font = bold_font
    ws_inv.cell(row=5, column=2).alignment = center
    ws_inv.cell(row=5, column=2).number_format = '#,##0.00'
    ws_inv.cell(row=5, column=2).border = ib
    ws_inv.cell(row=5, column=4, value='=D2+D3+D4').font = bold_font
    ws_inv.cell(row=5, column=4).alignment = center
    ws_inv.cell(row=5, column=4).number_format = '#,##0.00'
    ws_inv.cell(row=5, column=4).border = ib
    ws_inv.cell(row=5, column=5, value='=E2+E3+E4').font = bold_font
    ws_inv.cell(row=5, column=5).alignment = center
    ws_inv.cell(row=5, column=5).number_format = '#,##0.00'
    ws_inv.cell(row=5, column=5).border = ib

    # ── 报价表A: update with actual price data ──
    ws_quote = wb['报价表A']
    # Clear old price rows (keep headers R1-R3)
    for r in range(4, 30):
        for c in range(1, 9):
            ws_quote.cell(row=r, column=c).value = None

    # Determine week info from date range
    week_str = ''
    month = datetime.now().month  # fallback if no date_range_str
    if date_range_str:
        yr = year or 2026
        # Extract month and start_day (handle "6月" or "6.1-6.7")
        nums = [int(x) for x in re.findall(r'\d+', date_range_str)]
        month = nums[0] if len(nums) > 0 else month
        start_day = nums[1] if len(nums) > 1 else 1
        from datetime import date as dt_date
        monday_dt = dt_date(yr, month, start_day)
        first_of_month = dt_date(yr, month, 1)
        days_to_monday = (7 - first_of_month.weekday()) % 7
        first_monday = first_of_month + timedelta(days=days_to_monday) if days_to_monday else first_of_month
        week_num = 1 if monday_dt < first_monday else 2 + (monday_dt - first_monday).days // 7
        week_map = {1:'第一周', 2:'第二周', 3:'第三周', 4:'第四周', 5:'第五周', 6:'第六周'}
        week_str = week_map.get(week_num, f'第{week_num}周')

    # Write price data rows
    if price_rows_raw:
        for i, vals in enumerate(price_rows_raw):
            r = 4 + i
            ch, wh, price, customs = vals[0], vals[1], vals[2], vals[3] if len(vals) > 3 else None
            if not wh:
                continue
            ws_quote.cell(row=r, column=1, value=month).font = data_font
            ws_quote.cell(row=r, column=1).alignment = center
            ws_quote.cell(row=r, column=2, value=week_str).font = data_font
            ws_quote.cell(row=r, column=2).alignment = center
            ws_quote.cell(row=r, column=3, value=ch).font = data_font
            ws_quote.cell(row=r, column=3).alignment = center
            ws_quote.cell(row=r, column=4, value=wh).font = data_font
            ws_quote.cell(row=r, column=4).alignment = center
            ws_quote.cell(row=r, column=5, value=price).font = data_font
            ws_quote.cell(row=r, column=5).alignment = center
            ws_quote.cell(row=r, column=5).number_format = '0.0'
            # Calculated columns as formulas
            for c_idx, formula_tmpl in [(6, f'=E{r}*0.07/1.06'), (7, f'=E{r}*0.35'), (8, f'=E{r}*0.58')]:
                cell = ws_quote.cell(row=r, column=c_idx)
                cell.value = formula_tmpl
                cell.font = data_font
                cell.alignment = center
                cell.number_format = '0.0000'

    # ── Save ──
    # Apply template column fills to all data rows (after all writes/merges)
    from openpyxl.cell.cell import MergedCell
    for rn in range(4, 4 + n):
        for col_idx in template_fills:
            cell = ws.cell(row=rn, column=col_idx)
            if not isinstance(cell, MergedCell):
                try:
                    cell.fill = template_fills[col_idx]
                except:
                    pass

    wb.save(output_path)
    return True


def main():
    if len(sys.argv) < 4:
        print("用法: python3 gen_bill.py <订单列表.xlsx> <拣货数据.xlsx> <应收价格.xlsx> [输出文件名]")
        sys.exit(1)

    order_path = sys.argv[1]
    pick_path = sys.argv[2]
    price_path = sys.argv[3]

    print(f"📂 订单列表: {order_path}")
    print(f"📂 拣货数据: {pick_path}")
    print(f"📂 应收价格: {price_path}")

    # Load
    orders, picks, prices, price_rows_raw = load_data(order_path, pick_path, price_path)
    print(f"✅ 订单: {len(orders)} 条, 拣货: {len(picks)} 条, 价格: {len(prices)} 条")

    # Compute date range from orders
    date_serials = []
    for o in orders.values():
        d = o.get('发货日期', o.get('工作日期'))
        if d:
            date_serials.append(d)

    if date_serials:
        base = datetime(1899, 12, 30)
        min_dt = base + timedelta(days=min(date_serials))
        max_dt = base + timedelta(days=max(date_serials))
        # Find Monday of the week containing min_dt
        mon = min_dt - timedelta(days=min_dt.weekday())
        sun = mon + timedelta(days=6)
        year = mon.year
        date_range_str = f"{mon.month}.{mon.day}-{sun.month}.{sun.day}"
        title_str = f"至：广州拓锐科技有限公司（{mon.month}.{mon.day}-{sun.month}.{sun.day}）"
    else:
        year = datetime.now().year
        date_range_str = f"{datetime.now().month}.1-{datetime.now().month}.7"
        title_str = "至：广州拓锐科技有限公司"

    # Build rows
    rows = build_rows(orders, picks, prices)

    # Sort
    rows = sort_rows(rows)
    print(f"✅ 账单行: {len(rows)} 行")

    # Compute total accurately (same formula path as Excel)
    sum_O = sum(r['weight'] * r['unit_price'] * 0.07/1.06 for r in rows)
    sum_P = sum_O * 0.06
    sum_Q = sum(r['weight'] * r['unit_price'] * 0.35 for r in rows)
    sum_R = sum(r['weight'] * r['unit_price'] * 0.58 for r in rows)
    channels = set(r['service'] for r in rows)
    customs_count = len(channels)
    customs_S = customs_count * 350 / 1.06
    customs_T = customs_S * 0.06
    total = sum_O + sum_P + sum_Q + sum_R + customs_S + customs_T
    total_rounded = round(total, 1)

    # Auto-generate filename (dynamic month/year)
    file_month = date_range_str.split('.')[0] if date_range_str and '.' in date_range_str else f'{datetime.now().month}'
    output_path = sys.argv[4] if len(sys.argv) > 4 else \
        f'{file_month}月拓锐FBA仓-分段开票账单-JTT({date_range_str}) RMB {total_rounded}.xlsx'

    print(f"📄 输出: {output_path}")

    # Generate bill
    success = generate_bill(rows, output_path, title_str=title_str, date_range_str=date_range_str, price_rows_raw=price_rows_raw, year=year)

    if success:
        print(f"\n📊 费用汇总:")
        print(f"   渠道数: {customs_count} (报关费 {customs_count}×330.19 = {customs_S:.2f})")
        print(f"   国内运费: {sum_O:.2f} → 含税 {sum_O+sum_P:.2f}")
        print(f"   国际运费: {sum_Q+sum_R:.2f}")
        print(f"   报关费:   {customs_S:.2f} → 含税 {customs_S+customs_T:.2f}")
        print(f"   ─────────────────────────────")
        print(f"   ✅ 总计应收: {total_rounded} RMB")
        print(f"\n✅ 账单已生成: {output_path}")


if __name__ == '__main__':
    main()
