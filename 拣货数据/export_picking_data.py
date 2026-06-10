"""
拣货数据导出工具 v2（2026-06-10 新规则）

输入：
  1. 发票 (.xlsx)  — 头部含服务/仓库代码/箱数，数据区含货箱编号+重量+尺寸+品名
  2. 系统导出拣货数据 (.xlsx) — 运单号+扩展箱号对照
  3. 箱规历史数据库 (.xlsx) — 品名+客户箱规 → 标准箱规
  4. 内部拣货数据参考值模版 (.xlsx) — 输出格式+公式

输出：
  内部拣货数据参考值 (.xlsx) — 按 SO 号归组，匹配历史箱规

匹配逻辑：
  - 发票每行货箱编号 → 前12位 = FBA ID → 匹配系统导出的扩展箱号 → 取运单号(SO号)
  - 箱规历史：品名+重量/尺寸匹配 → 取标准箱规(V/W/X/Y)
  - 无历史匹配 → V/W/X/Y 留空，标红

用法：
  python3 export_picking_data.py
  python3 export_picking_data.py <发票文件> <系统导出文件> [输出文件]
"""
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import os
import re
import sys

# ── 配置 ──
DATA_DIR = os.path.dirname(os.path.abspath(__file__))
HISTORY_FILE = os.path.join(DATA_DIR, "箱规历史数据库.xlsx")
TEMPLATE_FILE = os.path.join(DATA_DIR, "内部拣货数据参考值模版.xlsx")

DEFAULT_COUNTRY_FILL = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

# ── 辅助函数 ──

def extract_fba_id(box_no):
    """从货箱编号提取前12位 FBA ID"""
    box_no = str(box_no).strip()
    return box_no[:12] if len(box_no) >= 12 else box_no


def calc_box_count(box_no):
    """计算箱数：从货箱编号 U{start}-{end} 或 U{num} 格式"""
    box_no = str(box_no).strip()
    # 匹配 U 后面跟数字，可选 -{数字}
    m = re.search(r'U(\d+)(?:-(\d+))?$', box_no)
    if m:
        start = int(m.group(1))
        end = int(m.group(2)) if m.group(2) else start
        return end - start + 1
    return 1


def country_from_warehouse(warehouse):
    """根据仓库代码推断国家（仅供参考，最终手动确认）"""
    w = str(warehouse or '').upper().strip()
    us_warehouses = {'RFD2', 'IAH3', 'LGB8', 'LAX9', 'SMF3', 'ONT8', 'FTW1', 'MEM6', 'MDW2'}
    if w in us_warehouses or w.startswith(('RFD', 'IAH', 'LGB', 'LAX', 'SMF', 'ONT', 'FTW', 'MEM', 'MDW')):
        return 'US'
    # 欧洲常见仓库前缀
    eu_prefixes = ('DTM', 'WRO', 'HAJ', 'FBA', 'POZ', 'AVP', 'YVR')
    if any(w.startswith(p) for p in eu_prefixes):
        return 'DE'  # 默认欧洲-德国
    return ''


# ── 解析函数 ──

def parse_invoice(filepath):
    """解析发票文件，返回 (data_rows, service, warehouse)"""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    sn = wb.sheetnames
    # 查找主 Sheet
    main_sheet = None
    for name in sn:
        if '发票' in name or 'page' in name.lower():
            main_sheet = name
            break
    if not main_sheet:
        main_sheet = sn[0]

    ws = wb[main_sheet]

    # ── 提取头部信息 ──
    service = ''
    warehouse = ''
    for r in range(1, 25):
        a = str(ws.cell(row=r, column=1).value or '').strip()
        b = ws.cell(row=r, column=2).value
        if a == '服务':
            service = str(b or '').strip()
        elif a == '收件人姓名':
            warehouse = str(b or '').strip()

    # ── 查找数据表头行 ──
    data_start = None
    for r in range(1, ws.max_row + 1):
        a = str(ws.cell(row=r, column=1).value or '').strip()
        if '货箱编号' in a:
            data_start = r
            break

    if data_start is None:
        wb.close()
        return [], service, warehouse

    # ── 读取数据行 ──
    data_rows = []
    for r in range(data_start + 1, ws.max_row + 1):
        box_no = str(ws.cell(row=r, column=1).value or '').strip()
        if not box_no:
            break
        row = {
            'box_no': box_no,
            'weight': ws.cell(row=r, column=2).value,    # B
            'length': ws.cell(row=r, column=3).value,     # C
            'width': ws.cell(row=r, column=4).value,       # D
            'height': ws.cell(row=r, column=5).value,      # E
            'en_name': ws.cell(row=r, column=6).value,     # F
            'cn_name': ws.cell(row=r, column=7).value,     # G
        }
        data_rows.append(row)

    wb.close()
    return data_rows, service, warehouse


def parse_system_export(filepath):
    """解析系统导出拣货数据，返回 {FBA前缀 → SO号} 映射"""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    prefix_to_so = {}
    for r in range(2, ws.max_row + 1):
        ext_box = str(ws.cell(row=r, column=3).value or '').strip()
        so_no = str(ws.cell(row=r, column=2).value or '').strip()
        if not ext_box or not so_no:
            continue
        prefix = ext_box[:12]
        if prefix not in prefix_to_so:
            prefix_to_so[prefix] = so_no
    wb.close()
    return prefix_to_so


def parse_history(filepath):
    """解析箱规历史数据库，返回记录列表 [{name, w, l, wid, h, rw, rl, rwid, rh}]"""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    records = []
    for r in range(3, ws.max_row + 1):
        name = str(ws.cell(row=r, column=1).value or '').strip()
        if not name:
            continue
        try:
            rec = {
                'name': name,
                'weight': _to_float(ws.cell(row=r, column=2).value),
                'length': _to_float(ws.cell(row=r, column=3).value),
                'width': _to_float(ws.cell(row=r, column=4).value),
                'height': _to_float(ws.cell(row=r, column=5).value),
                'rw': _to_float(ws.cell(row=r, column=6).value),
                'rl': _to_float(ws.cell(row=r, column=7).value),
                'rwid': _to_float(ws.cell(row=r, column=8).value),
                'rh': _to_float(ws.cell(row=r, column=9).value),
            }
        except (ValueError, TypeError):
            continue
        if all(v is not None for v in [rec['rw'], rec['rl'], rec['rwid'], rec['rh']]):
            records.append(rec)
    wb.close()
    return records


def _to_float(v):
    """安全转 float"""
    if v is None:
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def find_history_match(records, name, weight, length, width, height):
    """在箱规历史数据库中查找匹配记录，返回 {rw, rl, rwid, rh} 或 None

    支持长宽高 6 种排列组合（实际数据中尺寸顺序常不一致）。
    """
    if not name:
        return None
    name = name.strip()

    # 生成所有 (l, w, h) 排列
    dims_permutations = [
        (length, width, height),
        (length, height, width),
        (width, length, height),
        (width, height, length),
        (height, length, width),
        (height, width, length),
    ]

    candidates = []
    for rec in records:
        if rec['name'] != name:
            continue
        if not _dims_match(rec['weight'], weight, 0.5):
            continue
        # 尝试各种排列
        best_score = -1
        for pl, pw, ph in dims_permutations:
            l_ok = _dims_match(rec['length'], pl, 1.0)
            w_ok = _dims_match(rec['width'], pw, 1.0)
            h_ok = _dims_match(rec['height'], ph, 1.0)
            if l_ok and w_ok and h_ok:
                score = 0
                if _dims_match(rec['length'], pl, 0.5): score += 1
                if _dims_match(rec['width'], pw, 0.5): score += 1
                if _dims_match(rec['height'], ph, 0.5): score += 1
                if score > best_score:
                    best_score = score
        if best_score >= 0:
            # weight 精确度加分
            if _dims_match(rec['weight'], weight, 0.1):
                best_score += 1
            candidates.append((best_score, rec))

    if not candidates:
        return None
    candidates.sort(key=lambda x: -x[0])
    best = candidates[0][1]
    return {'rw': best['rw'], 'rl': best['rl'], 'rwid': best['rwid'], 'rh': best['rh']}


def _dims_match(a, b, tolerance):
    """检查两个数值是否在允许误差内（任一为 None 则跳过该维度）"""
    if a is None or b is None:
        return True  # 无法比较时视为匹配
    return abs(a - b) < tolerance


# ── 输出生成 ──

def generate_picking_output(invoice_file, system_file, output_path,
                              history_file=None, template_file=None):
    """核心入口：生成内部拣货数据参考值"""
    if history_file is None:
        history_file = HISTORY_FILE
    if template_file is None:
        template_file = TEMPLATE_FILE

    # ── 解析输入 ──
    data_rows, service, warehouse = parse_invoice(invoice_file)
    prefix_to_so = parse_system_export(system_file)
    history_records = parse_history(history_file)

    if not data_rows:
        raise ValueError("发票中未找到有效数据行")

    print(f"  📄 发票数据: {len(data_rows)} 行")
    print(f"  🔗 系统SO映射: {len(prefix_to_so)} 个FBA前缀")
    print(f"  📚 箱规历史: {len(history_records)} 条记录")

    # ── 构建输出行数据 ──
    output_rows = []
    missing_history = []  # 记录无历史匹配的行
    for row in data_rows:
        box_no = row['box_no']
        fba_id = extract_fba_id(box_no)
        box_count = calc_box_count(box_no)
        so_no = prefix_to_so.get(fba_id, '')

        cn_name = str(row['cn_name'] or '').strip()
        weight = _to_float(row['weight'])
        length = _to_float(row['length'])
        width = _to_float(row['width'])
        height = _to_float(row['height'])

        hm = find_history_match(history_records, cn_name, weight, length, width, height)

        out = {
            'so_no': so_no,
            'service': service,
            'country': country_from_warehouse(warehouse),
            'warehouse': warehouse,
            'e_price': '',    # 手动输入
            'f_price': '',    # 手动输入
            'supplier_ch': '', # 手动输入
            'ship_name': '',  # 手动输入
            'customs': '',    # 手动输入
            'fba_id': fba_id,
            'cn_name': cn_name,
            'pickup_fee': '', # 手动输入
            'box_count': box_count,
            'weight': weight,
            'length': length,
            'width': width,
            'height': height,
            'history_note': '', # 手动输入
            'ref_w': hm['rw'] if hm else None,
            'ref_l': hm['rl'] if hm else None,
            'ref_wid': hm['rwid'] if hm else None,
            'ref_h': hm['rh'] if hm else None,
        }
        output_rows.append(out)
        if hm is None:
            missing_history.append(out)

    total_boxes = sum(r['box_count'] for r in output_rows)
    print(f"  📊 输出行: {len(output_rows)} 行, 总箱数: {total_boxes}")
    print(f"  ⚠️  无历史匹配: {len(missing_history)} 行")

    # ── 写入模板 ──
    wb = openpyxl.load_workbook(template_file)
    ws = wb.active

    # 先解除所有合并单元格（Row 2+），再删除空行
    merges_to_remove = [str(m) for m in list(ws.merged_cells.ranges) if m.min_row >= 2]
    for m_str in merges_to_remove:
        ws.unmerge_cells(m_str)

    # 删除旧数据行 (Row 2 起)，保留表头 Row 1
    max_existing = ws.max_row
    if max_existing >= 2:
        ws.delete_rows(2, max_existing - 1)

    # ── 按 SO 号分组 ──
    groups = []
    cur = []
    for row in output_rows:
        if not cur:
            cur = [row]
        elif (cur[0]['so_no'] == row['so_no']
              and cur[0]['service'] == row['service']
              and cur[0]['warehouse'] == row['warehouse']):
            cur.append(row)
        else:
            groups.append(cur)
            cur = [row]
    if cur:
        groups.append(cur)

    # ── 写入数据 ──
    current_row = 2
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    for group in groups:
        start_row = current_row
        # 写共享值（A-I列，仅首行）
        ws.cell(row=start_row, column=1).value = group[0]['so_no']   # A
        ws.cell(row=start_row, column=2).value = group[0]['service']  # B
        ws.cell(row=start_row, column=3).value = group[0]['country']  # C
        ws.cell(row=start_row, column=4).value = group[0]['warehouse']  # D
        # E-I 手动输入，留空

        for row_data in group:
            r = current_row
            # J, K
            ws.cell(row=r, column=10).value = row_data['fba_id']   # J
            ws.cell(row=r, column=11).value = row_data['cn_name']  # K

            # L 手动输入

            # M
            ws.cell(row=r, column=13).value = row_data['box_count']  # M

            # N-Q
            ws.cell(row=r, column=14).value = row_data['weight']  # N
            ws.cell(row=r, column=15).value = row_data['length']  # O
            ws.cell(row=r, column=16).value = row_data['width']   # P
            ws.cell(row=r, column=17).value = row_data['height']  # Q

            # R = O*P*Q/6000  (formula)
            ws.cell(row=r, column=18).value = f'=O{r}*P{r}*Q{r}/6000'

            # S, T (formulas)
            ws.cell(row=r, column=19).value = f'=R{r}-Z{r}'
            ws.cell(row=r, column=20).value = f'=O{r}+P{r}+Q{r}-Y{r}-X{r}-W{r}'

            # U 手动输入

            # V-Y (历史匹配值)
            v_val = row_data['ref_w']
            w_val = row_data['ref_l']
            x_val = row_data['ref_wid']
            y_val = row_data['ref_h']

            ws.cell(row=r, column=22).value = v_val  # V
            ws.cell(row=r, column=23).value = w_val  # W
            ws.cell(row=r, column=24).value = x_val  # X
            ws.cell(row=r, column=25).value = y_val  # Y

            # 无匹配时标红
            if all(v is None for v in [v_val, w_val, x_val, y_val]):
                for col in [22, 23, 24, 25]:
                    ws.cell(row=r, column=col).fill = red_fill
            elif v_val is None:
                ws.cell(row=r, column=22).fill = red_fill
            elif w_val is None:
                ws.cell(row=r, column=23).fill = red_fill
            elif x_val is None:
                ws.cell(row=r, column=24).fill = red_fill
            elif y_val is None:
                ws.cell(row=r, column=25).fill = red_fill

            # Z ~ AD (formulas)
            ws.cell(row=r, column=26).value = f'=W{r}*X{r}*Y{r}/6000'
            ws.cell(row=r, column=27).value = f'=W{r}*X{r}*Y{r}*M{r}/1000000'
            ws.cell(row=r, column=28).value = f'=V{r}*M{r}'
            ws.cell(row=r, column=29).value = f'=Z{r}*M{r}'
            ws.cell(row=r, column=30).value = f'=ROUND(MAX(AB{r}:AC{r}),0)'

            # AE-AJ 手动输入，留空

            current_row += 1

        end_row = current_row - 1

        # 同一 SO 号组：合并 A, B 列
        if start_row < end_row:
            ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
            ws.merge_cells(start_row=start_row, start_column=2, end_row=end_row, end_column=2)
            # 也可合并 C, D 等，但为安全只合并 A, B

    # ── 保存 ──
    wb.save(output_path)
    print(f"  ✅ 已保存: {output_path}")
    return output_path, total_boxes


# ── CLI 入口 ──

def main(invoice_file=None, system_file=None, output_file=None):
    """CLI 主入口"""
    if invoice_file is None:
        files = sorted([
            os.path.join(DATA_DIR, f) for f in os.listdir(DATA_DIR)
            if f.startswith("(测试用发票)") and f.endswith('.xlsx')
        ])
        if not files:
            print("❌ 未找到测试发票文件")
            return None, None, None
        invoice_file = files[-1]

    if system_file is None:
        files = sorted([
            os.path.join(DATA_DIR, f) for f in os.listdir(DATA_DIR)
            if f.startswith("(测试用)导出拣货数据") and f.endswith('.xlsx')
        ])
        if not files:
            print("❌ 未找到系统导出拣货数据文件")
            return None, None, None
        system_file = files[-1]

    # 先处理，获取总箱数（用临时文件名）
    import re
    date_str = ''
    base = os.path.basename(system_file)
    m = re.search(r'(\d{4}-\d{2}-\d{2})', base)
    if m:
        date_str = m.group(1)

    tmp_path = os.path.join(DATA_DIR, '_temp_output.xlsx')
    result_path, total_boxes = generate_picking_output(invoice_file, system_file, tmp_path)

    # 生成带日期+箱数的文件名并重命名
    if output_file is None:
        output_name = f"内部拣货数据参考值_{date_str}_{total_boxes}箱.xlsx"
        output_file = os.path.join(DATA_DIR, output_name)

    if os.path.exists(result_path):
        os.rename(result_path, output_file)
        print(f"  ✅ 最终文件: {output_file}")

    return invoice_file, system_file, output_file

    print(f"📂 发票: {os.path.basename(invoice_file)}")
    print(f"📂 系统导出拣货数据: {os.path.basename(system_file)}")
    print(f"📂 输出: {os.path.basename(output_file)}")
    print()

    result = generate_picking_output(invoice_file, system_file, output_file)
    return invoice_file, system_file, result


if __name__ == '__main__':
    if len(sys.argv) > 1:
        if len(sys.argv) >= 4:
            main(sys.argv[1], sys.argv[2], sys.argv[3])
        elif len(sys.argv) >= 3:
            main(sys.argv[1], sys.argv[2])
        else:
            main(sys.argv[1])
    else:
        main()
