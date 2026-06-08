"""
拣货数据导出工具
规则：根据品名+客户下单数据(尺寸/重量)匹配历史数据，取系统历史拣货数据的值
用法：输入"导出拣货数据"即可运行

匹配逻辑：
1. 品名包含匹配（从主品名的逗号分隔列表中提取）
2. 客户维度（重量KG、长度CM、宽度CM、高度CM）精确匹配
3. 如有多个匹配，优先取扩展箱号前缀与FBA ID匹配的记录
4. 取历史拣货数据：高度(L)、宽度(K)、长度(J)、重量(I)
5. D(高度)=L, E(宽度)=K, F(长度)=J, G(重量)=I
6. H(材积重)=D*E*F/6000, I(收费重)=max(G, H)
"""
import openpyxl
import os
import sys

# === 配置 ===
DATA_DIR = os.path.dirname(os.path.abspath(__file__))
HISTORY_FILE = os.path.join(DATA_DIR, "历史数据（客户数据+拣货数据对比表）.xlsx")
MAX_FBA_PREFIX_LEN = 14


def load_history(filepath):
    """加载历史数据"""
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active

    records = []
    for r in range(3, ws.max_row + 1):
        fba_id = str(ws.cell(row=r, column=1).value or '').strip()
        prod_name = str(ws.cell(row=r, column=2).value or '').strip()
        cust = {
            'weight': ws.cell(row=r, column=4).value,   # D=实重
            'len': ws.cell(row=r, column=5).value,       # E=长
            'width': ws.cell(row=r, column=6).value,     # F=宽
            'height': ws.cell(row=r, column=7).value,    # G=高
        }
        pick = {
            'weight': ws.cell(row=r, column=9).value,    # I=实重
            'len': ws.cell(row=r, column=10).value,      # J=长
            'width': ws.cell(row=r, column=11).value,    # K=宽
            'height': ws.cell(row=r, column=12).value,   # L=高
        }

        if not prod_name or not all(v is not None for v in pick.values()):
            continue

        try:
            records.append({
                'fba_id': fba_id,
                'prod_name': prod_name,
                'cust': {k: float(v) if v else None for k, v in cust.items()},
                'pick': {k: float(v) for k, v in pick.items()},
            })
        except (ValueError, TypeError):
            continue

    return records


def extract_product_names(main_product_str):
    """从逗号分隔的主品名中提取单个产品名称"""
    if not main_product_str:
        return []
    return [p.strip() for p in main_product_str.split(',') if p.strip()]


def get_box_prefix(box_num):
    """从扩展箱号提取前缀"""
    return box_num.strip()[:MAX_FBA_PREFIX_LEN]


def name_matches(prod_name, product_names):
    """检查产品名是否匹配"""
    return any(
        pn and (pn in prod_name or prod_name.startswith(pn) or prod_name.find(pn) >= 0)
        for pn in product_names
    )


def calc_dims_score(rec_cust, test_cust):
    """计算维度匹配分数，返回 (维度是否匹配, 匹配分)"""
    ok, score = True, 0
    for key in ['weight', 'len', 'width', 'height']:
        rv, tv = rec_cust.get(key), test_cust.get(key)
        if tv is None or rv is None:
            ok = False
            continue
        diff = abs(rv - tv)
        threshold = 0.1 if key == 'weight' else 1.0
        if diff >= threshold:
            ok = False
        elif diff < (0.05 if key == 'weight' else 0.5):
            score += 1
    return ok, score


def find_best_match(records, product_names, box_prefix, test_cust):
    """多条件匹配：品名→尺寸→FBA前缀"""
    candidates = []
    for rec in records:
        if not name_matches(rec['prod_name'], product_names):
            continue
        dims_ok, dims_score = calc_dims_score(rec['cust'], test_cust)
        if not dims_ok:
            continue

        # FBA前缀匹配度
        fba_score = 0
        if rec['fba_id'] and box_prefix:
            if rec['fba_id'] in box_prefix or box_prefix.startswith(rec['fba_id']):
                fba_score = 2
            elif any(rec['fba_id'].startswith(box_prefix[:i]) for i in range(8, len(box_prefix))):
                fba_score = 1

        candidates.append((rec, fba_score, dims_score))

    if not candidates:
        return None

    # 排序：FBA匹配(高→低) → 尺寸分(高→低) → 保持原顺序(最早优先)
    candidates.sort(key=lambda c: (c[1], c[2]), reverse=True)
    return candidates[0][0]


def process_test_data(test_filepath, output_filepath, history_records):
    """处理测试数据"""
    wb = openpyxl.load_workbook(test_filepath)
    ws = wb.active

    matched, unmatched = 0, 0
    for r in range(2, ws.max_row + 1):
        box_num = str(ws.cell(row=r, column=3).value or '')
        if not box_num:
            continue

        product_names = extract_product_names(str(ws.cell(row=r, column=10).value or ''))
        box_prefix = get_box_prefix(box_num)
        test_cust = {
            'weight': ws.cell(row=r, column=13).value,  # M
            'len': ws.cell(row=r, column=14).value,      # N
            'width': ws.cell(row=r, column=15).value,    # O
            'height': ws.cell(row=r, column=16).value,   # P
        }

        match = find_best_match(history_records, product_names, box_prefix, test_cust)
        if not match:
            unmatched += 1
            continue

        p = match['pick']
        d_val, e_val, f_val = p['height'], p['width'], p['len']
        g_val = p['weight']
        h_val = round(d_val * e_val * f_val / 6000, 2)
        i_val = max(round(g_val, 2), h_val)

        ws.cell(row=r, column=4).value = int(d_val) if d_val == int(d_val) else round(d_val, 2)
        ws.cell(row=r, column=5).value = int(e_val) if e_val == int(e_val) else round(e_val, 2)
        ws.cell(row=r, column=6).value = int(f_val) if f_val == int(f_val) else round(f_val, 2)
        ws.cell(row=r, column=7).value = round(g_val, 2)
        ws.cell(row=r, column=8).value = h_val
        ws.cell(row=r, column=9).value = i_val
        matched += 1

    wb.save(output_filepath)
    print(f"✅ 处理完成: 匹配成功 {matched} 行, 失败 {unmatched} 行")
    print(f"📁 输出文件: {output_filepath}")
    return matched, unmatched


def main(input_filename=None, output_filename=None):
    """主入口"""
    if input_filename is None:
        # 自动查找最新的测试文件
        files = [f for f in os.listdir(DATA_DIR) if f.startswith("测试数据：导出拣货数据")]
        if not files:
            print("❌ 未找到测试数据文件")
            return
        input_filename = sorted(files)[-1]

    if output_filename is None:
        output_filename = input_filename.replace("测试数据：导出", "导出")

    test_path = os.path.join(DATA_DIR, input_filename)
    output_path = os.path.join(DATA_DIR, output_filename)

    if not os.path.exists(test_path):
        print(f"❌ 找不到文件: {test_path}")
        return

    print(f"📂 输入: {input_filename}")
    history = load_history(HISTORY_FILE)
    print(f"📚 历史数据: {len(history)} 条记录")
    process_test_data(test_path, output_path, history)


if __name__ == '__main__':
    if len(sys.argv) > 1:
        main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
    else:
        main()
