#!/usr/bin/env python3
"""
JTT电商AI助手 - Web应用

功能：
  1. TR账单自动生成 - 上传订单列表+拣货数据+应收价格，生成标准格式账单
  2. TR发票转换 - 上传TR发票，转换为天图/航乐等供应商模板

用法：
  python3 app.py
  # 浏览器访问 http://localhost:5000
"""

import os
import sys
import tempfile
import shutil
import atexit
from datetime import datetime, timedelta
from flask import Flask, request, render_template, send_file, flash, redirect

# ── 模块路径 ──
THIS_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(THIS_DIR)  # bb plan1
INVOICE_DIR = os.path.join(PROJECT_DIR, '发票转换')

sys.path.insert(0, THIS_DIR)
sys.path.insert(0, INVOICE_DIR)

from gen_bill import load_data, build_rows, sort_rows, generate_bill
from convert_invoice import TRInvoice, convert_to_tiantu, convert_to_hangle

app = Flask(__name__)
app.secret_key = 'jtt-ai-assistant-secret'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB
app.config['UPLOAD_FOLDER'] = os.path.join(THIS_DIR, 'uploads')

# 启动时清理旧上传
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
for f in os.listdir(app.config['UPLOAD_FOLDER']):
    p = os.path.join(app.config['UPLOAD_FOLDER'], f)
    try:
        if os.path.isdir(p):
            shutil.rmtree(p)
        else:
            os.remove(p)
    except:
        pass

# ── 发票转换目标格式 ──
TARGET_OPTIONS = {
    '天图': '天图下单发票',
    '航乐-uk': '航乐-英国发票',
    '航乐-eu': '航乐-欧洲发票',
}


# ═══════════════════════════════════════════════════════════
#  首页
# ═══════════════════════════════════════════════════════════

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html', targets=TARGET_OPTIONS, active_tab='bill')


# ═══════════════════════════════════════════════════════════
#  页脚
# ═══════════════════════════════════════════════════════════

@app.route('/invoice', methods=['GET'])
def invoice_page():
    return render_template('index.html', targets=TARGET_OPTIONS, active_tab='invoice')


# ═══════════════════════════════════════════════════════════
#  功能1：TR账单自动生成
# ═══════════════════════════════════════════════════════════

@app.route('/generate', methods=['POST'])
def generate():
    order_file = request.files.get('order_file')
    pick_file = request.files.get('pick_file')
    price_file = request.files.get('price_file')

    if not all([order_file, pick_file, price_file]):
        flash('请上传三个文件：订单列表、拣货数据、应收价格')
        return redirect('/')

    tmp_dir = tempfile.mkdtemp(dir=app.config['UPLOAD_FOLDER'])
    try:
        order_path = os.path.join(tmp_dir, 'order.xlsx')
        pick_path = os.path.join(tmp_dir, 'pick.xlsx')
        price_path = os.path.join(tmp_dir, 'price.xlsx')
        order_file.save(order_path)
        pick_file.save(pick_path)
        price_file.save(price_path)

        orders, picks, prices, price_rows_raw = load_data(order_path, pick_path, price_path)
        rows = build_rows(orders, picks, prices)
        rows = sort_rows(rows)

        # 日期范围
        date_serials = []
        for o in orders.values():
            d = o.get('发货日期', o.get('工作日期'))
            if d:
                date_serials.append(d)

        if date_serials:
            base = datetime(1899, 12, 30)
            min_dt = base + timedelta(days=min(date_serials))
            max_dt = base + timedelta(days=max(date_serials))
            mon = min_dt - timedelta(days=min_dt.weekday())
            sun = mon + timedelta(days=6)
            year = mon.year
            date_range_str = f"{mon.month}.{mon.day}-{sun.month}.{sun.day}"
            title_str = f"至：广州拓锐科技有限公司（{mon.month}.{mon.day}-{sun.month}.{sun.day}）"
        else:
            year = datetime.now().year
            date_range_str = f"{datetime.now().month}.1-{datetime.now().month}.7"
            title_str = "至：广州拓锐科技有限公司"

        def safe_float(v):
            try: return float(v) if v is not None and v != '' else 0
            except: return 0
        sum_O = sum(safe_float(r.get('weight', 0)) * safe_float(r.get('unit_price', 0)) * 0.07/1.06 for r in rows)
        sum_P = sum_O * 0.06
        sum_Q = sum(safe_float(r.get('weight', 0)) * safe_float(r.get('unit_price', 0)) * 0.35 for r in rows)
        sum_R = sum(safe_float(r.get('weight', 0)) * safe_float(r.get('unit_price', 0)) * 0.58 for r in rows)
        channels = set(r['service'] for r in rows)
        customs_S = len(channels) * 350 / 1.06
        customs_T = customs_S * 0.06
        total = round(sum_O + sum_P + sum_Q + sum_R + customs_S + customs_T, 1)

        file_month = date_range_str.split('.')[0] if '.' in date_range_str else f'{datetime.now().month}'
        output_name = f'{file_month}月拓锐FBA仓-分段开票账单-JTT({date_range_str}) RMB {total}.xlsx'
        output_path = os.path.join(tmp_dir, output_name)

        success = generate_bill(rows, output_path, title_str=title_str,
                                date_range_str=date_range_str, price_rows_raw=price_rows_raw, year=year)

        if not success or not os.path.exists(output_path):
            flash('生成账单失败，请检查文件内容')
            return redirect('/')

        return send_file(output_path,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         as_attachment=True,
                         download_name=output_name)

    except Exception as e:
        flash(f'处理出错: {str(e)}')
        return redirect('/')
    finally:
        pass


# ═══════════════════════════════════════════════════════════
#  功能2：TR发票 → 供应商模板 转换
# ═══════════════════════════════════════════════════════════

@app.route('/invoice_convert', methods=['POST'])
def invoice_convert():
    invoice_file = request.files.get('invoice_file')
    target = request.form.get('target', '天图')

    if not invoice_file or invoice_file.filename == '':
        flash('请上传 TR 发票文件')
        return redirect('/invoice')

    if target not in TARGET_OPTIONS:
        flash('请选择有效的目标格式')
        return redirect('/invoice')

    tmp_dir = tempfile.mkdtemp(dir=app.config['UPLOAD_FOLDER'])
    try:
        invoice_path = os.path.join(tmp_dir, 'invoice.xlsx')
        invoice_file.save(invoice_path)

        tr = TRInvoice(invoice_path)

        base_name = os.path.splitext(invoice_file.filename)[0]
        ext_map = {'天图': '天图', '航乐-uk': '航乐-UK', '航乐-eu': '航乐-EU'}
        output_name = f'{base_name}-{ext_map[target]}.xlsx'
        output_path = os.path.join(tmp_dir, output_name)

        if target == '天图':
            ok = convert_to_tiantu(tr, output_path)
        elif target == '航乐-uk':
            ok = convert_to_hangle(tr, output_path, region='uk')
        elif target == '航乐-eu':
            ok = convert_to_hangle(tr, output_path, region='eu')

        if not ok:
            flash('转换失败，请检查源文件格式')
            return redirect('/invoice')

        return send_file(output_path,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         as_attachment=True,
                         download_name=output_name)

    except Exception as e:
        flash(f'转换出错: {str(e)}')
        return redirect('/invoice')
    finally:
        def cleanup():
            try: shutil.rmtree(tmp_dir)
            except: pass
        atexit.register(cleanup)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print("=" * 50)
    print("  JTT电商AI助手  — 一站式跨境物流工具")
    print("=" * 50)
    print(f"  📋 账单生成  → http://localhost:{port}")
    print(f"  📄 发票转换  → http://localhost:{port}/invoice")
    print("=" * 50)
    app.run(host='0.0.0.0', port=port)
