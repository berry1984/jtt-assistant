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
import uuid
from datetime import datetime, timedelta
from flask import Flask, request, render_template, send_file, send_from_directory, flash, redirect

# ── 模块路径 ──
THIS_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(THIS_DIR)  # bb plan1
INVOICE_DIR = os.path.join(PROJECT_DIR, '发票转换')

sys.path.insert(0, THIS_DIR)
sys.path.insert(0, INVOICE_DIR)

from gen_bill import load_data, build_rows, sort_rows, generate_bill
from convert_invoice import TRInvoice, convert_to_tiantu, convert_to_hangle

# ── 思锐(SR)账单生成模块（延迟导入，避免启动时依赖缺失） ──
SR_DIR = os.path.join(PROJECT_DIR, 'SR账单自动生成')
sys.path.insert(0, PROJECT_DIR)

# ── 拣货数据导出模块 ──
PICKING_DIR = os.path.join(PROJECT_DIR, '拣货数据')

# ── 投保区间拆分模块 ──
INSURANCE_DIR = os.path.join(PROJECT_DIR, '投保区间拆分发票')
sys.path.insert(0, INSURANCE_DIR)

def _get_sr_module():
    """延迟导入 gen_sr_bill，避免铁路部署时依赖问题导致整个app崩溃"""
    import importlib
    try:
        return importlib.import_module('gen_sr_bill')
    except ImportError as e:
        print(f"[SR] 导入gen_sr_bill失败: {e}")
        return None

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

@app.route('/debug_template', methods=['GET'])
def debug_template():
    """调试：查看模板文件信息和环境"""
    import os
    from openpyxl import load_workbook
    tpl_path = os.path.join(INVOICE_DIR, '天图单票专用模板20260601.xlsx')
    info = {'status': 'unknown', 'path': tpl_path, 'exists': os.path.exists(tpl_path)}
    # 检查Pillow
    try:
        import PIL
        info['pillow'] = PIL.__version__
    except ImportError:
        info['pillow'] = 'NOT INSTALLED'
    # git commit
    try:
        import subprocess
        sha = subprocess.run(['git', 'rev-parse', '--short', 'HEAD'], capture_output=True, text=True, cwd=os.path.dirname(THIS_DIR))
        info['commit'] = sha.stdout.strip() if sha.returncode == 0 else 'unknown'
    except:
        info['commit'] = 'error'
    if info['exists']:
        info['size'] = os.path.getsize(tpl_path)
        try:
            wb = load_workbook(tpl_path, data_only=True)
            info['sheet2_rows'] = wb['Sheet2'].max_row
            ws2 = wb['Sheet2']
            first3 = [ws2.cell(r, 1).value for r in range(1, 4)]
            last3 = [ws2.cell(r, 1).value for r in range(info['sheet2_rows']-2, info['sheet2_rows']+1)]
            info['first_3'] = first3
            info['last_3'] = last3
            wb.close()
        except Exception as e:
            info['error'] = str(e)
    return info


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

# 临时图片存储路径，用于 IMAGE() 公式引用
TEMP_IMAGE_DIR = os.path.join(app.config['UPLOAD_FOLDER'], 'temp_images')
os.makedirs(TEMP_IMAGE_DIR, exist_ok=True)

@app.route('/temp_images/<session_id>/<filename>')
def serve_temp_image(session_id, filename):
    """提供转换时提取的临时图片（IMAGE 公式引用）"""
    return send_from_directory(os.path.join(TEMP_IMAGE_DIR, session_id), filename)

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

        # ── 提取图片到临时目录（供 IMAGE() 公式以 HTTP URL 引用） ──
        image_session_id = uuid.uuid4().hex[:12]
        session_img_dir = os.path.join(TEMP_IMAGE_DIR, image_session_id)
        os.makedirs(session_img_dir, exist_ok=True)

        img_count = 0
        for src_row, img_bytes in tr.images.items():
            img_filename = f'image_tiantu_{img_count + 1}.png'
            with open(os.path.join(session_img_dir, img_filename), 'wb') as f:
                f.write(img_bytes)
            img_count += 1

        if img_count > 0:
            # 构建服务器 URL 前缀
            host_url = request.host_url.rstrip('/')
            image_url_base = f'{host_url}/temp_images/{image_session_id}'
        else:
            image_url_base = None

        base_name = os.path.splitext(invoice_file.filename)[0]

        # 文件名：航乐按 "客户名称 订单号 欧洲/英国发票.xlsx" 格式
        if target.startswith('航乐'):
            import re
            customer = ''
            m = re.search(r'（(.+?)发票）', base_name)
            if m:
                customer = m.group(1)
            order_no = tr.get('客户订单号', '') or ''
            region_label = '欧洲' if target == '航乐-eu' else '英国'
            output_name = f'{customer} {order_no} {region_label}发票.xlsx'.strip()
            if output_name.startswith(' '):
                output_name = output_name.lstrip()
        else:
            output_name = f'{base_name}-天图.xlsx'

        output_path = os.path.join(tmp_dir, output_name)

        if target == '天图':
            ok = convert_to_tiantu(tr, output_path, image_url_base=image_url_base)
        elif target == '航乐-uk':
            ok = convert_to_hangle(tr, output_path, region='uk', image_url_base=image_url_base)
        elif target == '航乐-eu':
            ok = convert_to_hangle(tr, output_path, region='eu', image_url_base=image_url_base)

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


# ═══════════════════════════════════════════════════════════
#  功能3：思锐(SR)账单自动生成  （使用订单列表渠道+可调汇率）
# ═══════════════════════════════════════════════════════════

@app.route('/generate_sr', methods=['POST'])
def generate_sr():
    sr = _get_sr_module()
    if not sr:
        flash('思锐账单模块加载失败，请联系管理员')
        return redirect('/sr')

    bill_file = request.files.get('sr_bill_file')
    order_file = request.files.get('sr_order_file')

    if not bill_file:
        flash('请上传系统账单文件')
        return redirect('/sr')

    tmp_dir = tempfile.mkdtemp(dir=app.config['UPLOAD_FOLDER'])
    try:
        # 保存上传文件
        bill_path = os.path.join(tmp_dir, 'system_bill.xls')
        bill_file.save(bill_path)

        order_list = {}
        order_path = None
        if order_file and order_file.filename:
            order_path = os.path.join(tmp_dir, 'order_list.xlsx')
            order_file.save(order_path)
            order_list = sr.read_order_list(order_path)

        # AB2 汇率（用户手动输入，默认0.1282）
        try:
            ab2_rate = float(request.form.get('ab2_rate', '0.1282'))
        except ValueError:
            ab2_rate = 0.1282

        # 读取系统账单
        waybills = sr.read_system_bill(bill_path)
        if not waybills:
            flash('系统账单中没有找到运单数据')
            return redirect('/sr')

        # 自动查找模板
        template_path = os.path.join(SR_DIR, '思锐账单模板模板 思锐开票账单-JTT（5.1-5.31）.xlsx')
        if not os.path.exists(template_path):
            for f in os.listdir(SR_DIR):
                if '思锐' in f and f.endswith('.xlsx') and '模板' in f:
                    template_path = os.path.join(SR_DIR, f)
                    break

        if not os.path.exists(template_path):
            flash('找不到思锐账单模板文件')
            return redirect('/sr')

        # 生成输出文件名
        from datetime import date
        today = date.today()
        output_name = f'思锐开票账单-JTT（{today.year}.{today.month:02d}.01-{today.month:02d}.{today.day:02d}）.xlsx'
        output_path = os.path.join(tmp_dir, output_name)

        # 生成账单（传入自定义AB2汇率）
        success = sr.generate_bill(waybills, template_path, output_path, order_list, ab2_rate=ab2_rate)

        if not success or not os.path.exists(output_path):
            flash('生成思锐账单失败')
            return redirect('/sr')

        return send_file(output_path,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         as_attachment=True,
                         download_name=output_name)

    except Exception as e:
        flash(f'处理出错: {str(e)}')
        return redirect('/sr')
    finally:
        pass


@app.route('/sr', methods=['GET'])
def sr_page():
    return render_template('index.html', targets=TARGET_OPTIONS, active_tab='sr')


# ═══════════════════════════════════════════════════════════
#  功能4：拣货数据导出
# ═══════════════════════════════════════════════════════════

@app.route('/picking', methods=['GET'])
def picking_page():
    return render_template('index.html', targets=TARGET_OPTIONS, active_tab='picking')


@app.route('/picking_export', methods=['POST'])
def picking_export():
    """上传测试拣货数据，导出填充好的拣货数据表"""
    import sys
    sys.path.insert(0, PICKING_DIR)

    from export_picking_data import load_history, find_best_match, extract_product_names, get_box_prefix, HISTORY_FILE

    test_file = request.files.get('picking_file')
    if not test_file or test_file.filename == '':
        flash('请上传拣货数据文件')
        return redirect('/picking')

    tmp_dir = tempfile.mkdtemp(dir=app.config['UPLOAD_FOLDER'])
    try:
        import openpyxl

        # 保存上传文件
        input_path = os.path.join(tmp_dir, 'input.xlsx')
        test_file.save(input_path)

        # 加载历史数据
        if not os.path.exists(HISTORY_FILE):
            flash(f'找不到历史数据文件，请确认 拣货数据/历史数据（客户数据+拣货数据对比表）.xlsx 存在')
            return redirect('/picking')

        history_records = load_history(HISTORY_FILE)

        # 处理
        wb = openpyxl.load_workbook(input_path)
        ws = wb.active

        matched, unmatched = 0, 0
        for r in range(2, ws.max_row + 1):
            box_num = str(ws.cell(row=r, column=3).value or '')
            if not box_num:
                continue

            product_names = extract_product_names(str(ws.cell(row=r, column=10).value or ''))
            box_prefix = get_box_prefix(box_num)
            test_cust = {
                'weight': ws.cell(row=r, column=13).value,
                'len': ws.cell(row=r, column=14).value,
                'width': ws.cell(row=r, column=15).value,
                'height': ws.cell(row=r, column=16).value,
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

        output_name = test_file.filename.replace('测试数据：导出', '导出')
        if output_name == test_file.filename:
            output_name = f'导出_{test_file.filename}'

        output_path = os.path.join(tmp_dir, output_name)
        wb.save(output_path)

        if matched == 0:
            flash('⚠️ 所有行均匹配失败，请检查数据格式')
            return redirect('/picking')

        return send_file(output_path,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         as_attachment=True,
                         download_name=output_name)

    except Exception as e:
        flash(f'处理出错: {str(e)}')
        return redirect('/picking')
    finally:
        pass


# ═══════════════════════════════════════════════════════════
#  功能5：投保区间拆分
# ═══════════════════════════════════════════════════════════

@app.route('/insurance', methods=['GET'])
def insurance_page():
    return render_template('index.html', targets=TARGET_OPTIONS, active_tab='insurance')


@app.route('/insurance_split', methods=['POST'])
def insurance_split():
    """
    上传下单发票 → 按每箱RMB拆分为 5 个区间文件 → 打包 ZIP 下载
    """
    invoice_file = request.files.get('insurance_file')
    if not invoice_file or invoice_file.filename == '':
        flash('请上传下单发票文件')
        return redirect('/insurance')

    tmp_dir = tempfile.mkdtemp(dir=app.config['UPLOAD_FOLDER'])
    try:
        invoice_path = os.path.join(tmp_dir, 'invoice.xlsx')
        invoice_file.save(invoice_path)

        # 调用拆分模块
        from split_insurance_v2 import split_invoice_to_ranges
        out_dir = os.path.join(tmp_dir, 'ranges')
        out_files = split_invoice_to_ranges(invoice_path, output_dir=out_dir)

        if not out_files:
            flash('拆分失败，未生成任何文件')
            return redirect('/insurance')

        # 打包为 ZIP
        import zipfile
        base_name = os.path.splitext(invoice_file.filename)[0]
        zip_name = f'{base_name}-投保拆分.zip'
        zip_path = os.path.join(tmp_dir, zip_name)

        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for range_name, file_path in out_files.items():
                arcname = os.path.basename(file_path)
                zf.write(file_path, arcname)

        return send_file(zip_path,
                         mimetype='application/zip',
                         as_attachment=True,
                         download_name=zip_name)

    except Exception as e:
        flash(f'拆分出错: {str(e)}')
        return redirect('/insurance')
    finally:
        def cleanup():
            try:
                shutil.rmtree(tmp_dir)
            except Exception:
                pass
        import atexit
        atexit.register(cleanup)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print("=" * 50)
    print("  JTT电商AI助手  — 一站式跨境物流工具")
    print("=" * 50)
    print(f"  📋 账单生成  → http://localhost:{port}")
    print(f"  📊 思锐账单  → http://localhost:{port}/sr")
    print(f"  📄 发票转换  → http://localhost:{port}/invoice")
    print(f"  📦 拣货导出  → http://localhost:{port}/picking")
    print(f"  🛡️ 投保拆分  → http://localhost:{port}/insurance")
    print("=" * 50)
    app.run(host='0.0.0.0', port=port)
