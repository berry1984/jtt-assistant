#!/usr/bin/env python3
"""
发票转换 - Web应用

将TR/思锐/赛诺吉发票(.xlsx)在线转换为供应商发票模板：
  - 天图
  - 航乐-UK
  - 航乐-EU

用法:
  python3 app.py [端口号]
  # 然后浏览器访问 http://localhost:5001
"""

import os
import sys
import tempfile
import shutil

from flask import Flask, request, render_template, send_file, flash, redirect

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from convert_invoice import TRInvoice, convert_to_tiantu, convert_to_hangle

app = Flask(__name__)
app.secret_key = 'invoice-converter-secret'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB
app.config['UPLOAD_FOLDER'] = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), 'uploads'
)

# 启动时清理旧上传
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
for f in os.listdir(app.config['UPLOAD_FOLDER']):
    p = os.path.join(app.config['UPLOAD_FOLDER'], f)
    try:
        if os.path.isdir(p):
            shutil.rmtree(p)
        else:
            os.remove(p)
    except Exception:
        pass


TARGET_OPTIONS = {
    '天图': '天图下单发票',
    '航乐-uk': '航乐-英国发票',
    '航乐-eu': '航乐-欧洲发票',
}


@app.route('/', methods=['GET'])
def index():
    return render_template('index.html', targets=TARGET_OPTIONS)


@app.route('/convert', methods=['POST'])
def convert():
    # 检查文件
    invoice_file = request.files.get('invoice_file')
    target = request.form.get('target', '天图')

    if not invoice_file or invoice_file.filename == '':
        flash('请上传 TR 发票文件')
        return redirect('/')

    if target not in TARGET_OPTIONS:
        flash('请选择有效的目标格式')
        return redirect('/')

    # 保存到临时目录
    tmp_dir = tempfile.mkdtemp(dir=app.config['UPLOAD_FOLDER'])
    try:
        invoice_path = os.path.join(tmp_dir, 'invoice.xlsx')
        invoice_file.save(invoice_path)

        # 读取发票
        tr = TRInvoice(invoice_path)

        # 生成输出文件名
        base_name = os.path.splitext(invoice_file.filename)[0]
        ext_map = {'天图': '天图', '航乐-uk': '航乐-UK', '航乐-eu': '航乐-EU'}
        output_name = f'{base_name}-{ext_map[target]}.xlsx'
        output_path = os.path.join(tmp_dir, output_name)

        # 执行转换
        if target == '天图':
            ok = convert_to_tiantu(tr, output_path)
        elif target == '航乐-uk':
            ok = convert_to_hangle(tr, output_path, region='uk')
        elif target == '航乐-eu':
            ok = convert_to_hangle(tr, output_path, region='eu')

        if not ok:
            flash('转换失败，请检查源文件格式')
            return redirect('/')

        # 返回文件
        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_name,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        flash(f'转换出错: {str(e)}')
        return redirect('/')
    finally:
        # 保留文件直到请求结束（send_file 后清理）
        def cleanup():
            try:
                shutil.rmtree(tmp_dir)
            except Exception:
                pass
        import atexit
        atexit.register(cleanup)


if __name__ == '__main__':
    print('🚀 发票转换 Web 服务已启动')
    print('📍 http://localhost:5000')
    print('📂 模板: 天图 | 航乐-UK | 航乐-EU')
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 5001
    print(f'🌐 访问地址: http://localhost:{port}')
    app.run(host='0.0.0.0', port=port, debug=True)
