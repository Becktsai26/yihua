import os
import webbrowser
import tempfile
from flask import Flask, request, render_template_string, send_file
from reconcile_core import load_invoice, load_trade, reconcile, get_summary_text, export_report

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <title>每日發票對帳工具</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: -apple-system, "Microsoft JhengHei", sans-serif; background: #f0f2f5; padding: 20px; }
        .container { max-width: 700px; margin: 0 auto; }
        h1 { text-align: center; color: #333; margin-bottom: 24px; font-size: 24px; }
        .card { background: #fff; border-radius: 12px; padding: 24px; margin-bottom: 16px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
        .card h2 { font-size: 16px; color: #666; margin-bottom: 16px; }
        label { display: block; font-weight: 600; margin-bottom: 6px; color: #333; }
        .hint { font-size: 12px; color: #999; margin-bottom: 12px; }
        input[type="file"] { width: 100%; padding: 10px; border: 2px dashed #ccc; border-radius: 8px; margin-bottom: 4px; background: #fafafa; cursor: pointer; }
        input[type="file"]:hover { border-color: #4472C4; }
        .btn { display: inline-block; padding: 12px 32px; border: none; border-radius: 8px; font-size: 16px; font-weight: 600; cursor: pointer; text-decoration: none; }
        .btn-primary { background: #4472C4; color: #fff; }
        .btn-primary:hover { background: #3561b3; }
        .btn-success { background: #28a745; color: #fff; }
        .btn-success:hover { background: #218838; }
        .btn-center { text-align: center; }
        .result { background: #f8f9fa; border-radius: 8px; padding: 16px; font-family: "Courier New", monospace; white-space: pre-line; line-height: 1.6; font-size: 14px; }
        .result.ok { border-left: 4px solid #28a745; }
        .result.warn { border-left: 4px solid #ffc107; }
        .actions { display: flex; gap: 12px; justify-content: center; margin-top: 16px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>每日發票對帳工具</h1>

        {% if not result %}
        <form method="POST" enctype="multipart/form-data">
            <div class="card">
                <h2>上傳檔案</h2>
                <label>發票檔案</label>
                <input type="file" name="invoice_file" accept=".xlsx,.xls,.csv" required>
                <div class="hint">支援 .xlsx / .csv</div>
                <label>交易檔案</label>
                <input type="file" name="trade_file" accept=".xlsx,.xls,.csv" required>
                <div class="hint">支援 .xlsx / .csv</div>
            </div>
            <div class="btn-center">
                <button type="submit" class="btn btn-primary">開始對帳</button>
            </div>
        </form>
        {% else %}
        <div class="card">
            <h2>對帳結果</h2>
            <div class="result {{ 'ok' if has_diff == False else 'warn' }}">{{ result }}</div>
            <div class="actions">
                <a href="/export" class="btn btn-success">匯出 Excel 報表</a>
                <a href="/" class="btn btn-primary">重新對帳</a>
            </div>
        </div>
        {% endif %}

        {% if error %}
        <div class="card">
            <div class="result warn">{{ error }}</div>
            <div class="actions">
                <a href="/" class="btn btn-primary">返回</a>
            </div>
        </div>
        {% endif %}
    </div>
</body>
</html>
'''

_last_result = {}


def _save_upload(file_obj, tmp_dir, prefix):
    """保存上傳檔案，保留原始副檔名"""
    original_name = file_obj.filename or ''
    ext = os.path.splitext(original_name)[1].lower() or '.xlsx'
    path = os.path.join(tmp_dir, f'{prefix}{ext}')
    file_obj.save(path)
    return path


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'GET':
        return render_template_string(HTML_TEMPLATE, result=None, error=None)

    invoice_file = request.files.get('invoice_file')
    trade_file = request.files.get('trade_file')
    if not invoice_file or not trade_file:
        return render_template_string(HTML_TEMPLATE, result=None, error='請上傳兩個檔案')

    try:
        tmp_dir = tempfile.mkdtemp()
        invoice_path = _save_upload(invoice_file, tmp_dir, 'invoice')
        trade_path = _save_upload(trade_file, tmp_dir, 'trade')

        invoice_df = load_invoice(invoice_path)
        trade_df = load_trade(trade_path)
        result = reconcile(invoice_df, trade_df)
        _last_result['data'] = result

        summary = get_summary_text(result)
        has_diff = (len(result['amount_mismatch']) + len(result['only_in_xlsx']) + len(result['only_in_csv'])) > 0

        return render_template_string(HTML_TEMPLATE, result=summary, has_diff=has_diff, error=None)
    except Exception as e:
        return render_template_string(HTML_TEMPLATE, result=None, error=f'對帳錯誤：{str(e)}')


@app.route('/export')
def export():
    if 'data' not in _last_result:
        return '請先執行對帳', 400

    tmp_path = os.path.join(tempfile.mkdtemp(), '對帳報表.xlsx')
    export_report(_last_result['data'], tmp_path)
    return send_file(tmp_path, as_attachment=True, download_name='對帳報表.xlsx')


if __name__ == '__main__':
    port = 5678
    print(f'對帳工具已啟動，請開啟瀏覽器：http://localhost:{port}')
    webbrowser.open(f'http://localhost:{port}')
    app.run(host='127.0.0.1', port=port, debug=False)
