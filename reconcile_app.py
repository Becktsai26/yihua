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
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>每日發票對帳工具</title>
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500;600;700&family=Inter:wght@300;400;500;600&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg-start: #0c0e1a;
            --bg-end: #1a1040;
            --glass: rgba(255, 255, 255, 0.06);
            --glass-border: rgba(255, 255, 255, 0.12);
            --glass-hover: rgba(255, 255, 255, 0.10);
            --accent: #7c6aef;
            --accent-light: #a78bfa;
            --accent-glow: rgba(124, 106, 239, 0.35);
            --green: #34d399;
            --green-glow: rgba(52, 211, 153, 0.2);
            --red: #f87171;
            --red-glow: rgba(248, 113, 113, 0.2);
            --text: #f0f0f8;
            --text-dim: rgba(240, 240, 248, 0.55);
            --text-mid: rgba(240, 240, 248, 0.75);
        }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: "Noto Sans TC", "Inter", -apple-system, sans-serif;
            min-height: 100vh;
            background: linear-gradient(135deg, var(--bg-start) 0%, var(--bg-end) 50%, #0f1828 100%);
            color: var(--text);
            overflow-x: hidden;
        }

        /* 背景光暈裝飾 */
        body::before, body::after {
            content: '';
            position: fixed;
            border-radius: 50%;
            filter: blur(120px);
            z-index: 0;
            pointer-events: none;
        }
        body::before {
            width: 600px; height: 600px;
            background: radial-gradient(circle, rgba(124, 106, 239, 0.15), transparent 70%);
            top: -200px; left: -100px;
        }
        body::after {
            width: 500px; height: 500px;
            background: radial-gradient(circle, rgba(52, 211, 153, 0.08), transparent 70%);
            bottom: -150px; right: -100px;
        }

        .container {
            max-width: 640px;
            margin: 0 auto;
            padding: 48px 20px 36px;
            position: relative;
            z-index: 1;
        }

        /* 標題區 */
        .header {
            text-align: center;
            margin-bottom: 40px;
        }
        .header h1 {
            font-size: 28px;
            font-weight: 700;
            letter-spacing: 2px;
            background: linear-gradient(135deg, var(--text), var(--accent-light));
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
            margin-bottom: 8px;
        }
        .header .subtitle {
            font-size: 13px;
            color: var(--text-dim);
            font-weight: 300;
            letter-spacing: 3px;
            text-transform: uppercase;
        }

        /* 玻璃卡片 */
        .card {
            background: var(--glass);
            backdrop-filter: blur(24px);
            -webkit-backdrop-filter: blur(24px);
            border: 1px solid var(--glass-border);
            border-radius: 20px;
            padding: 32px;
            margin-bottom: 24px;
            position: relative;
            overflow: hidden;
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3), inset 0 1px 0 rgba(255, 255, 255, 0.08);
        }
        .card::before {
            content: '';
            position: absolute;
            top: 0; left: 0; right: 0;
            height: 1px;
            background: linear-gradient(90deg, transparent, rgba(255,255,255,0.15), transparent);
        }
        .card h2 {
            font-size: 14px;
            font-weight: 500;
            color: var(--text-mid);
            margin-bottom: 24px;
            letter-spacing: 1px;
            text-transform: uppercase;
        }

        /* 檔案上傳 */
        label {
            display: block;
            font-weight: 500;
            margin-bottom: 8px;
            color: var(--text);
            font-size: 14px;
        }
        .upload-zone {
            width: 100%;
            padding: 28px 20px;
            border: 1.5px dashed rgba(255, 255, 255, 0.15);
            border-radius: 14px;
            margin-bottom: 6px;
            background: rgba(255, 255, 255, 0.02);
            cursor: pointer;
            transition: all 0.3s ease;
            text-align: center;
            position: relative;
        }
        .upload-zone:hover, .upload-zone.drag-over {
            border-color: var(--accent-light);
            background: rgba(124, 106, 239, 0.08);
            box-shadow: 0 0 24px rgba(124, 106, 239, 0.1);
        }
        .upload-zone .icon {
            font-size: 28px;
            margin-bottom: 6px;
            opacity: 0.7;
        }
        .upload-zone .text {
            color: var(--text-dim);
            font-size: 13px;
            font-weight: 300;
        }
        .upload-zone .filename {
            color: var(--accent-light);
            font-size: 14px;
            font-weight: 600;
            display: none;
        }
        .upload-zone input[type="file"] {
            position: absolute; top: 0; left: 0; width: 100%; height: 100%;
            opacity: 0; cursor: pointer;
        }
        .hint {
            font-size: 11px;
            color: var(--text-dim);
            margin-bottom: 20px;
        }

        /* 按鈕 */
        .btn {
            display: inline-block;
            padding: 13px 36px;
            border: none;
            border-radius: 12px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            text-decoration: none;
            transition: all 0.3s ease;
            letter-spacing: 0.5px;
        }
        .btn-primary {
            background: linear-gradient(135deg, var(--accent), var(--accent-light));
            color: #fff;
            box-shadow: 0 4px 20px var(--accent-glow);
        }
        .btn-primary:hover {
            box-shadow: 0 8px 32px var(--accent-glow);
            transform: translateY(-2px);
        }
        .btn-export {
            background: rgba(52, 211, 153, 0.15);
            border: 1px solid rgba(52, 211, 153, 0.3);
            color: var(--green);
            backdrop-filter: blur(12px);
        }
        .btn-export:hover {
            background: rgba(52, 211, 153, 0.25);
            box-shadow: 0 4px 20px var(--green-glow);
            transform: translateY(-2px);
        }
        .btn-outline {
            background: rgba(255, 255, 255, 0.05);
            color: var(--text-dim);
            border: 1px solid var(--glass-border);
            backdrop-filter: blur(12px);
        }
        .btn-outline:hover {
            border-color: rgba(255, 255, 255, 0.25);
            color: var(--text);
            background: rgba(255, 255, 255, 0.08);
        }
        .btn-center { text-align: center; }
        .actions {
            display: flex;
            gap: 14px;
            justify-content: center;
            margin-top: 24px;
            flex-wrap: wrap;
        }

        /* 結果文字 */
        .result-box {
            background: rgba(0, 0, 0, 0.25);
            border: 1px solid rgba(255, 255, 255, 0.06);
            border-radius: 14px;
            padding: 20px;
            font-family: "SF Mono", "Fira Code", "Courier New", monospace;
            font-size: 13px;
            line-height: 1.9;
            white-space: pre-line;
            color: var(--text-mid);
        }
        .result-box.ok {
            border-color: rgba(52, 211, 153, 0.25);
            box-shadow: inset 0 0 30px rgba(52, 211, 153, 0.03);
        }
        .result-box.warn {
            border-color: rgba(248, 113, 113, 0.25);
            box-shadow: inset 0 0 30px rgba(248, 113, 113, 0.03);
        }
        .result-box.error {
            border-color: rgba(248, 113, 113, 0.25);
            box-shadow: inset 0 0 30px rgba(248, 113, 113, 0.03);
        }

        /* 狀態標籤 */
        .badge {
            display: inline-block;
            font-size: 11px;
            font-weight: 600;
            padding: 4px 14px;
            border-radius: 20px;
            margin-bottom: 16px;
            letter-spacing: 1.5px;
        }
        .badge-ok {
            background: var(--green-glow);
            color: var(--green);
            border: 1px solid rgba(52, 211, 153, 0.2);
        }
        .badge-warn {
            background: var(--red-glow);
            color: var(--red);
            border: 1px solid rgba(248, 113, 113, 0.2);
        }

        /* Loading */
        .loading-overlay {
            display: none;
            position: fixed; top: 0; left: 0; width: 100%; height: 100%;
            background: rgba(12, 14, 26, 0.85);
            backdrop-filter: blur(8px);
            -webkit-backdrop-filter: blur(8px);
            z-index: 100;
            justify-content: center;
            align-items: center;
            flex-direction: column;
            gap: 20px;
        }
        .loading-overlay.active { display: flex; }
        .spinner {
            width: 48px; height: 48px;
            border: 2px solid rgba(255, 255, 255, 0.08);
            border-top-color: var(--accent-light);
            border-radius: 50%;
            animation: spin 0.9s linear infinite;
        }
        @keyframes spin { to { transform: rotate(360deg); } }
        .loading-text {
            font-size: 13px;
            color: var(--text-dim);
            letter-spacing: 3px;
            font-weight: 300;
        }

        /* Footer */
        .footer {
            text-align: center;
            margin-top: 40px;
            padding-top: 20px;
            color: var(--text-dim);
            font-size: 11px;
            font-weight: 300;
            letter-spacing: 1px;
        }

        /* 動畫 */
        .card { animation: fadeUp 0.5s ease both; }
        @keyframes fadeUp {
            from { opacity: 0; transform: translateY(16px); }
            to { opacity: 1; transform: translateY(0); }
        }
    </style>
</head>
<body>
    <div class="loading-overlay" id="loading">
        <div class="spinner"></div>
        <div class="loading-text">對帳處理中</div>
    </div>

    <div class="container">
        <div class="header">
            <h1>對帳系統</h1>
            <div class="subtitle">Daily Invoice Reconciliation</div>
        </div>

        {% if not result and not error %}
        <form method="POST" enctype="multipart/form-data" id="mainForm">
            <div class="card">
                <h2>上傳對帳檔案</h2>

                <label>發票紀錄</label>
                <div class="upload-zone" id="zone1">
                    <div class="icon">&#128196;</div>
                    <div class="text">點擊或拖曳檔案至此</div>
                    <div class="filename" id="fname1"></div>
                    <input type="file" name="invoice_file" accept=".xlsx,.xls,.csv" required
                           onchange="showName(this, 'fname1', 'zone1')">
                </div>
                <div class="hint">支援 XLSX / CSV 格式</div>

                <label>8591交易紀錄</label>
                <div class="upload-zone" id="zone2">
                    <div class="icon">&#128178;</div>
                    <div class="text">點擊或拖曳檔案至此</div>
                    <div class="filename" id="fname2"></div>
                    <input type="file" name="trade_file" accept=".xlsx,.xls,.csv" required
                           onchange="showName(this, 'fname2', 'zone2')">
                </div>
                <div class="hint">支援 XLSX / CSV 格式</div>
            </div>
            <div class="btn-center">
                <button type="submit" class="btn btn-primary">開始對帳</button>
            </div>
        </form>
        {% endif %}

        {% if result %}
        <div class="card">
            <h2>對帳結果</h2>
            {% if has_diff == False %}
            <span class="badge badge-ok">ALL MATCHED</span>
            {% else %}
            <span class="badge badge-warn">MISMATCH FOUND</span>
            {% endif %}
            <div class="result-box {{ 'ok' if has_diff == False else 'warn' }}">{{ result }}</div>
            <div class="actions">
                <a href="/export" class="btn btn-export">匯出 Excel 報表</a>
                <a href="/" class="btn btn-outline">重新對帳</a>
            </div>
        </div>
        {% endif %}

        {% if error %}
        <div class="card">
            <h2>錯誤</h2>
            <span class="badge badge-warn">ERROR</span>
            <div class="result-box error">{{ error }}</div>
            <div class="actions">
                <a href="/" class="btn btn-outline">返回</a>
            </div>
        </div>
        {% endif %}

        <div class="footer">Invoice Reconciliation System v1.0</div>
    </div>

    <script>
    function showName(input, fnameId, zoneId) {
        const zone = document.getElementById(zoneId);
        const fname = document.getElementById(fnameId);
        if (input.files.length > 0) {
            fname.textContent = input.files[0].name;
            fname.style.display = 'block';
            zone.querySelector('.text').style.display = 'none';
            zone.querySelector('.icon').textContent = '\u2705';
        }
    }
    document.querySelectorAll('.upload-zone').forEach(zone => {
        zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
        zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
        zone.addEventListener('drop', () => zone.classList.remove('drag-over'));
    });
    const form = document.getElementById('mainForm');
    if (form) {
        form.addEventListener('submit', () => {
            document.getElementById('loading').classList.add('active');
        });
    }
    </script>
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
        has_diff = (len(result['amount_mismatch']) + len(result['only_in_xlsx']) + len(result['only_in_csv']) + len(result.get('reissued_invoices', []))) > 0

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
