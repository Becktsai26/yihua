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
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --navy: #0f1b3d;
            --navy-light: #162044;
            --navy-card: #1a2650;
            --border: #2a3a6a;
            --gold: #d4a843;
            --gold-light: #f0c75e;
            --gold-dim: rgba(212, 168, 67, 0.15);
            --blue-accent: #4a7dff;
            --green: #34c77b;
            --green-dim: rgba(52, 199, 123, 0.12);
            --red: #ef5565;
            --red-dim: rgba(239, 85, 101, 0.12);
            --text: #e8ecf4;
            --text-secondary: #8b99b8;
        }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: "Noto Sans TC", -apple-system, "Microsoft JhengHei", sans-serif;
            background: var(--navy);
            color: var(--text);
            min-height: 100vh;
        }

        /* 頂部裝飾線 */
        .top-bar {
            height: 4px;
            background: linear-gradient(90deg, var(--gold), var(--gold-light), var(--gold));
        }

        .container { max-width: 700px; margin: 0 auto; padding: 32px 20px; }

        /* Logo / 標題區 */
        .header { text-align: center; margin-bottom: 36px; }
        .header .logo {
            display: inline-flex;
            align-items: center;
            gap: 12px;
            margin-bottom: 10px;
        }
        .header .logo-icon {
            width: 48px; height: 48px;
            background: linear-gradient(135deg, var(--gold), var(--gold-light));
            border-radius: 12px;
            display: flex; align-items: center; justify-content: center;
            font-size: 24px;
            box-shadow: 0 4px 16px rgba(212, 168, 67, 0.3);
        }
        .header h1 {
            font-size: 26px;
            font-weight: 700;
            color: var(--text);
            letter-spacing: 1px;
        }
        .header h1 span { color: var(--gold); }
        .header .subtitle {
            font-size: 13px;
            color: var(--text-secondary);
            margin-top: 4px;
        }

        /* 卡片 */
        .card {
            background: var(--navy-card);
            border: 1px solid var(--border);
            border-radius: 14px;
            padding: 28px;
            margin-bottom: 20px;
            position: relative;
        }
        .card h2 {
            font-size: 15px;
            font-weight: 600;
            color: var(--gold);
            margin-bottom: 20px;
            padding-bottom: 12px;
            border-bottom: 1px solid rgba(212, 168, 67, 0.2);
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .card h2 .dot {
            width: 6px; height: 6px;
            background: var(--gold);
            border-radius: 50%;
        }

        /* 檔案上傳 */
        label { display: block; font-weight: 500; margin-bottom: 8px; color: var(--text); font-size: 14px; }
        .upload-zone {
            width: 100%;
            padding: 22px;
            border: 2px dashed var(--border);
            border-radius: 10px;
            margin-bottom: 6px;
            background: rgba(212, 168, 67, 0.03);
            cursor: pointer;
            transition: all 0.25s;
            text-align: center;
            position: relative;
        }
        .upload-zone:hover, .upload-zone.drag-over {
            border-color: var(--gold);
            background: var(--gold-dim);
        }
        .upload-zone .icon { font-size: 26px; margin-bottom: 4px; }
        .upload-zone .text { color: var(--text-secondary); font-size: 13px; }
        .upload-zone .filename {
            color: var(--gold-light);
            font-size: 14px;
            font-weight: 600;
            display: none;
        }
        .upload-zone input[type="file"] {
            position: absolute; top: 0; left: 0; width: 100%; height: 100%;
            opacity: 0; cursor: pointer;
        }
        .hint { font-size: 11px; color: var(--text-secondary); margin-bottom: 18px; }

        /* 按鈕 */
        .btn {
            display: inline-block;
            padding: 13px 34px;
            border: none;
            border-radius: 8px;
            font-size: 15px;
            font-weight: 600;
            cursor: pointer;
            text-decoration: none;
            transition: all 0.25s;
        }
        .btn-primary {
            background: linear-gradient(135deg, var(--gold), var(--gold-light));
            color: var(--navy);
            box-shadow: 0 4px 16px rgba(212, 168, 67, 0.3);
        }
        .btn-primary:hover {
            box-shadow: 0 6px 24px rgba(212, 168, 67, 0.45);
            transform: translateY(-2px);
        }
        .btn-export {
            background: linear-gradient(135deg, var(--green), #2db86a);
            color: #fff;
            box-shadow: 0 4px 16px rgba(52, 199, 123, 0.25);
        }
        .btn-export:hover {
            box-shadow: 0 6px 24px rgba(52, 199, 123, 0.4);
            transform: translateY(-2px);
        }
        .btn-outline {
            background: transparent;
            color: var(--text-secondary);
            border: 1px solid var(--border);
        }
        .btn-outline:hover { border-color: var(--gold); color: var(--gold); }
        .btn-center { text-align: center; }
        .actions { display: flex; gap: 14px; justify-content: center; margin-top: 20px; flex-wrap: wrap; }

        /* 統計卡片 */
        .stats {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 12px;
            margin-bottom: 20px;
        }
        .stat-item {
            background: rgba(255,255,255,0.03);
            border: 1px solid var(--border);
            border-radius: 10px;
            padding: 14px;
            text-align: center;
        }
        .stat-item .value {
            font-size: 24px;
            font-weight: 700;
            margin-bottom: 2px;
        }
        .stat-item .label {
            font-size: 11px;
            color: var(--text-secondary);
        }
        .stat-item.ok .value { color: var(--green); }
        .stat-item.warn .value { color: var(--red); }
        .stat-item.neutral .value { color: var(--gold); }

        /* 結果文字 */
        .result-box {
            background: rgba(0, 0, 0, 0.2);
            border: 1px solid var(--border);
            border-radius: 10px;
            padding: 18px;
            font-family: "Courier New", monospace;
            font-size: 13px;
            line-height: 1.8;
            white-space: pre-line;
        }
        .result-box.ok {
            border-color: rgba(52, 199, 123, 0.3);
        }
        .result-box.warn {
            border-color: rgba(239, 85, 101, 0.3);
        }
        .result-box.error {
            border-color: rgba(239, 85, 101, 0.3);
        }

        /* 狀態標籤 */
        .badge {
            display: inline-block;
            font-size: 11px;
            font-weight: 600;
            padding: 3px 12px;
            border-radius: 4px;
            margin-bottom: 14px;
            letter-spacing: 1px;
        }
        .badge-ok { background: var(--green-dim); color: var(--green); }
        .badge-warn { background: var(--red-dim); color: var(--red); }

        /* Loading */
        .loading-overlay {
            display: none;
            position: fixed; top: 0; left: 0; width: 100%; height: 100%;
            background: rgba(15, 27, 61, 0.9);
            z-index: 100;
            justify-content: center;
            align-items: center;
            flex-direction: column;
            gap: 16px;
        }
        .loading-overlay.active { display: flex; }
        .spinner {
            width: 44px; height: 44px;
            border: 3px solid var(--border);
            border-top-color: var(--gold);
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
        }
        @keyframes spin { to { transform: rotate(360deg); } }
        .loading-text {
            font-size: 14px;
            color: var(--gold);
            letter-spacing: 2px;
        }

        /* Footer */
        .footer {
            text-align: center;
            margin-top: 36px;
            padding-top: 20px;
            border-top: 1px solid rgba(212, 168, 67, 0.1);
            color: var(--text-secondary);
            font-size: 12px;
        }
    </style>
</head>
<body>
    <div class="top-bar"></div>
    <div class="loading-overlay" id="loading">
        <div class="spinner"></div>
        <div class="loading-text">對帳處理中...</div>
    </div>

    <div class="container">
        <div class="header">
            <div class="logo">
                <div class="logo-icon">&#128176;</div>
                <div>
                    <h1>8591 <span>對帳系統</span></h1>
                    <div class="subtitle">每日發票自動核對工具</div>
                </div>
            </div>
        </div>

        {% if not result and not error %}
        <form method="POST" enctype="multipart/form-data" id="mainForm">
            <div class="card">
                <h2><span class="dot"></span> 上傳對帳檔案</h2>

                <label>發票檔案</label>
                <div class="upload-zone" id="zone1">
                    <div class="icon">&#128196;</div>
                    <div class="text">點擊或拖曳檔案至此</div>
                    <div class="filename" id="fname1"></div>
                    <input type="file" name="invoice_file" accept=".xlsx,.xls,.csv" required
                           onchange="showName(this, 'fname1', 'zone1')">
                </div>
                <div class="hint">支援 XLSX / CSV 格式</div>

                <label>8591 交易檔案</label>
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
            <h2><span class="dot"></span> 對帳結果</h2>
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
            <h2><span class="dot"></span> 錯誤</h2>
            <span class="badge badge-warn">ERROR</span>
            <div class="result-box error">{{ error }}</div>
            <div class="actions">
                <a href="/" class="btn btn-outline">返回</a>
            </div>
        </div>
        {% endif %}

        <div class="footer">8591 Invoice Reconciliation System v1.0</div>
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
