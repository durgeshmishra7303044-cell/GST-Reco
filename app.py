from flask import Flask, request, send_file, render_template_string, jsonify
import os, uuid, traceback, tempfile
from reconcile import reconcile
from export import export_excel

app = Flask(__name__)
UPLOAD_FOLDER = tempfile.gettempdir()
OUTPUT_FOLDER = tempfile.gettempdir()

HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>GST ITC Reconciliation Tool</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', sans-serif; background: #f0f4f8; color: #1a202c; }
  .header { background: linear-gradient(135deg, #1a365d 0%, #2b6cb0 100%); color: white; padding: 24px 40px; display: flex; align-items: center; gap: 16px; }
  .header h1 { font-size: 22px; font-weight: 700; }
  .header p  { font-size: 13px; opacity: 0.8; margin-top: 4px; }
  .logo { font-size: 36px; }
  .container { max-width: 900px; margin: 36px auto; padding: 0 20px; }
  .card { background: white; border-radius: 12px; box-shadow: 0 2px 12px rgba(0,0,0,.08); padding: 32px; margin-bottom: 24px; }
  .card h2 { font-size: 16px; font-weight: 700; color: #2b6cb0; margin-bottom: 20px; border-bottom: 2px solid #ebf4ff; padding-bottom: 10px; }
  .upload-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
  .upload-box { border: 2px dashed #bee3f8; border-radius: 10px; padding: 28px 20px; text-align: center; cursor: pointer; transition: all .2s; position: relative; background: #f7fbff; }
  .upload-box:hover { border-color: #2b6cb0; background: #ebf8ff; }
  .upload-box input { position: absolute; inset: 0; opacity: 0; cursor: pointer; width: 100%; height: 100%; }
  .upload-box .icon { font-size: 32px; margin-bottom: 10px; }
  .upload-box .label { font-size: 13px; font-weight: 600; color: #2d3748; margin-bottom: 4px; }
  .upload-box .hint  { font-size: 11px; color: #718096; }
  .upload-box .fname { font-size: 12px; color: #2b6cb0; font-weight: 600; margin-top: 8px; word-break: break-all; }
  .btn-reconcile { width: 100%; padding: 14px; font-size: 15px; font-weight: 700; background: linear-gradient(135deg, #1a7a4a, #2f855a); color: white; border: none; border-radius: 8px; cursor: pointer; margin-top: 24px; transition: opacity .2s; letter-spacing: .5px; }
  .btn-reconcile:hover { opacity: .9; }
  .btn-reconcile:disabled { background: #a0aec0; cursor: not-allowed; }
  .progress { display: none; margin-top: 16px; }
  .progress-bar-bg { background: #e2e8f0; border-radius: 999px; height: 8px; overflow: hidden; }
  .progress-bar { background: linear-gradient(90deg,#2b6cb0,#2f855a); height: 100%; width: 0; border-radius: 999px; transition: width .4s ease; }
  .progress-msg { font-size: 12px; color: #718096; margin-top: 6px; text-align: center; }
  .results { display: none; }
  .stats-grid { display: grid; grid-template-columns: repeat(4,1fr); gap: 12px; margin-bottom: 20px; }
  .stat-box { border-radius: 8px; padding: 16px; text-align: center; }
  .stat-box.green  { background: #f0fff4; border: 1px solid #c6f6d5; }
  .stat-box.red    { background: #fff5f5; border: 1px solid #fed7d7; }
  .stat-box.orange { background: #fffaf0; border: 1px solid #feebc8; }
  .stat-box .num { font-size: 26px; font-weight: 800; }
  .stat-box.green .num { color: #276749; }
  .stat-box.red   .num { color: #c53030; }
  .stat-box.orange .num { color: #c05621; }
  .stat-box .lbl { font-size: 11px; color: #718096; margin-top: 4px; }
  .stat-box .itc { font-size: 12px; font-weight: 700; margin-top: 6px; color: #c53030; }
  .btn-download { width: 100%; padding: 13px; font-size: 14px; font-weight: 700; background: linear-gradient(135deg, #2b6cb0, #3182ce); color: white; border: none; border-radius: 8px; cursor: pointer; letter-spacing: .5px; transition: opacity .2s; }
  .btn-download:hover { opacity: .9; }
  .error-box { display: none; background: #fff5f5; border: 1px solid #feb2b2; border-radius: 8px; padding: 16px; color: #c53030; font-size: 12px; margin-top: 16px; white-space: pre-wrap; }
  .steps { background: #f7fafc; border-radius: 8px; padding: 20px; }
  .steps h3 { font-size: 13px; font-weight: 700; color: #4a5568; margin-bottom: 12px; }
  .step { display: flex; gap: 12px; margin-bottom: 10px; align-items: flex-start; }
  .step-num { background: #2b6cb0; color: white; border-radius: 50%; width: 22px; height: 22px; font-size: 11px; font-weight: 700; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
  .step-text { font-size: 12px; color: #4a5568; line-height: 1.5; }
  @media(max-width:600px) { .upload-grid { grid-template-columns: 1fr; } .stats-grid { grid-template-columns: 1fr 1fr; } }
</style>
</head>
<body>
<div class="header">
  <div class="logo">📊</div>
  <div>
    <h1>GST ITC Reconciliation Tool</h1>
    <p>Auto-reconcile GSTR-2B vs Purchase Register — Upload · Click · Download</p>
  </div>
</div>
<div class="container">
  <div class="card">
    <h2>📁 Upload Files</h2>
    <div class="upload-grid">
      <div class="upload-box">
        <input type="file" id="file2b" accept=".xlsx,.xls" onchange="setFile('2b',this)">
        <div class="icon">🏛️</div>
        <div class="label">GSTR-2B File</div>
        <div class="hint">Downloaded from GST Portal (.xlsx)</div>
        <div class="fname" id="name2b">No file chosen</div>
      </div>
      <div class="upload-box">
        <input type="file" id="filebk" accept=".xlsx,.xls" onchange="setFile('bk',this)">
        <div class="icon">📒</div>
        <div class="label">Purchase Register</div>
        <div class="hint">Tally export or manual Excel (.xlsx)</div>
        <div class="fname" id="namebk">No file chosen</div>
      </div>
    </div>
    <button class="btn-reconcile" id="btnRec" onclick="runRecon()" disabled>🔄 Run Reconciliation</button>
    <div class="progress" id="progress">
      <div class="progress-bar-bg"><div class="progress-bar" id="pbar"></div></div>
      <div class="progress-msg" id="pmsg">Uploading files...</div>
    </div>
    <div class="error-box" id="errBox"></div>
  </div>
  <div class="card results" id="results">
    <h2>✅ Reconciliation Complete</h2>
    <div class="stats-grid" id="statsGrid"></div>
    <button class="btn-download" onclick="downloadReport()">⬇️ Download Color-Coded Excel Report</button>
  </div>
  <div class="card">
    <div class="steps">
      <h3>ℹ️ How it works</h3>
      <div class="step"><div class="step-num">1</div><div class="step-text"><b>Upload GSTR-2B</b> — Excel downloaded from GST portal</div></div>
      <div class="step"><div class="step-num">2</div><div class="step-text"><b>Upload Purchase Register</b> — Tally export with GSTIN, invoice no, date</div></div>
      <div class="step"><div class="step-num">3</div><div class="step-text"><b>Click Reconcile</b> — Key: GSTIN + Invoice No + Date (auto format fix)</div></div>
      <div class="step"><div class="step-num">4</div><div class="step-text"><b>Download Report</b> — Matched ✅ | 2B not in Books ❌ | Books not in 2B ❌ | Format Mismatch ⚠️</div></div>
    </div>
  </div>
</div>
<script>
let files = {}, reportId = null;
function setFile(key, inp) {
  files[key] = inp.files[0];
  document.getElementById('name' + key).textContent = inp.files[0].name;
  document.getElementById('btnRec').disabled = !(files['2b'] && files['bk']);
}
function setProgress(pct, msg) {
  document.getElementById('pbar').style.width = pct + '%';
  document.getElementById('pmsg').textContent = msg;
}
async function runRecon() {
  document.getElementById('errBox').style.display = 'none';
  document.getElementById('results').style.display = 'none';
  document.getElementById('progress').style.display = 'block';
  document.getElementById('btnRec').disabled = true;
  setProgress(10, 'Uploading files...');
  const fd = new FormData();
  fd.append('gstr2b', files['2b']);
  fd.append('books', files['bk']);
  try {
    setProgress(30, 'Reading GSTR-2B...');
    const res = await fetch('/reconcile', { method: 'POST', body: fd });
    setProgress(70, 'Matching invoices...');
    const data = await res.json();
    setProgress(100, 'Done!');
    if (data.error) throw new Error(data.error);
    reportId = data.report_id;
    showResults(data.summary);
  } catch(e) {
    document.getElementById('progress').style.display = 'none';
    document.getElementById('errBox').textContent = 'Error: ' + e.message;
    document.getElementById('errBox').style.display = 'block';
  } finally {
    document.getElementById('btnRec').disabled = false;
  }
}
function showResults(s) {
  document.getElementById('results').style.display = 'block';
  const fmt = n => '₹' + Number(n).toLocaleString('en-IN', {maximumFractionDigits:0});
  document.getElementById('statsGrid').innerHTML = `
    <div class="stat-box green"><div class="num">${s.matched}</div><div class="lbl">Matched ✅</div></div>
    <div class="stat-box orange"><div class="num">${s.format_mismatch}</div><div class="lbl">Format Mismatch ⚠️</div></div>
    <div class="stat-box red"><div class="num">${s.not_in_books_count}</div><div class="lbl">2B Not In Books</div><div class="itc">${fmt(s.not_in_books_itc)} ITC</div></div>
    <div class="stat-box red"><div class="num">${s.not_in_2b_count}</div><div class="lbl">Books Not In 2B</div><div class="itc">${fmt(s.not_in_2b_itc)} ITC</div></div>
  `;
  document.getElementById('results').scrollIntoView({behavior:'smooth'});
}
function downloadReport() { if (reportId) window.location = '/download/' + reportId; }
</script>
</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML)

@app.route('/reconcile', methods=['POST'])
def run_reconcile():
    try:
        f2b = request.files['gstr2b']
        fbk = request.files['books']
        uid = str(uuid.uuid4())[:8]
        path_2b = os.path.join(UPLOAD_FOLDER, f'{uid}_2b.xlsx')
        path_bk = os.path.join(UPLOAD_FOLDER, f'{uid}_bk.xlsx')
        f2b.save(path_2b)
        fbk.save(path_bk)
        results = reconcile(path_2b, path_bk)
        out_path = os.path.join(OUTPUT_FOLDER, f'GST_Recon_{uid}.xlsx')
        export_excel(results, out_path)
        try:
            os.remove(path_2b)
            os.remove(path_bk)
        except Exception:
            pass
        return jsonify({'report_id': uid, 'summary': results['summary']})
    except Exception as e:
        return jsonify({'error': str(e) + '\n\nTraceback:\n' + traceback.format_exc()}), 500

@app.route('/download/<uid>')
def download(uid):
    if not uid.replace('-','').isalnum():
        return 'Invalid ID', 400
    path = os.path.join(OUTPUT_FOLDER, f'GST_Recon_{uid}.xlsx')
    if os.path.exists(path):
        return send_file(path, as_attachment=True, download_name=f'GST_Recon_{uid}.xlsx')
    return 'File not found — please re-run reconciliation', 404

if __name__ == '__main__':
    print("\n" + "="*55)
    print("  GST Reconciliation Tool is RUNNING")
    print("  Open your browser: http://localhost:5000")
    print("="*55 + "\n")
    app.run(debug=False, port=5000)
