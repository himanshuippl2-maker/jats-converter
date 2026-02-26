import os, uuid, traceback
from flask import Flask, request, jsonify, send_file, render_template_string
from werkzeug.utils import secure_filename
import converter as cv

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024  # 20MB max

UPLOAD_FOLDER = '/tmp/jats_uploads'
OUTPUT_FOLDER = '/tmp/jats_outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {'docx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# ============================================================
# WEB UI
# ============================================================
HTML = '''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>DOCX → JATS XML Converter</title>
<link href="https://fonts.googleapis.com/css2?family=Sora:wght@300;400;600;700&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
<style>
*{margin:0;padding:0;box-sizing:border-box}
:root{
  --bg:#0a0e1a;--surface:#111827;--surface2:#1a2235;
  --border:rgba(99,179,237,.12);--accent:#63b3ed;--accent2:#68d391;
  --accent3:#f6ad55;--danger:#fc8181;--text:#e2e8f0;--muted:#718096;
}
body{font-family:'Sora',sans-serif;background:var(--bg);color:var(--text);min-height:100vh}
.header{background:linear-gradient(135deg,#0f1f3d,#0a0e1a);border-bottom:1px solid var(--border);padding:20px 32px;display:flex;align-items:center;gap:14px}
.header h1{font-size:18px;font-weight:700;color:var(--accent)}
.header p{font-size:12px;color:var(--muted);margin-top:2px}
.badge{background:rgba(99,179,237,.1);border:1px solid rgba(99,179,237,.3);color:var(--accent);font-size:10px;font-weight:700;padding:3px 10px;border-radius:50px;letter-spacing:1px;margin-left:auto}
.wrap{max-width:860px;margin:0 auto;padding:32px 20px;display:grid;grid-template-columns:1fr 1fr;gap:24px}
@media(max-width:640px){.wrap{grid-template-columns:1fr}}
.card{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:24px}
.sec-title{font-size:11px;font-weight:700;color:var(--muted);letter-spacing:2px;text-transform:uppercase;margin-bottom:14px}
.upload-zone{border:2px dashed rgba(99,179,237,.3);border-radius:10px;padding:32px 16px;text-align:center;cursor:pointer;transition:.2s;background:rgba(99,179,237,.03)}
.upload-zone:hover,.upload-zone.drag{border-color:var(--accent);background:rgba(99,179,237,.08)}
.upload-zone .icon{font-size:36px;margin-bottom:10px}
.upload-zone p{font-size:13px;color:var(--muted);line-height:1.6}
.upload-zone strong{color:var(--text)}
#fileInput{display:none}
.field{display:flex;flex-direction:column;gap:6px;margin-bottom:12px}
.field label{font-size:11px;font-weight:600;color:var(--muted);letter-spacing:.5px}
.field input,.field select{background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:9px 12px;color:var(--text);font-family:'Sora',sans-serif;font-size:12px;outline:none;transition:.2s;width:100%}
.field input:focus,.field select:focus{border-color:var(--accent)}
.field select option{background:#1a2235}
.row2{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px}
.btn{background:linear-gradient(135deg,#3182ce,#2b6cb0);color:#fff;border:none;padding:13px;border-radius:10px;font-family:'Sora',sans-serif;font-size:14px;font-weight:700;cursor:pointer;width:100%;transition:.2s;box-shadow:0 4px 20px rgba(49,130,206,.3);margin-top:8px}
.btn:hover{transform:translateY(-1px);box-shadow:0 8px 28px rgba(49,130,206,.4)}
.btn:disabled{opacity:.5;cursor:not-allowed;transform:none}
.result{display:none;margin-top:20px}
.result .dl-btn{background:linear-gradient(135deg,var(--accent2),#48bb78);color:#0a2010;border:none;padding:11px 20px;border-radius:8px;font-family:'Sora',sans-serif;font-size:13px;font-weight:700;cursor:pointer;text-decoration:none;display:inline-block;transition:.2s}
.result .dl-btn:hover{transform:translateY(-1px)}
.error-box{background:rgba(252,129,129,.08);border:1px solid rgba(252,129,129,.2);border-radius:10px;padding:14px;font-size:12px;color:var(--danger);display:none;margin-top:12px}
.progress{display:none;margin-top:16px}
.bar{height:4px;background:var(--surface2);border-radius:4px;overflow:hidden;margin-top:6px}
.bar-fill{height:100%;background:linear-gradient(90deg,var(--accent),var(--accent2));width:0%;transition:width .4s;border-radius:4px}
.prog-label{font-size:12px;color:var(--muted)}
.stats{display:flex;flex-wrap:wrap;gap:12px;margin-top:14px}
.stat{background:var(--surface2);border-radius:8px;padding:8px 14px;font-size:12px}
.stat span{color:var(--accent);font-weight:700}
.api-card{grid-column:1/-1}
.code{background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:14px;font-family:'JetBrains Mono',monospace;font-size:12px;color:#98c379;overflow-x:auto;white-space:pre;line-height:1.7;margin-top:10px}
::-webkit-scrollbar{width:5px;height:5px}
::-webkit-scrollbar-thumb{background:rgba(255,255,255,.08);border-radius:3px}
</style>
</head>
<body>
<div class="header">
  <div style="font-size:26px">🧬</div>
  <div>
    <h1>DOCX → JATS XML Converter</h1>
    <p>PMC-Compliant Medical Journal Converter · JATS 1.2</p>
  </div>
  <div class="badge">PMC READY</div>
</div>

<div class="wrap">
  <!-- Upload + Convert -->
  <div class="card">
    <div class="sec-title">📂 Upload & Convert</div>

    <div class="upload-zone" id="uploadZone" onclick="document.getElementById('fileInput').click()">
      <div class="icon">📄</div>
      <p><strong>Click to upload .docx</strong><br>or drag & drop here</p>
    </div>
    <input type="file" id="fileInput" accept=".docx">

    <div class="progress" id="progress">
      <div class="prog-label" id="progLabel">Processing...</div>
      <div class="bar"><div class="bar-fill" id="barFill"></div></div>
    </div>

    <div class="error-box" id="errorBox"></div>

    <div class="result" id="result">
      <div class="stats" id="stats"></div>
      <br>
      <a class="dl-btn" id="dlBtn" href="#">⬇️ Download JATS XML</a>
    </div>
  </div>

  <!-- Journal Metadata -->
  <div class="card">
    <div class="sec-title">📰 Journal Metadata</div>

    <div class="field">
      <label>Article Type</label>
      <select id="articleType">
        <option value="research-article">Original Research Article</option>
        <option value="review-article">Review Article</option>
        <option value="case-report">Case Report</option>
        <option value="letter">Letter to Editor</option>
        <option value="editorial">Editorial</option>
        <option value="brief-report">Brief Report</option>
        <option value="systematic-review">Systematic Review</option>
      </select>
    </div>
    <div class="field"><label>Journal Full Name</label><input id="journal" placeholder="IP Indian Journal of Clinical and Experimental Dermatology"></div>
    <div class="field"><label>Publisher</label><input id="publisher" placeholder="IP Innovative Publication" value="IP Innovative Publication"></div>
    <div class="field"><label>Journal URL</label><input id="journalUrl" placeholder="https://ijced.org"></div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">
      <div class="field"><label>ISSN Print</label><input id="issnPrint" placeholder="2581-4710"></div>
      <div class="field"><label>ISSN Electronic</label><input id="issnElec" placeholder="2581-4729"></div>
    </div>
    <div class="field"><label>Article DOI</label><input id="doi" placeholder="10.18231/j.ijced.2025.001"></div>
    <div class="row2">
      <div class="field"><label>Volume</label><input id="volume" placeholder="11"></div>
      <div class="field"><label>Issue</label><input id="issue" placeholder="4"></div>
      <div class="field"><label>Year</label><input id="year" placeholder="2025" value="2025"></div>
    </div>
    <div class="row2">
      <div class="field"><label>Day</label><input id="day" placeholder="30"></div>
      <div class="field"><label>Month</label><input id="month" placeholder="12"></div>
      <div class="field"><label>License</label>
        <select id="license">
          <option value="cc-by-nc-4.0">CC BY-NC 4.0</option>
          <option value="cc-by-4.0">CC BY 4.0</option>
          <option value="cc-by-nc-nd-4.0">CC BY-NC-ND 4.0</option>
        </select>
      </div>
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">
      <div class="field"><label>First Page</label><input id="fpage" placeholder="473"></div>
      <div class="field"><label>Last Page</label><input id="lpage" placeholder="478"></div>
    </div>

    <button class="btn" id="convertBtn" onclick="doConvert()" disabled>🔄 Convert to JATS XML</button>
  </div>

  <!-- API Docs -->
  <div class="card api-card">
    <div class="sec-title">🔌 REST API</div>
    <p style="font-size:13px;color:var(--muted);line-height:1.7">
      POST a <code style="color:var(--accent)">.docx</code> file with journal metadata as form fields. Returns JATS XML file.
    </p>
    <div class="code">curl -X POST https://your-app.railway.app/api/convert \\
  -F "file=@article.docx" \\
  -F "journal=IP Indian Journal of Clinical and Experimental Dermatology" \\
  -F "issn_print=2581-4710" \\
  -F "issn_elec=2581-4729" \\
  -F "publisher=IP Innovative Publication" \\
  -F "doi=10.18231/j.ijced.2025.001" \\
  -F "volume=11" -F "issue=4" \\
  -F "year=2025" -F "month=12" -F "day=30" \\
  -F "fpage=473" -F "lpage=478" \\
  -F "article_type=research-article" \\
  -o output.xml</div>
    <p style="font-size:12px;color:var(--muted);margin-top:12px">
      ✅ Returns: <code style="color:var(--accent2)">application/xml</code> &nbsp;|&nbsp;
      ❌ Error: <code style="color:var(--danger)">{"error": "message"}</code>
    </p>
  </div>
</div>

<script>
let uploadedFile = null;

// Drag & drop
const zone = document.getElementById('uploadZone');
const fi = document.getElementById('fileInput');
zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag'); });
zone.addEventListener('dragleave', () => zone.classList.remove('drag'));
zone.addEventListener('drop', e => {
  e.preventDefault(); zone.classList.remove('drag');
  if (e.dataTransfer.files[0]) setFile(e.dataTransfer.files[0]);
});
fi.addEventListener('change', e => { if (e.target.files[0]) setFile(e.target.files[0]); });

function setFile(f) {
  if (!f.name.endsWith('.docx')) { showError('Only .docx files allowed'); return; }
  uploadedFile = f;
  zone.innerHTML = `<div class="icon">✅</div><p><strong>${f.name}</strong><br>${(f.size/1024).toFixed(1)} KB</p>`;
  document.getElementById('convertBtn').disabled = false;
  hideError();
}

async function doConvert() {
  if (!uploadedFile) return;
  document.getElementById('convertBtn').disabled = true;
  document.getElementById('result').style.display = 'none';
  hideError();
  showProgress(10, 'Uploading file...');

  const fd = new FormData();
  fd.append('file', uploadedFile);
  fd.append('journal', document.getElementById('journal').value);
  fd.append('publisher', document.getElementById('publisher').value);
  fd.append('journal_url', document.getElementById('journalUrl').value);
  fd.append('issn_print', document.getElementById('issnPrint').value);
  fd.append('issn_elec', document.getElementById('issnElec').value);
  fd.append('doi', document.getElementById('doi').value);
  fd.append('volume', document.getElementById('volume').value);
  fd.append('issue', document.getElementById('issue').value);
  fd.append('year', document.getElementById('year').value);
  fd.append('month', document.getElementById('month').value);
  fd.append('day', document.getElementById('day').value);
  fd.append('fpage', document.getElementById('fpage').value);
  fd.append('lpage', document.getElementById('lpage').value);
  fd.append('article_type', document.getElementById('articleType').value);
  fd.append('license', document.getElementById('license').value);

  try {
    showProgress(30, 'Parsing document...');
    const resp = await fetch('/api/convert', { method: 'POST', body: fd });

    if (!resp.ok) {
      const err = await resp.json();
      throw new Error(err.error || 'Conversion failed');
    }

    showProgress(90, 'Building XML...');
    const blob = await resp.blob();
    const url = URL.createObjectURL(blob);
    const fname = uploadedFile.name.replace('.docx', '-jats.xml');

    // Get stats from header
    const statsH = resp.headers.get('X-Stats');
    if (statsH) {
      try {
        const s = JSON.parse(statsH);
        document.getElementById('stats').innerHTML =
          `<div class="stat">👥 <span>${s.authors}</span> Authors</div>` +
          `<div class="stat">📝 <span>${s.sections}</span> Sections</div>` +
          `<div class="stat">📚 <span>${s.refs}</span> References</div>` +
          `<div class="stat">📊 <span>${s.tables}</span> Tables</div>` +
          `<div class="stat">📏 <span>${s.size}</span> KB</div>`;
      } catch(e) {}
    }

    document.getElementById('dlBtn').href = url;
    document.getElementById('dlBtn').download = fname;
    document.getElementById('result').style.display = 'block';
    showProgress(100, 'Done! ✅');
    setTimeout(() => document.getElementById('progress').style.display = 'none', 1500);

  } catch(err) {
    showError(err.message);
    document.getElementById('progress').style.display = 'none';
  }

  document.getElementById('convertBtn').disabled = false;
}

function showProgress(pct, label) {
  document.getElementById('progress').style.display = 'block';
  document.getElementById('barFill').style.width = pct + '%';
  document.getElementById('progLabel').textContent = label;
}
function showError(msg) {
  const b = document.getElementById('errorBox');
  b.textContent = '❌ ' + msg; b.style.display = 'block';
}
function hideError() {
  document.getElementById('errorBox').style.display = 'none';
}
</script>
</body>
</html>'''

@app.route('/')
def index():
    return render_template_string(HTML)

# ============================================================
# API ENDPOINT
# ============================================================
@app.route('/api/convert', methods=['POST'])
def api_convert():
    # Validate file
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded. Use field name: file'}), 400

    f = request.files['file']
    if not f.filename or not allowed_file(f.filename):
        return jsonify({'error': 'Only .docx files allowed'}), 400

    # Save uploaded file
    uid = str(uuid.uuid4())[:8]
    fname = secure_filename(f.filename)
    input_path = os.path.join(UPLOAD_FOLDER, f'{uid}_{fname}')
    output_path = os.path.join(OUTPUT_FOLDER, f'{uid}_jats.xml')
    f.save(input_path)

    try:
        # Build journal metadata from form fields
        TYPE_LABELS = {
            'research-article': 'Original Research Article',
            'review-article': 'Review Article',
            'case-report': 'Case Report',
            'letter': 'Letter to Editor',
            'editorial': 'Editorial',
            'brief-report': 'Brief Report',
            'systematic-review': 'Systematic Review',
        }
        article_type = request.form.get('article_type', 'research-article')

        LICENSE_URLS = {
            'cc-by-4.0': 'https://creativecommons.org/licenses/by/4.0/',
            'cc-by-nc-4.0': 'https://creativecommons.org/licenses/by-nc/4.0/',
            'cc-by-nc-nd-4.0': 'https://creativecommons.org/licenses/by-nc-nd/4.0/',
        }
        license_key = request.form.get('license', 'cc-by-nc-4.0')

        jm = {
            'name':        request.form.get('journal', ''),
            'publisher':   request.form.get('publisher', 'IP Innovative Publication'),
            'journalUrl':  request.form.get('journal_url', ''),
            'issnPrint':   request.form.get('issn_print', ''),
            'issnElec':    request.form.get('issn_elec', ''),
            'doi':         request.form.get('doi', ''),
            'volume':      request.form.get('volume', ''),
            'issue':       request.form.get('issue', ''),
            'year':        request.form.get('year', '2025'),
            'month':       request.form.get('month', ''),
            'day':         request.form.get('day', ''),
            'fpage':       request.form.get('fpage', ''),
            'lpage':       request.form.get('lpage', ''),
            'articleType': article_type,
            'articleTypeLabel': TYPE_LABELS.get(article_type, article_type),
            'licenseUrl':  LICENSE_URLS.get(license_key, LICENSE_URLS['cc-by-nc-4.0']),
        }

        # Use CrossRef only if explicitly requested (slow, skip by default)
        use_crossref = request.form.get('crossref', 'false').lower() == 'true'

        # Parse and convert
        parsed = cv.parse_docx(input_path, use_crossref=use_crossref)
        xml = cv.build_xml(parsed, jm)
        xml = cv.post_process(xml)

        # Write output
        with open(output_path, 'w', encoding='utf-8') as out:
            out.write(xml)

        # Stats for frontend
        import json
        stats = json.dumps({
            'authors':  len(parsed.get('authors', [])),
            'sections': len(parsed.get('sections', [])),
            'refs':     len(parsed.get('references', [])),
            'tables':   len(parsed.get('tables', [])),
            'size':     round(len(xml) / 1024, 1),
        })

        out_fname = fname.replace('.docx', '-jats.xml')
        response = send_file(
            output_path,
            mimetype='application/xml',
            as_attachment=True,
            download_name=out_fname,
        )
        response.headers['X-Stats'] = stats
        return response

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

    finally:
        # Cleanup input file
        if os.path.exists(input_path):
            os.remove(input_path)

# ============================================================
# HEALTH CHECK (Railway/Render use this)
# ============================================================
@app.route('/health')
def health():
    return jsonify({'status': 'ok', 'version': '3.0'})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
