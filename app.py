"""
Inventory Reconciliation Web App
Upload QBO files + TicketVault report → download Excel reconciliation report.
"""

import os
import traceback
from datetime import datetime
from io import BytesIO

import pandas as pd
from flask import Flask, render_template_string, request, send_file, jsonify

from reconciler import (
    parse_qbo_consolidated,
    parse_qbo_single,
    parse_ticketvault,
    build_report,
)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB

# ---------------------------------------------------------------------------
# HTML Template
# ---------------------------------------------------------------------------
HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>Inventory Reconciliation</title>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

    body {
      font-family: 'Inter', 'Segoe UI', system-ui, sans-serif;
      background: #eef0f4;
      color: #1a2233;
      min-height: 100vh;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      padding: 40px 16px;
    }

    .card {
      background: #fff;
      border-radius: 16px;
      box-shadow: 0 2px 16px rgba(0,0,0,0.08);
      padding: 36px 40px 32px;
      width: 100%;
      max-width: 560px;
    }

    h1 {
      font-size: 1.25rem;
      color: #111827;
      font-weight: 700;
      margin-bottom: 4px;
      letter-spacing: -0.01em;
    }
    .subtitle {
      font-size: 0.875rem;
      color: #6b7280;
      margin-bottom: 28px;
    }

    .field-group {
      margin-bottom: 20px;
    }

    .field-label {
      font-size: 0.72rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      color: #6b7280;
      margin-bottom: 8px;
      display: block;
    }

    /* Upload zone */
    .upload-zone {
      border: 1.5px dashed #c9d0db;
      border-radius: 10px;
      padding: 28px 20px;
      background: #f9fafb;
      cursor: pointer;
      transition: border-color 0.2s, background 0.2s;
      position: relative;
      text-align: center;
    }
    .upload-zone:hover, .upload-zone.dragover {
      border-color: #6b7aab;
      background: #f1f3fa;
    }
    .upload-zone.has-files {
      border-style: solid;
      border-color: #9aacd4;
      background: #f4f6fb;
      padding: 14px 16px;
      text-align: left;
    }
    .upload-zone input[type="file"] {
      position: absolute;
      inset: 0;
      opacity: 0;
      cursor: pointer;
      width: 100%;
      height: 100%;
    }

    .upload-icon {
      width: 36px;
      height: 36px;
      margin: 0 auto 10px;
      color: #9ca3af;
      pointer-events: none;
    }
    .upload-zone.has-files .upload-icon { display: none; }

    .zone-hint {
      font-size: 0.875rem;
      color: #6b7280;
      pointer-events: none;
      line-height: 1.5;
    }
    .zone-hint a { color: #4b5fa6; text-decoration: none; font-weight: 500; }
    .zone-hint small { display: block; font-size: 0.78rem; color: #9ca3af; margin-top: 2px; }
    .upload-zone.has-files .zone-hint { display: none; }

    .file-list {
      font-size: 0.84rem;
      color: #374151;
      display: none;
    }
    .upload-zone.has-files .file-list { display: block; }
    .file-item {
      display: flex;
      align-items: center;
      gap: 8px;
      padding: 3px 0;
    }
    .file-item svg { flex-shrink: 0; color: #6b7aab; }
    .file-item span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }


    /* Button row */
    .btn-row {
      display: flex;
      gap: 10px;
      margin-top: 28px;
    }

    .btn {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 7px;
      padding: 11px 20px;
      border-radius: 8px;
      font-size: 0.9rem;
      font-weight: 600;
      border: none;
      cursor: pointer;
      transition: background 0.18s, box-shadow 0.18s;
      font-family: inherit;
    }
    .btn-primary {
      background: #5a6fa8;
      color: #fff;
      flex: 1;
    }
    .btn-primary:hover:not(:disabled) { background: #4a5d94; box-shadow: 0 2px 8px rgba(90,111,168,0.3); }
    .btn-primary:disabled { background: #b0b8cc; cursor: not-allowed; }

    .btn-clear {
      background: #fff;
      color: #374151;
      border: 1.5px solid #d1d5db;
      padding: 11px 22px;
      flex-shrink: 0;
    }
    .btn-clear:hover { background: #f3f4f6; border-color: #b0b8cc; }

    /* Status messages */
    .status-box {
      border-radius: 8px;
      padding: 12px 16px;
      font-size: 0.875rem;
      margin-top: 16px;
      display: none;
      line-height: 1.5;
    }
    .status-error   { background: #fef2f2; border: 1px solid #fecaca; color: #991b1b; }
    .status-success { background: #f0fdf4; border: 1px solid #bbf7d0; color: #166534; }
    .status-loading { background: #eff6ff; border: 1px solid #bfdbfe; color: #1e40af; }

    .spinner {
      display: inline-block;
      width: 14px; height: 14px;
      border: 2px solid rgba(30,64,175,0.25);
      border-top-color: #1e40af;
      border-radius: 50%;
      animation: spin 0.7s linear infinite;
      vertical-align: middle;
      margin-right: 4px;
    }
    @keyframes spin { to { transform: rotate(360deg); } }

    /* Checks info accordion */
    .checks-toggle {
      display: flex;
      align-items: center;
      gap: 6px;
      font-size: 0.8rem;
      color: #6b7280;
      cursor: pointer;
      margin-bottom: 20px;
      user-select: none;
      width: fit-content;
    }
    .checks-toggle:hover { color: #4b5fa6; }
    .checks-toggle svg { transition: transform 0.2s; }
    .checks-toggle.open svg { transform: rotate(90deg); }
    .checks-body {
      display: none;
      background: #f9fafb;
      border-radius: 8px;
      padding: 14px 16px;
      font-size: 0.83rem;
      color: #374151;
      margin-bottom: 20px;
      border: 1px solid #e5e7eb;
    }
    .checks-body.open { display: block; }
    .checks-body ol { padding-left: 18px; }
    .checks-body li { margin-bottom: 5px; line-height: 1.5; }
    .checks-body li:last-child { margin-bottom: 0; }
  </style>
</head>
<body>
  <div class="card">
    <h1>Inventory Reconciliation</h1>
    <p class="subtitle">Month-end QBO vs TicketVault reconciliation</p>

    <!-- Checks accordion -->
    <div class="checks-toggle" id="checks-toggle">
      <svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <polyline points="6 3 11 8 6 13"/>
      </svg>
      What does this check?
    </div>
    <div class="checks-body" id="checks-body">
      <ol>
        <li><strong>Daily $ Reconciliation</strong> — QBO Bill totals vs TicketVault Total Cost, grouped by date &amp; company</li>
        <li><strong>Duplicate Bills (Bill #s)</strong> — Bills or Expenses where the same Company + Bill # appears more than once in QBO</li>
        <li><strong>Duplicate Bills (Detail)</strong> — Bills or Expenses where Company, Date, Name, Description, and Amount all match</li>
        <li><strong>Description Mismatch</strong> — Company name in parentheses in the Description field must match the QBO Company column (Bills &amp; Expenses)</li>
        <li><strong>TV Date Mismatch</strong> — TicketVault rows where the PO Created date does not match the Created date</li>
      </ol>
    </div>

    <form id="recon-form" enctype="multipart/form-data">

      <!-- QBO files -->
      <div class="field-group">
        <label class="field-label">QBO Daily Inventory Summary Reports</label>
        <div class="upload-zone" id="qbo-zone">
          <input type="file" name="qbo_files" id="qbo-input" multiple accept=".xlsx,.xls">
          <svg class="upload-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
            <polyline points="17 8 12 3 7 8"/>
            <line x1="12" y1="3" x2="12" y2="15"/>
          </svg>
          <div class="zone-hint">
            Drag &amp; drop your <strong>Excel</strong> files here, or <a href="#">browse</a>
            <small>Consolidated and/or single-company QBO reports (.xlsx) — multiple files accepted</small>
          </div>
          <div class="file-list" id="qbo-file-list"></div>
        </div>
      </div>

      <!-- TicketVault files -->
      <div class="field-group">
        <label class="field-label">TicketVault Purchase Details Report(s)</label>
        <div class="upload-zone" id="tv-zone">
          <input type="file" name="tv_files" id="tv-input" multiple accept=".xlsx,.xls">
          <svg class="upload-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
            <polyline points="17 8 12 3 7 8"/>
            <line x1="12" y1="3" x2="12" y2="15"/>
          </svg>
          <div class="zone-hint">
            Drag &amp; drop your <strong>Excel</strong> files here, or <a href="#">browse</a>
            <small>One or more Purchase Details exports (.xlsx) — merged automatically</small>
          </div>
          <div class="file-list" id="tv-file-list"></div>
        </div>
      </div>

      <!-- Buttons -->
      <div class="btn-row">
        <button type="submit" class="btn btn-primary" id="run-btn" disabled>
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
            <polyline points="7 10 12 15 17 10"/>
            <line x1="12" y1="15" x2="12" y2="3"/>
          </svg>
          Run &amp; Export
        </button>
        <button type="button" class="btn btn-clear" id="clear-btn">Clear</button>
      </div>
    </form>

    <div class="status-box status-loading" id="status-loading">
      <span class="spinner"></span> Processing files, please wait…
    </div>
    <div class="status-box status-error" id="status-error"></div>
    <div class="status-box status-success" id="status-success"></div>
  </div>

  <script>
    const qboInput  = document.getElementById('qbo-input');
    const tvInput   = document.getElementById('tv-input');
    const runBtn    = document.getElementById('run-btn');
    const clearBtn  = document.getElementById('clear-btn');

    // File icon SVG
    const fileIcon = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>`;

    function renderFileList(input, zone, listEl) {
      listEl.innerHTML = '';
      if (input.files.length === 0) {
        zone.classList.remove('has-files');
      } else {
        zone.classList.add('has-files');
        for (const f of input.files) {
          const div = document.createElement('div');
          div.className = 'file-item';
          div.innerHTML = `${fileIcon}<span>${f.name}</span>`;
          listEl.appendChild(div);
        }
      }
      checkReady();
    }

    function checkReady() {
      runBtn.disabled = !(qboInput.files.length > 0 && tvInput.files.length > 0);
    }

    qboInput.addEventListener('change', () => renderFileList(qboInput, document.getElementById('qbo-zone'), document.getElementById('qbo-file-list')));
    tvInput.addEventListener('change',  () => renderFileList(tvInput,  document.getElementById('tv-zone'),  document.getElementById('tv-file-list')));

    // Drag and drop
    document.querySelectorAll('.upload-zone').forEach(zone => {
      zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('dragover'); });
      zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
      zone.addEventListener('drop', e => {
        e.preventDefault();
        zone.classList.remove('dragover');
        const input = zone.querySelector('input[type=file]');
        const list  = zone.querySelector('.file-list');
        // Merge dropped files with existing (for multi-file QBO zone)
        const dt = new DataTransfer();
        for (const f of input.files) dt.items.add(f);
        for (const f of e.dataTransfer.files) dt.items.add(f);
        input.files = dt.files;
        renderFileList(input, zone, list);
      });
    });

    // Clear button
    clearBtn.addEventListener('click', () => {
      // Reset file inputs
      [qboInput, tvInput].forEach(input => {
        const dt = new DataTransfer();
        input.files = dt.files;
      });
      renderFileList(qboInput, document.getElementById('qbo-zone'), document.getElementById('qbo-file-list'));
      renderFileList(tvInput,  document.getElementById('tv-zone'),  document.getElementById('tv-file-list'));
      ['status-loading','status-error','status-success'].forEach(id => {
        document.getElementById(id).style.display = 'none';
      });
    });

    // Checks accordion
    const toggle = document.getElementById('checks-toggle');
    const body   = document.getElementById('checks-body');
    toggle.addEventListener('click', () => {
      toggle.classList.toggle('open');
      body.classList.toggle('open');
    });

    // Form submit
    document.getElementById('recon-form').addEventListener('submit', async (e) => {
      e.preventDefault();

      document.getElementById('status-loading').style.display = 'block';
      document.getElementById('status-error').style.display   = 'none';
      document.getElementById('status-success').style.display = 'none';
      runBtn.disabled = true;
      runBtn.innerHTML = '<span class="spinner"></span> Processing…';

      const fd = new FormData(e.target);

      try {
        const resp = await fetch('/run', { method: 'POST', body: fd });

        if (!resp.ok) {
          const err = await resp.json();
          throw new Error(err.error || `Server error ${resp.status}`);
        }

        const blob = await resp.blob();
        const url  = URL.createObjectURL(blob);
        const a    = document.createElement('a');
        const ts   = new Date().toISOString().slice(0, 7).replace('-', '_');
        a.href     = url;
        a.download = `Reconciliation_${ts}.xlsx`;
        a.click();
        URL.revokeObjectURL(url);

        document.getElementById('status-loading').style.display = 'none';
        const successBox = document.getElementById('status-success');
        successBox.innerHTML = '✅ Report generated and downloaded successfully!';
        successBox.style.display = 'block';

      } catch (err) {
        document.getElementById('status-loading').style.display = 'none';
        const errBox = document.getElementById('status-error');
        errBox.innerHTML = `❌ <strong>Error:</strong> ${err.message}`;
        errBox.style.display = 'block';
      } finally {
        runBtn.disabled = false;
        runBtn.innerHTML = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg> Run &amp; Export`;
        checkReady();
      }
    });
  </script>
</body>
</html>
"""

# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route('/')
def index():
    return render_template_string(HTML)


@app.route('/run', methods=['POST'])
def run():
    try:
        qbo_files = request.files.getlist('qbo_files')
        tv_files  = request.files.getlist('tv_files')

        if not qbo_files or all(f.filename == '' for f in qbo_files):
            return jsonify(error="No QBO files uploaded."), 400
        if not tv_files or all(f.filename == '' for f in tv_files):
            return jsonify(error="No TicketVault files uploaded."), 400

        # Read all TV file bytes; merge into one via parse_ticketvault
        tv_input_files = []
        tv_buffers = []
        for f in tv_files:
            if f.filename == '':
                continue
            raw = f.read()
            tv_input_files.append((f.filename, raw))
            tv_buffers.append(BytesIO(raw))

        tv_recon_df, tv_raw_df = parse_ticketvault(tv_buffers)

        # Collect input file bytes for embedding in output
        input_files = tv_input_files[:]

        # Parse all QBO files and combine
        qbo_frames = []
        errors = []
        for f in qbo_files:
            if f.filename == '':
                continue
            try:
                raw_bytes = f.read()
                input_files.append((f.filename, raw_bytes))
                raw  = BytesIO(raw_bytes)
                raw2 = BytesIO(raw_bytes)
                try:
                    df = parse_qbo_consolidated(raw)
                    qbo_frames.append(df)
                except Exception:
                    df = parse_qbo_single(raw2, filename=f.filename)
                    qbo_frames.append(df)
            except Exception as ex:
                errors.append(f"{f.filename}: {str(ex)}")

        if not qbo_frames:
            return jsonify(error="Could not parse any QBO files. " + "; ".join(errors)), 400

        qbo_df = pd.concat(qbo_frames, ignore_index=True)

        # Derive period label from QBO date range
        dates = qbo_df['date'].dropna()
        if len(dates):
            period_label = dates.min().strftime('%B %Y')
            if dates.min().month != dates.max().month:
                period_label += ' – ' + dates.max().strftime('%B %Y')
        else:
            period_label = ''

        report_bytes = build_report(
            qbo_df, tv_recon_df, tv_raw_df=tv_raw_df,
            period_label=period_label, input_files=input_files
        )

        safe_label = period_label.replace(' ', '_').replace('–', '-') if period_label else datetime.now().strftime('%Y%m%d')
        filename = f"Inventory_Reconciliation_{safe_label}.xlsx"
        return send_file(
            report_bytes,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as ex:
        traceback.print_exc()
        return jsonify(error=str(ex)), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
