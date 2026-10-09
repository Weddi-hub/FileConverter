"""
PDF to Excel Web Application
============================
"""

import os
import uuid
import logging
from flask import Flask, request, jsonify, send_file, render_template_string
from werkzeug.utils import secure_filename
import pandas as pd

from pdf_to_excel import convert_pdf_to_excel

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB limit

UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "outputs")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("pdf_web_app")

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PDF to Excel Converter</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css">
  <style>
    :root {
      --primary-color: #1a56db;
      --excel-color: #107c41;
      --bg-light: #f8fafc;
    }
    body {
      background-color: var(--bg-light);
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
      color: #1e293b;
    }
    .header-bar {
      background: linear-gradient(135deg, #1e3a8a 0%, #1e40af 50%, #1d4ed8 100%);
      color: white;
      padding: 2rem 0;
      margin-bottom: 2rem;
      box-shadow: 0 4px 20px rgba(30, 58, 138, 0.15);
    }
    .card-custom {
      border: 1px solid #e2e8f0;
      border-radius: 12px;
      box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05), 0 2px 4px -2px rgba(0,0,0,0.05);
      background: #ffffff;
      margin-bottom: 1.5rem;
    }
    .dropzone {
      border: 2px dashed #94a3b8;
      border-radius: 12px;
      padding: 3rem 1.5rem;
      text-align: center;
      cursor: pointer;
      background: #f8fafc;
      transition: all 0.2s ease-in-out;
    }
    .dropzone:hover, .dropzone.dragover {
      border-color: var(--primary-color);
      background-color: #eff6ff;
    }
    .badge-excel {
      background-color: #e6f4ea;
      color: var(--excel-color);
      font-weight: 600;
      border: 1px solid #b7e1cd;
      padding: 0.4rem 0.8rem;
      border-radius: 6px;
    }
    .preview-table-container {
      max-height: 520px;
      overflow: auto;
      border: 1px solid #e2e8f0;
      border-radius: 8px;
    }
    .preview-table th {
      position: sticky;
      top: 0;
      background: #1e3a8a;
      color: white;
      font-weight: 600;
      z-index: 2;
      font-size: 0.85rem;
      padding: 0.6rem 0.75rem;
      white-space: nowrap;
    }
    .preview-table td {
      font-size: 0.85rem;
      padding: 0.5rem 0.75rem;
      white-space: nowrap;
    }
    .btn-convert {
      background-color: var(--excel-color);
      color: white;
      font-weight: 600;
      font-size: 1.05rem;
      padding: 0.85rem 2rem;
      border: none;
      border-radius: 8px;
      transition: all 0.2s;
    }
    .btn-convert:hover {
      background-color: #0c6133;
      color: white;
    }
    .stat-badge {
      display: inline-flex;
      align-items: center;
      gap: 0.4rem;
      background: #f1f5f9;
      padding: 0.45rem 0.85rem;
      border-radius: 8px;
      font-size: 0.88rem;
      font-weight: 500;
    }
  </style>
</head>
<body>

  <!-- Header -->
  <header class="header-bar">
    <div class="container">
      <div class="d-flex align-items-center justify-content-between flex-wrap gap-3">
        <div>
          <h1 class="h3 mb-1 fw-bold"><i class="bi bi-file-earmark-spreadsheet me-2"></i>PDF to Excel Converter</h1>
          <p class="mb-0 text-blue-100 opacity-75">Upload your PDF to extract all tables into a <strong>single continuous Excel sheet ('Sheet1')</strong>.</p>
        </div>
        <div>
          <span class="badge badge-excel"><i class="bi bi-check-circle-fill me-1"></i> Continuous Sheet1</span>
        </div>
      </div>
    </div>
  </header>

  <div class="container pb-5">
    <div class="row">
      <!-- Upload Column -->
      <div class="col-lg-5">
        <div class="card card-custom p-4">
          <h5 class="fw-bold mb-3"><i class="bi bi-cloud-arrow-up text-primary me-2"></i>Upload PDF File</h5>
          
          <!-- Dropzone -->
          <div class="dropzone mb-3" id="dropzone" onclick="document.getElementById('fileInput').click()">
            <i class="bi bi-file-earmark-pdf fs-1 text-danger"></i>
            <h5 class="mt-2 mb-1 fw-bold" id="dropzoneText">Click or Drag & Drop PDF here</h5>
            <p class="text-muted small mb-0">Select any multi-page PDF document</p>
            <input type="file" id="fileInput" accept=".pdf" style="display:none" onchange="handleFileSelect(this.files)">
          </div>

          <!-- Selected File Details -->
          <div id="fileSelectedInfo" class="alert alert-info py-2 px-3 d-none mb-3">
            <div class="d-flex align-items-center justify-content-between">
              <div class="text-truncate me-2">
                <i class="bi bi-file-earmark-check-fill text-success me-1"></i>
                <span id="fileName" class="fw-semibold">No file selected</span>
                <span id="fileSize" class="text-muted ms-2 small"></span>
              </div>
              <button class="btn btn-sm btn-outline-secondary py-0" onclick="clearSelectedFile()">Clear</button>
            </div>
          </div>

          <!-- Convert Button -->
          <button class="btn btn-convert w-100" id="convertBtn" onclick="runConversion()">
            <span id="convertBtnSpinner" class="spinner-border spinner-border-sm me-2 d-none" role="status"></span>
            <span id="convertBtnText"><i class="bi bi-file-earmark-excel me-2"></i>Convert to Excel</span>
          </button>

          <!-- Storage Status & Cleanup -->
          <div class="mt-3 p-2 bg-light rounded border text-muted small d-flex align-items-center justify-content-between">
            <div>
              <i class="bi bi-shield-check text-success me-1"></i>
              <span><strong>Auto-cleanup:</strong> Uploads deleted immediately. Outputs kept 15m.</span>
            </div>
            <button type="button" class="btn btn-sm btn-outline-danger py-0 px-2 ms-2 text-nowrap" onclick="clearServerStorage()" title="Manually purge cached files">
              <i class="bi bi-trash3 me-1"></i>Clear
            </button>
          </div>
        </div>
      </div>

      <!-- Preview & Download Column -->
      <div class="col-lg-7">
        <div class="card card-custom p-4">
          <div class="d-flex align-items-center justify-content-between mb-3 flex-wrap gap-2">
            <h5 class="fw-bold mb-0"><i class="bi bi-table text-primary me-2"></i>Excel Preview ('Sheet1')</h5>
            <div id="downloadBtnContainer" class="d-none">
              <a id="downloadBtn" href="#" class="btn btn-success fw-bold px-3">
                <i class="bi bi-download me-1"></i>Download .xlsx
              </a>
            </div>
          </div>

          <!-- Alert area -->
          <div id="alertBox" class="alert alert-secondary mb-3">
            <i class="bi bi-info-circle me-1"></i> Upload a PDF file and click <strong>Convert to Excel</strong>.
          </div>

          <!-- Stats Badges -->
          <div id="statsRow" class="d-flex flex-wrap gap-2 mb-3 d-none">
            <div class="stat-badge"><i class="bi bi-file-earmark-text text-primary"></i> <span id="statPages">0</span> Pages</div>
            <div class="stat-badge"><i class="bi bi-list-ol text-success"></i> <span id="statRows">0</span> Rows</div>
            <div class="stat-badge"><i class="bi bi-columns text-warning"></i> <span id="statCols">0</span> Columns</div>
            <div class="stat-badge"><i class="bi bi-file-spreadsheet text-success"></i> Sheet: <strong>Sheet1</strong></div>
          </div>

          <!-- Table Preview -->
          <div id="previewContainer" class="preview-table-container d-none">
            <table class="table table-sm table-striped table-hover mb-0 preview-table" id="previewTable">
              <thead id="previewThead"></thead>
              <tbody id="previewTbody"></tbody>
            </table>
          </div>

          <!-- Empty State Graphic -->
          <div id="emptyState" class="text-center py-5 text-muted">
            <i class="bi bi-file-spreadsheet fs-1 opacity-50"></i>
            <p class="mt-2 mb-0">Extracted table data will appear here once converted.</p>
          </div>
        </div>
      </div>
    </div>
  </div>

  <script>
    let selectedFile = null;

    const dropzone = document.getElementById('dropzone');
    const fileInput = document.getElementById('fileInput');

    ['dragenter', 'dragover'].forEach(name => {
      dropzone.addEventListener(name, (e) => {
        e.preventDefault();
        dropzone.classList.add('dragover');
      });
    });

    ['dragleave', 'drop'].forEach(name => {
      dropzone.addEventListener(name, (e) => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
      });
    });

    dropzone.addEventListener('drop', (e) => {
      if (e.dataTransfer.files.length) {
        handleFileSelect(e.dataTransfer.files);
      }
    });

    function handleFileSelect(files) {
      if (!files.length) return;
      selectedFile = files[0];
      document.getElementById('fileName').textContent = selectedFile.name;
      document.getElementById('fileSize').textContent = `(${(selectedFile.size / 1024).toFixed(1)} KB)`;
      document.getElementById('fileSelectedInfo').classList.remove('d-none');
      document.getElementById('dropzoneText').textContent = selectedFile.name;
    }

    function clearSelectedFile() {
      selectedFile = null;
      fileInput.value = '';
      document.getElementById('fileSelectedInfo').classList.add('d-none');
      document.getElementById('dropzoneText').textContent = 'Click or Drag & Drop PDF here';
      document.getElementById('previewContainer').classList.add('d-none');
      document.getElementById('downloadBtnContainer').classList.add('d-none');
      document.getElementById('statsRow').classList.add('d-none');
      document.getElementById('emptyState').classList.remove('d-none');
      showAlert('secondary', '<i class="bi bi-info-circle me-1"></i> Upload a PDF file and click <strong>Convert to Excel</strong>.');
    }

    async function runConversion() {
      if (!selectedFile) {
        showAlert('warning', 'Please select or drop a PDF file first.');
        return;
      }

      setLoading(true);
      showAlert('info', 'Converting PDF... Extracting all pages into continuous Sheet1...');

      const formData = new FormData();
      formData.append('pdf_file', selectedFile);

      try {
        const response = await fetch('/convert', {
          method: 'POST',
          body: formData
        });

        const data = await response.json();
        setLoading(false);

        if (!response.ok || !data.success) {
          showAlert('danger', data.error || 'Conversion failed.');
          return;
        }

        showAlert('success', `<i class="bi bi-check-circle-fill me-1"></i> Successfully converted! Extracted <strong>${data.total_rows} row(s)</strong> into Sheet1.`);
        
        // Show download link
        const downloadBtn = document.getElementById('downloadBtn');
        downloadBtn.href = data.download_url;
        document.getElementById('downloadBtnContainer').classList.remove('d-none');

        // Update stats
        document.getElementById('statPages').textContent = data.total_pages;
        document.getElementById('statRows').textContent = data.total_rows;
        document.getElementById('statCols').textContent = data.total_cols;
        document.getElementById('statsRow').classList.remove('d-none');

        // Render preview table
        renderTablePreview(data.columns, data.preview_rows);

      } catch (err) {
        setLoading(false);
        showAlert('danger', 'Network or server error: ' + err.message);
      }
    }

    function renderTablePreview(columns, rows) {
      const thead = document.getElementById('previewThead');
      const tbody = document.getElementById('previewTbody');
      thead.innerHTML = '';
      tbody.innerHTML = '';

      if (!columns || !columns.length) {
        document.getElementById('emptyState').classList.remove('d-none');
        document.getElementById('previewContainer').classList.add('d-none');
        return;
      }

      // Headers
      const trHead = document.createElement('tr');
      columns.forEach(col => {
        const th = document.createElement('th');
        th.textContent = col;
        trHead.appendChild(th);
      });
      thead.appendChild(trHead);

      // Rows
      rows.forEach(row => {
        const tr = document.createElement('tr');
        columns.forEach(col => {
          const td = document.createElement('td');
          td.textContent = row[col] !== undefined && row[col] !== null ? row[col] : '';
          tr.appendChild(td);
        });
        tbody.appendChild(tr);
      });

      document.getElementById('emptyState').classList.add('d-none');
      document.getElementById('previewContainer').classList.remove('d-none');
    }

    function showAlert(type, message) {
      const box = document.getElementById('alertBox');
      box.className = `alert alert-${type} mb-3`;
      box.innerHTML = message;
    }

    function setLoading(isLoading) {
      const btn = document.getElementById('convertBtn');
      const spinner = document.getElementById('convertBtnSpinner');
      const btnText = document.getElementById('convertBtnText');
      if (isLoading) {
        btn.disabled = true;
        spinner.classList.remove('d-none');
        btnText.textContent = ' Converting...';
      } else {
        btn.disabled = false;
        spinner.classList.add('d-none');
        btnText.innerHTML = '<i class="bi bi-file-earmark-excel me-2"></i>Convert to Excel';
      }
    }

    async function clearServerStorage() {
      if (!confirm("Clear all temporary converted files from the server?")) return;
      try {
        const res = await fetch('/clear-storage', { method: 'POST' });
        const data = await res.json();
        showAlert('info', `<i class="bi bi-check2-circle me-1"></i> Storage cache cleared (${data.deleted_files} temporary files removed).`);
        clearSelectedFile();
      } catch (e) {
        showAlert('danger', 'Failed to clear server storage: ' + e.message);
      }
    }
  </script>
</body>
</html>
"""

import time
import threading

def purge_expired_files(max_age_seconds: int = 900) -> int:
    """
    Deletes files in uploads/ and outputs/ older than max_age_seconds (default: 15 minutes).
    Pass max_age_seconds=0 to clear all temporary files immediately.
    """
    now = time.time()
    deleted = 0
    for directory in [UPLOAD_DIR, OUTPUT_DIR]:
        if not os.path.exists(directory):
            continue
        for fname in os.listdir(directory):
            if fname.startswith("."):
                continue
            fpath = os.path.join(directory, fname)
            if os.path.isfile(fpath):
                try:
                    mtime = os.path.getmtime(fpath)
                    if now - mtime >= max_age_seconds:
                        os.remove(fpath)
                        deleted += 1
                        logger.info(f"Purged expired file: {fname}")
                except Exception as e:
                    logger.warning(f"Could not purge '{fpath}': {e}")
    return deleted


def start_storage_janitor():
    """Starts background janitor daemon that sweeps uploads/ and outputs/ every 5 minutes."""
    def janitor_loop():
        while True:
            time.sleep(300)
            try:
                purge_expired_files(max_age_seconds=900)
            except Exception as e:
                logger.error(f"Janitor error: {e}")

    janitor_thread = threading.Thread(target=janitor_loop, daemon=True)
    janitor_thread.start()
    logger.info("Storage Janitor thread active (15-min retention window).")


@app.route("/")
def index():
    """Serves the clean converter interface."""
    return render_template_string(HTML_TEMPLATE)


@app.route("/clear-storage", methods=["POST"])
def clear_storage():
    """Manually flushes temporary files from uploads and outputs."""
    deleted_count = purge_expired_files(max_age_seconds=0)
    return jsonify({
        "success": True,
        "deleted_files": deleted_count,
        "message": f"Purged {deleted_count} temporary files."
    })


@app.route("/convert", methods=["POST"])
def convert():
    """Converts uploaded PDF and returns preview data + download link."""
    # Periodic sweep on request
    purge_expired_files(max_age_seconds=900)

    if 'pdf_file' not in request.files:
        return jsonify({"success": False, "error": "No PDF file uploaded."}), 400

    uploaded_file = request.files['pdf_file']
    if uploaded_file.filename == '':
        return jsonify({"success": False, "error": "No file selected."}), 400

    original_filename = secure_filename(uploaded_file.filename)
    if not original_filename.lower().endswith('.pdf'):
        original_filename += '.pdf'

    unique_prefix = uuid.uuid4().hex[:8]
    input_pdf_path = os.path.join(UPLOAD_DIR, f"{unique_prefix}_{original_filename}")
    uploaded_file.save(input_pdf_path)

    # Output file path
    base_name = os.path.splitext(original_filename)[0]
    out_id = uuid.uuid4().hex[:8]
    output_filename = f"{base_name}_{out_id}.xlsx"
    output_excel_path = os.path.join(OUTPUT_DIR, output_filename)

    total_pages = 0
    try:
        import pdfplumber
        with pdfplumber.open(input_pdf_path) as pdf:
            total_pages = len(pdf.pages)
    except Exception as e:
        logger.warning(f"Could not read page count: {e}")

    try:
        df, out_path = convert_pdf_to_excel(
            pdf_path=input_pdf_path,
            output_excel_path=output_excel_path,
            sheet_name="Sheet1",
            mode="auto",
            drop_duplicate_headers=True,
            include_structured_text=False
        )

        # Build preview records (first 100 rows for preview)
        preview_rows = df.head(100).to_dict(orient="records")
        columns = list(df.columns)

        return jsonify({
            "success": True,
            "download_url": f"/download/{output_filename}",
            "filename": output_filename,
            "total_pages": total_pages,
            "total_rows": len(df),
            "total_cols": len(columns),
            "columns": columns,
            "preview_rows": preview_rows
        })

    except Exception as e:
        logger.error(f"Conversion error: {e}", exc_info=True)
        return jsonify({"success": False, "error": str(e)}), 500

    finally:
        # Ephemeral storage rule: immediately delete uploaded temporary input PDF
        if os.path.exists(input_pdf_path):
            try:
                os.remove(input_pdf_path)
                logger.info(f"Ephemeral purge: Removed upload '{input_pdf_path}'")
            except Exception as err:
                logger.warning(f"Could not remove temp upload '{input_pdf_path}': {err}")


@app.route("/download/<path:filename>")
def download(filename):
    """Provides download for the generated Excel file."""
    file_path = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(file_path):
        return "File not found or expired.", 404
    return send_file(
        file_path,
        as_attachment=True,
        download_name=filename.split("_")[-1] if "_" not in filename else filename
    )


@app.route("/health")
def health():
    return jsonify({"status": "healthy", "service": "pdf-to-excel-converter"})


if __name__ == "__main__":
    logger.info("Purging stale temporary files on startup...")
    purge_expired_files(max_age_seconds=0)
    start_storage_janitor()
    logger.info("Starting PDF to Excel Converter Web Server on http://127.0.0.1:5000 ...")
    app.run(host="127.0.0.1", port=5000, debug=False)
