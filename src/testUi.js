import { randomBytes } from 'node:crypto';

const OUTPUT_OPTIONS = [
  ['PDF', 'PDF', 'Fixed layout with selectable text'],
  ['DOCX_EDITABLE', 'Word · Editable', 'Native tables and editable content'],
  ['DOCX_VISUAL', 'Word · Visual', 'Exact PDF pages as images'],
  ['XLSX', 'Excel', 'Native cells and workbook formatting'],
];

export function testUiPage() {
  const nonce = randomBytes(18).toString('base64');
  const options = OUTPUT_OPTIONS.map(([value, label, note]) => (
    `<option value="${value}">${label} — ${note}</option>`
  )).join('');
  const html = `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="color-scheme" content="light">
  <title>RDL Render Lab</title>
  <style nonce="${nonce}">
    :root {
      color-scheme: light; --ink:#17213d; --muted:#65708a; --line:#dce2ed; --soft:#f4f6fb;
      --surface:#fff; --navy:#173a68; --blue:#2f6fed; --blue-dark:#2058c4; --green:#168562;
      --red:#b43a47; --shadow:0 24px 60px rgba(31,47,82,.12);
    }
    * { box-sizing:border-box; }
    body {
      margin:0; min-height:100vh; font-family:Arial,Helvetica,sans-serif; color:var(--ink);
      background:radial-gradient(circle at 8% 10%,rgba(47,111,237,.13),transparent 29rem),
        linear-gradient(180deg,#eef3fb 0,#f8f9fc 42%,#f2f4f8 100%);
    }
    button,input,select { font:inherit; }
    .shell { width:min(1080px,calc(100% - 32px)); margin:0 auto; padding:48px 0 64px; }
    .topbar { display:flex; align-items:center; justify-content:space-between; gap:24px; margin-bottom:30px; }
    .brand { display:flex; align-items:center; gap:14px; }
    .mark {
      width:42px; height:42px; display:grid; place-items:center; border-radius:13px; color:#fff;
      background:linear-gradient(145deg,var(--navy),var(--blue)); box-shadow:0 8px 20px rgba(37,83,158,.22);
      font-weight:800; letter-spacing:-.05em;
    }
    .brand-copy strong { display:block; font-size:15px; letter-spacing:.01em; }
    .brand-copy span { color:var(--muted); font-size:12px; }
    .health {
      display:inline-flex; align-items:center; gap:8px; min-height:34px; padding:7px 12px;
      border:1px solid var(--line); border-radius:999px; background:rgba(255,255,255,.75);
      color:var(--muted); font-size:12px; font-weight:700;
    }
    .health-dot { width:8px; height:8px; border-radius:50%; background:#a3aabc; }
    .health[data-state="ready"] .health-dot { background:var(--green); box-shadow:0 0 0 4px rgba(22,133,98,.1); }
    .health[data-state="unavailable"] .health-dot { background:var(--red); }
    .hero { max-width:750px; margin-bottom:30px; }
    .eyebrow { margin:0 0 12px; color:var(--blue); font-size:12px; font-weight:800; letter-spacing:.13em; text-transform:uppercase; }
    h1 { margin:0; font-size:clamp(32px,5vw,54px); line-height:1.03; letter-spacing:-.045em; }
    .intro { margin:17px 0 0; max-width:660px; color:var(--muted); font-size:16px; line-height:1.65; }
    .card {
      overflow:hidden; border:1px solid rgba(211,218,231,.9); border-radius:22px;
      background:rgba(255,255,255,.94); box-shadow:var(--shadow);
    }
    .form-body { padding:30px; }
    .file-grid { display:grid; grid-template-columns:1fr 1fr; gap:18px; }
    .file-field {
      position:relative; display:flex; min-height:180px; padding:22px; cursor:pointer; flex-direction:column;
      justify-content:space-between; border:1.5px dashed #b9c4d8; border-radius:16px; background:var(--soft);
      transition:border-color .16s,background .16s,transform .16s;
    }
    .file-field:hover,.file-field:focus-within,.file-field[data-dragging="true"] {
      border-color:var(--blue); background:#f0f5ff; transform:translateY(-1px);
    }
    .file-field[data-selected="true"] { border-style:solid; border-color:#8fa9dc; background:#f5f8ff; }
    .file-field input { position:absolute; width:1px; height:1px; opacity:0; pointer-events:none; }
    .file-heading { display:flex; align-items:flex-start; justify-content:space-between; gap:16px; }
    .file-icon {
      width:38px; height:38px; display:grid; place-items:center; border-radius:11px; color:var(--navy);
      background:#fff; border:1px solid var(--line); font-size:11px; font-weight:800;
    }
    .file-number { color:#97a1b5; font-size:12px; font-weight:800; letter-spacing:.08em; }
    .file-title { display:block; margin-top:22px; font-weight:800; }
    .file-note { display:block; margin-top:5px; color:var(--muted); font-size:12px; line-height:1.45; }
    .file-name {
      display:block; overflow:hidden; margin-top:14px; color:var(--blue-dark); font-size:12px; font-weight:700;
      text-overflow:ellipsis; white-space:nowrap;
    }
    .controls { display:grid; grid-template-columns:minmax(0,1fr) auto; align-items:end; gap:18px; margin-top:22px; }
    .control-label { display:block; margin:0 0 9px; font-size:12px; font-weight:800; letter-spacing:.02em; }
    select {
      width:100%; height:49px; padding:0 42px 0 14px; color:var(--ink); border:1px solid var(--line);
      border-radius:12px; background:#fff;
    }
    select:focus,button:focus-visible,.file-field:focus-within { outline:3px solid rgba(47,111,237,.2); outline-offset:2px; }
    .render-button {
      min-width:190px; height:49px; padding:0 22px; border:0; border-radius:12px; color:#fff;
      background:var(--blue); box-shadow:0 9px 20px rgba(47,111,237,.2); cursor:pointer; font-weight:800;
      transition:background .16s,transform .16s,opacity .16s;
    }
    .render-button:hover:not(:disabled) { background:var(--blue-dark); transform:translateY(-1px); }
    .render-button:disabled { cursor:wait; opacity:.62; box-shadow:none; }
    .status { min-height:22px; margin:18px 0 0; color:var(--muted); font-size:13px; line-height:1.5; }
    .status[data-kind="error"] { color:var(--red); }
    .status[data-kind="success"] { color:var(--green); }
    .privacy {
      display:flex; gap:12px; padding:18px 30px; border-top:1px solid var(--line); color:var(--muted);
      background:#fafbfe; font-size:12px; line-height:1.5;
    }
    .privacy strong { color:var(--ink); }
    .shield { flex:0 0 auto; width:18px; height:18px; border:2px solid #8090aa; border-radius:5px 5px 8px 8px; }
    @media (max-width:720px) {
      .shell { width:min(100% - 22px,1080px); padding-top:24px; }
      .topbar { align-items:flex-start; } .brand-copy span { display:none; }
      .file-grid,.controls { grid-template-columns:1fr; } .form-body { padding:20px; }
      .file-field { min-height:150px; } .render-button { width:100%; } .privacy { padding:16px 20px; }
    }
  </style>
</head>
<body>
  <main class="shell">
    <header class="topbar">
      <div class="brand">
        <div class="mark" aria-hidden="true">R</div>
        <div class="brand-copy"><strong>RDL Converter</strong><span>End-to-end render workspace</span></div>
      </div>
      <div class="health" id="health" data-state="checking" role="status">
        <span class="health-dot" aria-hidden="true"></span><span id="health-label">Checking service</span>
      </div>
    </header>
    <section class="hero" aria-labelledby="page-title">
      <p class="eyebrow">Open test utility</p>
      <h1 id="page-title">Turn report definitions into usable files.</h1>
      <p class="intro">Upload an RDL and its matching request JSON, choose an output, and download the rendered report. No account or role is required.</p>
    </section>
    <form class="card" id="render-form">
      <div class="form-body">
        <div class="file-grid">
          <label class="file-field" id="rdl-drop">
            <input id="rdl-file" name="rdl" type="file" accept=".rdl,.xml,application/xml,text/xml" required>
            <span class="file-heading"><span class="file-icon" aria-hidden="true">RDL</span><span class="file-number">01</span></span>
            <span>
              <span class="file-title">Report definition</span>
              <span class="file-note">Choose or drop the .rdl file to render.</span>
              <span class="file-name" id="rdl-name">No file selected</span>
            </span>
          </label>
          <label class="file-field" id="json-drop">
            <input id="json-file" name="json" type="file" accept=".json,application/json" required>
            <span class="file-heading"><span class="file-icon" aria-hidden="true">{ }</span><span class="file-number">02</span></span>
            <span>
              <span class="file-title">Request data</span>
              <span class="file-note">Parameters, datasets, and optional bundled subreports.</span>
              <span class="file-name" id="json-name">No file selected</span>
            </span>
          </label>
        </div>
        <div class="controls">
          <div>
            <label class="control-label" for="output">Output format</label>
            <select id="output" name="output">${options}</select>
          </div>
          <button class="render-button" id="render-button" type="submit">Render &amp; download</button>
        </div>
        <p class="status" id="status" role="status" aria-live="polite">Select both files to begin.</p>
      </div>
      <div class="privacy">
        <span class="shield" aria-hidden="true"></span>
        <span><strong>Direct processing.</strong> Files are sent only to this converter instance. The service does not execute RDL queries or retain report data.</span>
      </div>
    </form>
  </main>
  <script nonce="${nonce}">
    const form = document.getElementById('render-form');
    const rdlInput = document.getElementById('rdl-file');
    const jsonInput = document.getElementById('json-file');
    const outputInput = document.getElementById('output');
    const button = document.getElementById('render-button');
    const status = document.getElementById('status');
    function showStatus(message, kind = '') { status.textContent = message; status.dataset.kind = kind; }
    function displayFile(input, nameId, dropId) {
      const file = input.files && input.files[0];
      document.getElementById(nameId).textContent = file ? file.name : 'No file selected';
      document.getElementById(dropId).dataset.selected = String(Boolean(file));
      showStatus(rdlInput.files[0] && jsonInput.files[0] ? 'Ready to render.' : 'Select both files to begin.');
    }
    function installDropZone(input, dropId, nameId) {
      const zone = document.getElementById(dropId);
      input.addEventListener('change', () => displayFile(input, nameId, dropId));
      for (const name of ['dragenter','dragover']) zone.addEventListener(name, (event) => {
        event.preventDefault(); zone.dataset.dragging = 'true';
      });
      for (const name of ['dragleave','drop']) zone.addEventListener(name, (event) => {
        event.preventDefault(); zone.dataset.dragging = 'false';
      });
      zone.addEventListener('drop', (event) => {
        const file = event.dataTransfer.files && event.dataTransfer.files[0];
        if (!file) return;
        const transfer = new DataTransfer();
        transfer.items.add(file);
        input.files = transfer.files;
        displayFile(input, nameId, dropId);
      });
    }
    function downloadName(response, fallback) {
      const disposition = response.headers.get('content-disposition') || '';
      const encoded = disposition.match(/filename\\*=UTF-8''([^;]+)/i);
      if (encoded) { try { return decodeURIComponent(encoded[1]); } catch {} }
      const quoted = disposition.match(/filename="([^"]+)"/i);
      return quoted ? quoted[1] : fallback;
    }
    async function errorMessage(response) {
      try {
        const payload = await response.json();
        const code = payload && payload.error && payload.error.code;
        const message = payload && payload.error && payload.error.message;
        return [code,message].filter(Boolean).join(': ') || 'The render request failed.';
      } catch { return 'The render request failed.'; }
    }
    installDropZone(rdlInput, 'rdl-drop', 'rdl-name');
    installDropZone(jsonInput, 'json-drop', 'json-name');
    fetch('/readyz', { headers:{ accept:'application/json' } })
      .then((response) => {
        document.getElementById('health').dataset.state = response.ok ? 'ready' : 'unavailable';
        document.getElementById('health-label').textContent = response.ok ? 'Service ready' : 'Service not ready';
      })
      .catch(() => {
        document.getElementById('health').dataset.state = 'unavailable';
        document.getElementById('health-label').textContent = 'Service unavailable';
      });
    form.addEventListener('submit', async (event) => {
      event.preventDefault();
      const rdl = rdlInput.files[0];
      const json = jsonInput.files[0];
      if (!rdl || !json) { showStatus('Choose both an RDL file and a JSON file.','error'); return; }
      let request;
      try {
        request = JSON.parse(await json.text());
        if (!request || typeof request !== 'object' || Array.isArray(request)) throw new Error();
      } catch { showStatus('The request file must contain one valid JSON object.','error'); return; }
      delete request.rdlBase64;
      request.output = outputInput.value;
      if (!request.fileName) request.fileName = rdl.name;
      if (!request.outputFileName) request.outputFileName = rdl.name.replace(/\\.[^.]+$/,'') || 'report';
      const body = new FormData();
      body.append('request', JSON.stringify(request));
      body.append('rdl', rdl, rdl.name);
      button.disabled = true;
      button.textContent = 'Rendering…';
      showStatus('Rendering your report. Large reports can take a moment.');
      try {
        const response = await fetch('/v1/render', { method:'POST', body });
        if (!response.ok) throw new Error(await errorMessage(response));
        const blob = await response.blob();
        const extensions = { PDF:'pdf', DOCX_EDITABLE:'docx', DOCX_VISUAL:'docx', XLSX:'xlsx' };
        const fallback = (request.outputFileName || 'report') + '.' + extensions[request.output];
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = downloadName(response, fallback);
        document.body.appendChild(link);
        link.click();
        const completedName = link.download;
        link.remove();
        URL.revokeObjectURL(url);
        showStatus('Download ready: ' + completedName,'success');
      } catch (error) {
        showStatus(error && error.message ? error.message : 'The render request failed.','error');
      } finally {
        button.disabled = false;
        button.textContent = 'Render & download';
      }
    });
  </script>
</body>
</html>`;
  return {
    html,
    contentSecurityPolicy: [
      "default-src 'none'",
      `script-src 'nonce-${nonce}'`,
      `style-src 'nonce-${nonce}'`,
      "connect-src 'self'",
      "img-src 'self' data:",
      "base-uri 'none'",
      "form-action 'self'",
      "frame-ancestors 'none'",
    ].join('; '),
  };
}
