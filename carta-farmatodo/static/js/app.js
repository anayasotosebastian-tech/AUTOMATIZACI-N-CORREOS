// Estado global
let currentSessionId = null;
let currentResults = [];
let pendingLabId = null;
let pendingLabName = null;

// Fecha de hoy al cargar
document.addEventListener('DOMContentLoaded', () => {
  const today = new Date().toISOString().split('T')[0];
  document.getElementById('doc_date').value = today;
  setupDragDrop();
});

// ---- File Upload ----
function setupDragDrop() {
  const zone = document.getElementById('dropZone');
  const input = document.getElementById('excel_file');

  zone.addEventListener('click', (e) => {
    if (e.target === zone || e.target.closest('#dropMsg')) input.click();
  });

  ['dragenter', 'dragover'].forEach(ev => {
    zone.addEventListener(ev, e => { e.preventDefault(); zone.classList.add('dragover'); });
  });
  ['dragleave', 'drop'].forEach(ev => {
    zone.addEventListener(ev, e => { e.preventDefault(); zone.classList.remove('dragover'); });
  });
  zone.addEventListener('drop', e => {
    const file = e.dataTransfer.files[0];
    if (file) setFile(file, input);
  });
  input.addEventListener('change', () => {
    if (input.files[0]) setFile(input.files[0], input);
  });
}

function setFile(file, input) {
  // Transfer to real input if needed
  const dt = new DataTransfer();
  dt.items.add(file);
  input.files = dt.files;

  document.getElementById('dropMsg').classList.add('d-none');
  document.getElementById('fileSelected').classList.remove('d-none');
  document.getElementById('fileName').textContent = file.name;
  document.getElementById('fileSize').textContent = formatBytes(file.size);
  document.getElementById('dropZone').classList.add('has-file');
}

function formatBytes(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / 1048576).toFixed(1) + ' MB';
}

// ---- Form Submit ----
document.getElementById('uploadForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const fileInput = document.getElementById('excel_file');
  if (!fileInput.files.length) {
    showToast('error', 'Archivo requerido', 'Selecciona un archivo Excel primero.');
    return;
  }

  setLoading(true, 'Leyendo Excel y agrupando laboratorios...');

  const formData = new FormData(e.target);
  try {
    const resp = await fetch('/process', { method: 'POST', body: formData });
    const data = await resp.json();
    if (!resp.ok || data.error) {
      setLoading(false);
      showToast('error', 'Error al procesar', data.error || 'Error desconocido');
      return;
    }
    currentSessionId = data.session_id;
    currentResults = data.results;
    renderResults(data);
    setLoading(false);
    document.getElementById('step2').classList.remove('d-none');
    document.getElementById('step2').scrollIntoView({ behavior: 'smooth' });
    showToast('success', '¡Listo!', `${data.total_labs} cartas generadas para ${data.total_productos} productos.`);
  } catch (err) {
    setLoading(false);
    showToast('error', 'Error de conexión', err.message);
  }
});

function setLoading(on, msg) {
  document.getElementById('processBtn').disabled = on;
  document.getElementById('loadingPanel').classList.toggle('d-none', !on);
  if (msg) document.getElementById('loadingMsg').textContent = msg;
}

// ---- Render Results ----
function renderResults(data) {
  document.getElementById('labCount').textContent = `${data.total_labs} laboratorios`;
  const tbody = document.getElementById('resultsBody');
  tbody.innerHTML = '';

  data.results.forEach(r => {
    const hasPdf = !!r.pdf_file;
    const tr = document.createElement('tr');
    tr.id = `row-${r.lab_id}`;
    tr.innerHTML = `
      <td class="ps-3">
        <div class="fw-semibold">${r.lab_name}</div>
        <div class="text-muted small">${r.laboratorio}</div>
      </td>
      <td>${r.kam}</td>
      <td><span class="text-muted small">${r.email}</span></td>
      <td class="text-center"><span class="badge bg-secondary">${r.productos_count}</span></td>
      <td class="text-center">
        <div class="d-flex gap-1 justify-content-center flex-wrap">
          <a href="/download/${r.session_id}/${encodeURIComponent(r.word_file)}"
             class="btn btn-outline-primary btn-action" title="Descargar Word">
            <i class="bi bi-file-word"></i> .docx
          </a>
          ${hasPdf
            ? `<a href="/download/${r.session_id}/${encodeURIComponent(r.pdf_file)}"
                class="btn btn-outline-danger btn-action" title="Descargar PDF">
                <i class="bi bi-file-pdf"></i> .pdf
               </a>`
            : `<span class="badge bg-warning text-dark" title="LibreOffice no disponible">Sin PDF</span>`
          }
        </div>
      </td>
      <td class="text-center">
        <button class="btn btn-outline-secondary btn-action"
          onclick="openPreview('${r.session_id}', '${r.pdf_file || ''}', '${r.word_file}', '${escHtml(r.lab_name)}')">
          <i class="bi bi-eye"></i> Ver
        </button>
      </td>
      <td class="text-center">
        <button class="btn btn-outline-success btn-action"
          onclick="openEmailModal('${r.lab_id}', '${escHtml(r.lab_name)}', '${r.email}', '${r.session_id}')">
          <i class="bi bi-envelope"></i> Enviar
        </button>
      </td>
      <td class="text-center pe-3">
        <span class="badge bg-light text-secondary status-badge" id="status-${r.lab_id}">Pendiente</span>
      </td>
    `;
    tbody.appendChild(tr);
  });

  // Show warning if any have no PDF
  if (data.results.some(r => !r.pdf_file)) {
    document.getElementById('pdfWarning').style.removeProperty('display');
  }
}

// ---- Preview ----
function openPreview(sessionId, pdfFile, wordFile, labName) {
  document.getElementById('previewTitle').textContent = labName;
  const frame = document.getElementById('previewFrame');
  const noPdf = document.getElementById('previewNoPdf');
  const dlBtn = document.getElementById('previewDownload');

  if (pdfFile) {
    frame.src = `/preview/${sessionId}/${encodeURIComponent(pdfFile)}`;
    frame.classList.remove('d-none');
    noPdf.classList.add('d-none');
    dlBtn.href = `/download/${sessionId}/${encodeURIComponent(pdfFile)}`;
    dlBtn.download = pdfFile;
  } else {
    frame.src = '';
    frame.classList.add('d-none');
    noPdf.classList.remove('d-none');
    dlBtn.href = `/download/${sessionId}/${encodeURIComponent(wordFile)}`;
    dlBtn.download = wordFile;
  }
  new bootstrap.Modal(document.getElementById('previewModal')).show();
}

// ---- Email individual ----
function openEmailModal(labId, labName, email, sessionId) {
  pendingLabId = labId;
  document.getElementById('emailTo').textContent = email;
  document.getElementById('emailCc').textContent = 'sebastian_anaya@nadro.com.mx, josep.jaramillo@rfp.mx';
  document.getElementById('emailSubject').textContent =
    `SOLICITUD DE RENOVACIÓN PLAN DE LEALTAD - ${labName} ${getDocDate()}`;
  new bootstrap.Modal(document.getElementById('emailModal')).show();
}

async function confirmSendEmail() {
  if (!pendingLabId) return;
  const modal = bootstrap.Modal.getInstance(document.getElementById('emailModal'));

  const btn = document.getElementById('confirmSendBtn');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner-border spinner-border-sm me-1"></span>Enviando...';

  try {
    const resp = await fetch('/send_email', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        session_id: currentSessionId,
        lab_id: pendingLabId,
        smtp_config: getSmtpConfig()
      })
    });
    const data = await resp.json();
    modal.hide();
    btn.disabled = false;
    btn.innerHTML = '<i class="bi bi-send me-1"></i>Enviar';

    if (data.success) {
      markSent(pendingLabId);
      showToast('success', 'Correo enviado', `Correo enviado correctamente.`);
    } else {
      showToast('error', 'Error al enviar', data.error || 'Error desconocido');
    }
  } catch (err) {
    btn.disabled = false;
    btn.innerHTML = '<i class="bi bi-send me-1"></i>Enviar';
    showToast('error', 'Error', err.message);
  }
}

// ---- Enviar todos ----
async function sendAllEmails() {
  const btn = document.getElementById('sendAllBtn');
  const cancelBtn = document.getElementById('sendAllCancel');
  const progress = document.getElementById('sendAllProgress');
  const results = document.getElementById('sendAllResults');

  btn.disabled = true;
  cancelBtn.disabled = true;
  progress.classList.remove('d-none');
  results.classList.add('d-none');
  document.getElementById('sendProgressBar').style.width = '20%';
  document.getElementById('sendProgressMsg').textContent = 'Enviando correos...';

  try {
    const resp = await fetch('/send_all_emails', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        session_id: currentSessionId,
        smtp_config: getSmtpConfig()
      })
    });
    const data = await resp.json();

    document.getElementById('sendProgressBar').style.width = '100%';
    document.getElementById('sendProgressMsg').textContent = 'Completado';
    results.classList.remove('d-none');

    let html = '';
    if (data.sent && data.sent.length) {
      html += `<div class="alert alert-success p-2 mb-2">
        <strong><i class="bi bi-check-circle me-1"></i>Enviados (${data.sent.length}):</strong>
        <ul class="mb-0 mt-1">${data.sent.map(l => `<li>${l}</li>`).join('')}</ul>
      </div>`;
      data.sent.forEach(labName => {
        const r = currentResults.find(x => x.lab_name === labName);
        if (r) markSent(r.lab_id);
      });
    }
    if (data.errors && data.errors.length) {
      html += `<div class="alert alert-danger p-2 mb-0">
        <strong><i class="bi bi-x-circle me-1"></i>Errores (${data.errors.length}):</strong>
        <ul class="mb-0 mt-1">${data.errors.map(e => `<li>${e.lab}: ${e.error}</li>`).join('')}</ul>
      </div>`;
    }
    document.getElementById('sendSuccessList').innerHTML = html;

    btn.disabled = false;
    cancelBtn.disabled = false;
    cancelBtn.textContent = 'Cerrar';
    showToast(data.success ? 'success' : 'warning',
      data.success ? 'Correos enviados' : 'Enviado con errores',
      `${(data.sent || []).length} enviados, ${(data.errors || []).length} errores`);
  } catch (err) {
    btn.disabled = false;
    cancelBtn.disabled = false;
    showToast('error', 'Error', err.message);
  }
}

// ---- Download All ----
function downloadAll() {
  if (!currentSessionId) return;
  window.location.href = `/download_all/${currentSessionId}`;
}

// ---- Helpers ----
function getSmtpConfig() {
  return {
    host: document.getElementById('smtp_host').value,
    port: document.getElementById('smtp_port').value,
    from_email: document.getElementById('smtp_from').value,
    username: document.getElementById('smtp_user').value,
    password: document.getElementById('smtp_pass').value,
  };
}

function getDocDate() {
  const val = document.getElementById('doc_date').value;
  const d = new Date(val + 'T12:00:00');
  return d.toLocaleDateString('es-MX', { day: 'numeric', month: 'long', year: 'numeric' });
}

function markSent(labId) {
  const row = document.getElementById(`row-${labId}`);
  if (row) row.classList.add('row-sent');
  const badge = document.getElementById(`status-${labId}`);
  if (badge) {
    badge.className = 'badge bg-success status-badge';
    badge.textContent = 'Enviado ✓';
  }
}

function resetForm() {
  document.getElementById('uploadForm').reset();
  document.getElementById('dropMsg').classList.remove('d-none');
  document.getElementById('fileSelected').classList.add('d-none');
  document.getElementById('dropZone').classList.remove('has-file');
  document.getElementById('step2').classList.add('d-none');
  const today = new Date().toISOString().split('T')[0];
  document.getElementById('doc_date').value = today;
  currentSessionId = null;
  currentResults = [];
}

function togglePass() {
  const inp = document.getElementById('smtp_pass');
  const icon = document.getElementById('eyeIcon');
  if (inp.type === 'password') {
    inp.type = 'text';
    icon.className = 'bi bi-eye-slash';
  } else {
    inp.type = 'password';
    icon.className = 'bi bi-eye';
  }
}

function escHtml(s) {
  return String(s).replace(/'/g, "\\'").replace(/"/g, '&quot;');
}

function showToast(type, title, msg) {
  const toastEl = document.getElementById('liveToast');
  const icon = document.getElementById('toastIcon');
  const icons = { success: 'bi-check-circle-fill text-success', error: 'bi-x-circle-fill text-danger', warning: 'bi-exclamation-triangle-fill text-warning' };
  icon.className = `bi ${icons[type] || icons.success} me-2`;
  document.getElementById('toastTitle').textContent = title;
  document.getElementById('toastMsg').textContent = msg;
  new bootstrap.Toast(toastEl, { delay: 5000 }).show();
}
