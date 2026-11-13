import * as XLSX from 'xlsx';
import { createWorker } from 'tesseract.js';
import './style.css';

const DB_NAME = 'inventario-local';
const DB_VERSION = 1;
const SESSION_ID = 'current';
const $ = (id) => document.getElementById(id);
const ui = {
  file: $('excelFile'), fileName: $('fileName'), inventory: $('inventoryPanel'), cameraPanel: $('cameraPanel'),
  warehouse: $('warehouse'), itemCount: $('itemCount'), scannedCount: $('scannedCount'), pendingCount: $('pendingCount'),
  camera: $('camera'), placeholder: $('cameraPlaceholder'), cameraToggle: $('cameraToggle'), capture: $('capture'),
  canvas: $('captureCanvas'), ocrStatus: $('ocrStatus'), matchPanel: $('matchPanel'), matchCode: $('matchCode'),
  matchDescription: $('matchDescription'), matchWarehouse: $('matchWarehouse'), register: $('registerScan'),
  pendingList: $('pendingList'), togglePending: $('togglePending'), exportCsv: $('exportCsv'), reset: $('resetSession'),
  saveState: $('saveState'), toast: $('toast')
};
let items = [];
let fileLabel = '';
let cameraStream = null;
let workerPromise = null;
let currentMatch = null;
let toastTimer;

function openDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => req.result.createObjectStore('sessions', { keyPath: 'id' });
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}
async function readSession() {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const req = db.transaction('sessions', 'readonly').objectStore('sessions').get(SESSION_ID);
    req.onsuccess = () => { db.close(); resolve(req.result || null); };
    req.onerror = () => { db.close(); reject(req.error); };
  });
}
async function saveSession() {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction('sessions', 'readwrite');
    tx.objectStore('sessions').put({ id: SESSION_ID, items, fileLabel, updatedAt: Date.now() });
    tx.oncomplete = () => { db.close(); ui.saveState.textContent = 'Guardado en este teléfono'; resolve(); };
    tx.onerror = () => { db.close(); reject(tx.error); };
  });
}
async function clearSession() {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction('sessions', 'readwrite');
    tx.objectStore('sessions').delete(SESSION_ID);
    tx.oncomplete = () => { db.close(); resolve(); };
    tx.onerror = () => { db.close(); reject(tx.error); };
  });
}
function normalize(value) {
  return String(value ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
}
function findColumn(headers, ...names) {
  const wanted = names.map(normalize);
  return headers.findIndex((h) => wanted.includes(normalize(h)));
}
function readQty(v) {
  if (typeof v === 'number') return v;
  const n = Number(String(v ?? '').trim().replace(',', '.'));
  return Number.isFinite(n) ? n : 0;
}
function formatQty(n) {
  return Number.isInteger(n) ? String(n) : n.toLocaleString('es-CO', { maximumFractionDigits: 2 });
}
function makeKey(code, warehouse) { return `${code}__${warehouse}`; }

ui.file.addEventListener('change', async () => {
  const file = ui.file.files?.[0];
  if (!file) return;
  ui.saveState.textContent = 'Leyendo Excel…';
  try {
    const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', raw: false });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
    const headers = rows[0] || [];
    const codeCol = findColumn(headers, 'sabcodiitem');
    const descCol = findColumn(headers, 'itedesccort', 'itesccort');
    const qtyCol = findColumn(headers, 'sabcantmanoump');
    const warehouseCol = findColumn(headers, 'sabbodega');
    if (codeCol < 0 || descCol < 0 || qtyCol < 0) throw new Error('No encontré sabcodiitem, itedesccort/itesccort y sabcantmanoump en la primera fila.');

    const grouped = new Map();
    for (const row of rows.slice(1)) {
      let code = String(row[codeCol] ?? '').trim();
      if (/^\d+\.0+$/.test(code)) code = code.slice(0, code.indexOf('.'));
      if (!code) continue;
      const description = String(row[descCol] ?? '').trim();
      const warehouse = warehouseCol < 0 ? '' : String(row[warehouseCol] ?? '').trim();
      const expected = readQty(row[qtyCol]);
      const key = makeKey(code, warehouse);
      const found = grouped.get(key);
      if (found) found.expected += expected;
      else grouped.set(key, { key, code, description, warehouse, expected, scanned: 0 });
    }
    items = [...grouped.values()];
    fileLabel = file.name;
    render();
    await saveSession();
    showToast(`Cargados ${items.length} productos por bodega`);
  } catch (err) {
    ui.saveState.textContent = 'No se guardó';
    showToast(err.message || 'No pude leer ese archivo Excel.');
  } finally { ui.file.value = ''; }
});

function warehouses() { return [...new Set(items.map((x) => x.warehouse))].sort((a,b) => a.localeCompare(b, 'es')); }
function selectedWarehouse() { return ui.warehouse.value || '*'; }
function visibleItems() { const w = selectedWarehouse(); return items.filter((x) => w === '*' || x.warehouse === w); }
function pendingItems() { return visibleItems().filter((x) => x.scanned < x.expected); }
function render() {
  const loaded = items.length > 0;
  ui.inventory.classList.toggle('hidden', !loaded);
  ui.cameraPanel.classList.toggle('hidden', !loaded);
  ui.exportCsv.disabled = !loaded;
  ui.reset.disabled = !loaded;
  if (!loaded) {
    ui.saveState.textContent = 'Sin Excel'; ui.fileName.textContent = 'Todavía no hay un archivo cargado';
    return;
  }
  const current = ui.warehouse.value;
  const options = [{ id: '*', name: 'Todas las bodegas' }, ...warehouses().map((w) => ({ id: w, name: w || '(sin bodega)' }))];
  ui.warehouse.innerHTML = options.map((o) => `<option value="${escapeHtml(o.id)}">${escapeHtml(o.name)}</option>`).join('');
  if (options.some((o) => o.id === current)) ui.warehouse.value = current;
  ui.fileName.textContent = fileLabel || 'Inventario guardado en este teléfono';
  ui.itemCount.textContent = `${items.length} filas producto/bodega`;
  const selected = visibleItems();
  const scanned = selected.reduce((a, x) => a + x.scanned, 0);
  const pending = pendingItems();
  ui.scannedCount.textContent = formatQty(scanned);
  ui.pendingCount.textContent = pending.length;
  ui.togglePending.textContent = ui.pendingList.classList.contains('hidden') ? 'Ver lista' : 'Ocultar';
  ui.pendingList.innerHTML = pending.length
    ? pending.map((x) => `<article class="pending-item"><div><strong>${escapeHtml(x.code)}</strong><span>${escapeHtml(x.description)}</span><small>${x.warehouse ? `Bodega ${escapeHtml(x.warehouse)} · ` : ''}Escaneadas ${formatQty(x.scanned)} de ${formatQty(x.expected)}</small></div><b>Faltan ${formatQty(Math.max(0, x.expected - x.scanned))}</b></article>`).join('')
    : '<div class="complete-message">✓ No quedan cantidades pendientes en esta selección.</div>';
}
ui.warehouse.addEventListener('change', () => { ui.matchPanel.classList.add('hidden'); currentMatch = null; render(); });
ui.togglePending.addEventListener('click', () => { ui.pendingList.classList.toggle('hidden'); render(); });

async function toggleCamera() {
  if (cameraStream) { stopCamera(); return; }
  if (!navigator.mediaDevices?.getUserMedia) { ui.ocrStatus.textContent = 'Abre esta página en Chrome con HTTPS para habilitar la cámara.'; return; }
  try {
    cameraStream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: 'environment' }, width: { ideal: 1280 }, height: { ideal: 960 } }, audio: false });
    ui.camera.srcObject = cameraStream;
    await ui.camera.play();
    ui.placeholder.classList.add('hidden');
    ui.capture.disabled = false;
    ui.cameraToggle.textContent = 'Cerrar cámara';
    ui.ocrStatus.textContent = 'Alinea la etiqueta y toma una foto.';
  } catch (err) {
    ui.ocrStatus.textContent = err.name === 'NotAllowedError' ? 'Permite el acceso a la cámara en Chrome y vuelve a intentarlo.' : `No se pudo abrir la cámara: ${err.message}`;
  }
}
function stopCamera() {
  cameraStream?.getTracks().forEach((track) => track.stop());
  cameraStream = null; ui.camera.srcObject = null;
  ui.placeholder.classList.remove('hidden'); ui.capture.disabled = true; ui.cameraToggle.textContent = 'Abrir cámara';
}
ui.cameraToggle.addEventListener('click', toggleCamera);

async function getWorker() {
  if (!workerPromise) workerPromise = createWorker('eng', 1, { logger: (m) => {
    if (m.status === 'recognizing text' && m.progress != null) ui.ocrStatus.textContent = `Leyendo etiqueta… ${Math.round(m.progress * 100)}%`;
    if (m.status === 'loading language traineddata') ui.ocrStatus.textContent = 'Preparando lectura por primera vez…';
  } });
  return workerPromise;
}
ui.capture.addEventListener('click', async () => {
  if (!cameraStream || !ui.camera.videoWidth) return;
  ui.capture.disabled = true; ui.matchPanel.classList.add('hidden'); currentMatch = null;
  try {
    const ratio = Math.min(1, 1600 / ui.camera.videoWidth);
    ui.canvas.width = Math.round(ui.camera.videoWidth * ratio);
    ui.canvas.height = Math.round(ui.camera.videoHeight * ratio);
    ui.canvas.getContext('2d', { willReadFrequently: true }).drawImage(ui.camera, 0, 0, ui.canvas.width, ui.canvas.height);
    const worker = await getWorker();
    const { data } = await worker.recognize(ui.canvas);
    const match = findInventoryCode(data.lines || []);
    if (!match) {
      ui.ocrStatus.textContent = 'No encontré un código del Excel sobre Lote/Serial o Concepto. Acerca la etiqueta, evita reflejos y toma otra foto.';
      return;
    }
    const matches = items.filter((x) => x.code === match && (selectedWarehouse() === '*' || x.warehouse === selectedWarehouse()));
    if (!matches.length) { ui.ocrStatus.textContent = `Leí ${match}, pero no está en la bodega seleccionada.`; return; }
    if (matches.length > 1) {
      const choice = prompt(`El código ${match} está en varias bodegas:\n${matches.map((x, i) => `${i + 1}. ${x.warehouse} — ${x.description}`).join('\n')}\n\nEscribe el número de la bodega:`);
      const chosen = matches[Number(choice) - 1];
      if (!chosen) { ui.ocrStatus.textContent = 'Lectura cancelada. Puedes volver a tomar la foto.'; return; }
      currentMatch = chosen;
    } else currentMatch = matches[0];
    ui.matchCode.textContent = currentMatch.code;
    ui.matchDescription.textContent = currentMatch.description;
    ui.matchWarehouse.textContent = currentMatch.warehouse ? `Bodega ${currentMatch.warehouse} · ${formatQty(Math.max(0, currentMatch.expected - currentMatch.scanned))} pendientes` : `${formatQty(Math.max(0, currentMatch.expected - currentMatch.scanned))} pendientes`;
    ui.matchPanel.classList.remove('hidden');
    ui.ocrStatus.textContent = 'Confirma el producto antes de descontar una unidad.';
  } catch (err) {
    ui.ocrStatus.textContent = `No pude procesar la foto: ${err.message}`;
  } finally { ui.capture.disabled = false; }
});

function findInventoryCode(lines) {
  const anchor = lines.find((line) => {
    const n = normalize(line.text);
    return n.includes('loteserial') || n.includes('concepto');
  });
  if (!anchor?.bbox) return null;
  const ay = (anchor.bbox.y0 + anchor.bbox.y1) / 2;
  const ax = (anchor.bbox.x0 + anchor.bbox.x1) / 2;
  const above = lines.filter((line) => line.bbox && (line.bbox.y0 + line.bbox.y1) / 2 < ay)
    .sort((a,b) => ((b.bbox.y0+b.bbox.y1)/2) - ((a.bbox.y0+a.bbox.y1)/2));
  for (const line of above) {
    const cy = (line.bbox.y0 + line.bbox.y1) / 2;
    const cx = (line.bbox.x0 + line.bbox.x1) / 2;
    if (ay - cy > 260) break;
    if (Math.abs(cx - ax) > 420) continue;
    const candidates = (line.text.replace(/[Oo]/g, '0').match(/\d{5,12}/g) || []);
    for (const code of candidates) {
      if (items.some((x) => x.code === code && (selectedWarehouse() === '*' || x.warehouse === selectedWarehouse()))) return code;
    }
  }
  return null;
}

ui.register.addEventListener('click', async () => {
  if (!currentMatch) return;
  if (currentMatch.scanned >= currentMatch.expected) { showToast('La cantidad de este producto ya se completó.'); return; }
  currentMatch.scanned += 1;
  await saveSession();
  ui.matchPanel.classList.add('hidden'); currentMatch = null; render();
  ui.ocrStatus.textContent = 'Descuento registrado. Puedes tomar la siguiente foto.';
  showToast('Unidad registrada');
});

ui.exportCsv.addEventListener('click', () => {
  const rows = [['sabcodiitem','itedesccort','sabbodega','esperadas','escaneadas','faltantes']];
  for (const x of visibleItems()) rows.push([x.code, x.description, x.warehouse, formatQty(x.expected), formatQty(x.scanned), formatQty(Math.max(0, x.expected-x.scanned))]);
  const csv = '\ufeff' + rows.map((r) => r.map((v) => `"${String(v).replace(/"/g, '""')}"`).join(';')).join('\r\n');
  const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
  const a = document.createElement('a'); a.href = url; a.download = 'resultado_inventario.csv'; a.click(); URL.revokeObjectURL(url);
});
ui.reset.addEventListener('click', async () => {
  if (!confirm('¿Borrar el Excel cargado y todos los escaneos guardados en este teléfono?')) return;
  stopCamera(); items = []; fileLabel = ''; currentMatch = null; await clearSession();
  ui.matchPanel.classList.add('hidden'); ui.pendingList.classList.add('hidden'); render(); showToast('Sesión borrada');
});

function escapeHtml(value) { return String(value).replace(/[&<>"']/g, (c) => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[c])); }
function showToast(message) { ui.toast.textContent = message; ui.toast.classList.add('show'); clearTimeout(toastTimer); toastTimer = setTimeout(() => ui.toast.classList.remove('show'), 2800); }
window.addEventListener('beforeunload', stopCamera);

async function restoreSession() {
  try {
    const saved = await readSession();
    if (saved?.items?.length) { items = saved.items; fileLabel = saved.fileLabel || ''; render(); ui.saveState.textContent = 'Sesión recuperada'; }
  } catch { ui.saveState.textContent = 'No disponible el guardado local'; }
}
restoreSession();
