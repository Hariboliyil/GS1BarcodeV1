/* global XLSX, bwipjs, html2canvas */

const state = {
  workbookName: '',
  master: [],
  stock: [],
  index: [],
  currentProduct: null,
  lastExpiryWarningKey: '',
  form: {
    productCode: '',
    batch: '',
    expiry: '',
    genericCode: '',
    description: '',
    gtin: '',
    locator: '',
    timestamp: '',
  },
};

const el = {
  sourceStatus: document.getElementById('sourceStatus'),
  recordCount: document.getElementById('recordCount'),
  workbookFile: document.getElementById('workbookFile'),
  productCode: document.getElementById('productCode'),
  batch: document.getElementById('batch'),
  expiry: document.getElementById('expiry'),
  applyBtn: document.getElementById('applyBtn'),
  copyBtn: document.getElementById('copyBtn'),
  downloadBtn: document.getElementById('downloadBtn'),
  printBtn: document.getElementById('printBtn'),
  barcodeCanvas: document.getElementById('barcodeCanvas'),
  barcodeTextView: document.getElementById('barcodeTextView'),
  renderStatus: document.getElementById('renderStatus'),
  genericCodeView: document.getElementById('genericCodeView'),
  productCodeView: document.getElementById('productCodeView'),
  descriptionView: document.getElementById('descriptionView'),
  batchView: document.getElementById('batchView'),
  expiryView: document.getElementById('expiryView'),
  gtinView: document.getElementById('gtinView'),
  locatorView: document.getElementById('locatorView'),
  locatorQr: document.getElementById('locatorQr'),
  timestampView: document.getElementById('timestampView'),
  expiryWarning: document.getElementById('expiryWarning'),
  labelPreview: document.getElementById('labelPreview'),
};

function text(value) {
  return String(value ?? '').trim();
}

function optionalText(value) {
  const clean = text(value);
  if (!clean) return '';
  if (/^no reference$/i.test(clean)) return '';
  if (/^n\/a$/i.test(clean)) return '';
  return clean;
}

function normalize(value) {
  return text(value).toUpperCase();
}

function excelSerialToDate(serial) {
  const days = Number(serial);
  if (!Number.isFinite(days) || days <= 0) return null;
  return new Date(Date.UTC(1899, 11, 30) + days * 86400000);
}

function parseDate(value) {
  if (!value) return '';
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
  if (typeof value === 'number') {
    const serial = String(Math.trunc(value));
    if (serial.length === 8) {
      const d = Number(serial.slice(0, 2));
      const m = Number(serial.slice(2, 4)) - 1;
      const y = Number(serial.slice(4, 8));
      return new Date(Date.UTC(y, m, d)).toISOString().slice(0, 10);
    }
    const excelDate = excelSerialToDate(value);
    return excelDate ? excelDate.toISOString().slice(0, 10) : '';
  }
  const s = text(value);
  if (!s) return '';
  if (/^\d{8}$/.test(s)) {
    const d = Number(s.slice(0, 2));
    const m = Number(s.slice(2, 4)) - 1;
    const y = Number(s.slice(4, 8));
    return new Date(Date.UTC(y, m, d)).toISOString().slice(0, 10);
  }
  const parsed = new Date(s);
  return Number.isNaN(parsed.getTime()) ? '' : parsed.toISOString().slice(0, 10);
}

function displayDate(value) {
  const iso = parseDate(value);
  if (!iso) return '';
  return new Intl.DateTimeFormat('en-GB', {
    day: '2-digit',
    month: 'short',
    year: 'numeric',
    timeZone: 'UTC',
  }).format(new Date(`${iso}T00:00:00Z`));
}

function gs1Date(value) {
  const iso = parseDate(value);
  return iso ? `${iso.slice(2, 4)}${iso.slice(5, 7)}${iso.slice(8, 10)}` : '';
}

function getExpiryWarning(expiryValue) {
  const iso = parseDate(expiryValue);
  if (!iso) return '';
  const expiry = new Date(`${iso}T00:00:00`);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const future = new Date(today);
  future.setMonth(future.getMonth() + 3);
  if (expiry < today) return 'Warning: expiry date is already in the past.';
  if (expiry < future) return 'Warning: expiry date is less than 3 months away.';
  return '';
}

function renderExpiryWarning() {
  const warning = getExpiryWarning(state.form.expiry);
  el.expiryWarning.textContent = warning;
  return warning;
}

function timestampValue(date = new Date()) {
  const pad = (n) => String(n).padStart(2, '0');
  return `${pad(date.getDate())}${pad(date.getMonth() + 1)}${date.getFullYear()}${pad(date.getHours())}${pad(date.getMinutes())}${pad(date.getSeconds())}`;
}

function readSheet(workbook, name) {
  const sheet = workbook.Sheets[name];
  if (!sheet) return [];
  return XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false, cellDates: true });
}

function buildIndex() {
  state.index = state.master.map((row) => {
    const productCode = text(row.PRODUCT || row['Product Code'] || row.Product);
    const shortDesc = text(row['ITEM SHORT DESCRIPTION'] || row['Item Short Description']);
    const partNo = text(row['Manufacturer Part Number'] || row.Y);
    const generic = text(row['GENERIC CODE'] || row['Generic Code'] || row.J);
    const gtin = text(row['GTIN Number'] || row.AB);
    return {
      productCode,
      row,
      searchText: normalize([productCode, shortDesc, partNo, generic, gtin].filter(Boolean).join(' ')),
    };
  });
}

function findProduct(productCode) {
  const key = normalize(productCode);
  return state.index.find((entry) => normalize(entry.productCode) === key) || null;
}

function findStocks(productCode) {
  const key = normalize(productCode);
  return state.stock.filter((row) => normalize(row['Product Code']) === key);
}

function deriveProductFields(row, productCode) {
  const stocks = findStocks(productCode);
  const stock = stocks[0] || {};
  const partNo = text(row['Manufacturer Part Number'] || row.Y);
  const shortDesc = text(row['ITEM SHORT DESCRIPTION'] || row['Item Short Description']);
  const longDesc = [partNo, shortDesc].filter(Boolean).join('-');
  const gtin = text(row['GTIN Number'] || row.AB);
  return {
    productCode: text(productCode),
    genericCode: text(row['GENERIC CODE'] || row['Generic Code'] || row.J),
    description: longDesc || shortDesc,
    gtin,
    locator: optionalText(stock['Bin Wise'] || row['Bin Wise']),
  };
}

function setStatus(message) {
  el.renderStatus.textContent = message;
}

function updateFormViews() {
  el.productCode.value = state.form.productCode;
  el.batch.value = state.form.batch;
  el.expiry.value = state.form.expiry;

  el.genericCodeView.textContent = state.form.genericCode;
  el.productCodeView.textContent = state.form.productCode;
  el.descriptionView.textContent = state.form.description;
  el.batchView.textContent = state.form.batch;
  el.expiryView.textContent = displayDate(state.form.expiry);
  el.gtinView.textContent = state.form.gtin;
  el.locatorView.textContent = state.form.locator || '';
  el.timestampView.textContent = state.form.timestamp;
}

function fitDescription() {
  const node = el.descriptionView;
  const computed = window.getComputedStyle(node);
  const maxSize = Number.parseFloat(computed.getPropertyValue('--desc-max-size')) || 26;
  const minSize = Number.parseFloat(computed.getPropertyValue('--desc-min-size')) || 12;
  const step = 0.5;
  let size = maxSize;
  node.style.fontSize = `${size}px`;
  node.style.whiteSpace = 'normal';
  node.style.wordBreak = 'break-word';
  node.style.overflow = 'hidden';
  node.style.textOverflow = 'clip';

  const fits = () => node.scrollHeight <= node.clientHeight + 1 && node.scrollWidth <= node.clientWidth + 1;
  while (size > minSize && !fits()) {
    size -= step;
    node.style.fontSize = `${size}px`;
  }
}

function fitDescriptionForPrint() {
  const node = el.descriptionView;
  node.style.setProperty('--desc-max-size', '14');
  node.style.setProperty('--desc-min-size', '8');
  node.style.fontSize = '14px';
  fitDescription();
}

function buildBarcodePayload() {
  const gtin = text(state.form.gtin);
  const batch = text(state.form.batch);
  const expiry = gs1Date(state.form.expiry);
  if (!gtin || !batch || !expiry) return '';
  return `01${gtin}17${expiry}10${batch}`;
}

let _qrInstance = null;

function renderLocatorQr() {
  const locator = text(state.form.locator);
  el.locatorQr.innerHTML = '';
  _qrInstance = null;
  if (!locator) return;
  try {
    _qrInstance = new QRCode(el.locatorQr, {
      text: locator,
      width: 104,
      height: 104,
      colorDark: '#000000',
      colorLight: '#ffffff',
      correctLevel: QRCode.CorrectLevel.M,
    });
  } catch (e) {
    console.warn('QR render failed', e);
  }
}


  const payload = buildBarcodePayload();
  el.barcodeTextView.textContent = payload || 'Enter product code, batch number, and expiry date';
  if (!payload) {
    const ctx = el.barcodeCanvas.getContext('2d');
    ctx.clearRect(0, 0, el.barcodeCanvas.width, el.barcodeCanvas.height);
    ctx.fillStyle = '#6b7280';
    ctx.font = '600 18px "Space Grotesk", sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText('Awaiting input', el.barcodeCanvas.width / 2, el.barcodeCanvas.height / 2);
    setStatus('Waiting for workbook and inputs');
    return;
  }

  try {
    bwipjs.toCanvas(el.barcodeCanvas, {
      bcid: 'datamatrix',
      text: payload,
      scale: 6,
      padding: 2,
      backgroundcolor: 'FFFFFF',
    });
    setStatus('Rendered');
  } catch (error) {
    const ctx = el.barcodeCanvas.getContext('2d');
    ctx.clearRect(0, 0, el.barcodeCanvas.width, el.barcodeCanvas.height);
    ctx.fillStyle = '#dc2626';
    ctx.font = '600 14px "IBM Plex Mono", monospace';
    ctx.textAlign = 'center';
    ctx.fillText(String(error.message || error), el.barcodeCanvas.width / 2, el.barcodeCanvas.height / 2);
    setStatus('Barcode render failed');
  }
}

function renderLabel() {
  state.form.timestamp = timestampValue();
  updateFormViews();
  fitDescription();
  renderExpiryWarning();
  renderLocatorQr();
  renderBarcode();
}

function clearForm() {
  state.currentProduct = null;
  state.form = {
    productCode: '',
    batch: '',
    expiry: '',
    genericCode: '',
    description: '',
    gtin: '',
    locator: '',
    timestamp: '',
  };
  updateFormViews();
  renderBarcode();
}

function loadWorkbook(file) {
  if (!file) return;

  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx', 'xlsm', 'xls'].includes(ext)) {
    alert(`Unsupported file type ".${ext}". Please upload an Excel file (.xlsx, .xlsm, or .xls).`);
    return;
  }

  el.sourceStatus.textContent = `Loading "${file.name}"…`;
  el.recordCount.textContent = 'Reading…';
  setStatus('Reading workbook…');

  const reader = new FileReader();

  reader.onload = (e) => {
    try {
      const data = e.target.result;
      if (!data || data.byteLength === 0) throw new Error('File is empty or could not be read.');
      const workbook = XLSX.read(new Uint8Array(data), { type: 'array', cellDates: true });

      const sheetNames = workbook.SheetNames;
      if (!sheetNames.includes('Master')) {
        throw new Error(`Sheet "Master" not found. Available sheets: ${sheetNames.join(', ')}`);
      }

      state.workbookName = file.name;
      state.master = readSheet(workbook, 'Master');
      state.stock = readSheet(workbook, 'Tbl_StockList');
      buildIndex();

      el.sourceStatus.textContent = file.name;
      el.recordCount.textContent = `${state.master.length.toLocaleString()} products`;
      clearForm();
      setStatus(`Workbook loaded — ${state.master.length.toLocaleString()} products`);
    } catch (error) {
      console.error(error);
      el.sourceStatus.textContent = 'Load failed — see status for details';
      el.recordCount.textContent = 'No data';
      setStatus(`Error: ${error.message || error}`);
      alert(`Could not read the workbook:\n\n${error.message || error}`);
    }
  };

  reader.onerror = () => {
    const msg = reader.error ? reader.error.message : 'Unknown file read error';
    console.error('FileReader error:', reader.error);
    el.sourceStatus.textContent = 'File read failed';
    el.recordCount.textContent = 'No data';
    setStatus('File read error: ' + msg);
    alert('Could not read the file:\n\n' + msg);
  };

  reader.readAsArrayBuffer(file);
}

function applyProductLookup(productCode) {
  if (!state.index.length) return;
  const record = findProduct(productCode);
  if (!record) {
    state.currentProduct = null;
    state.form.productCode = text(productCode);
    state.form.genericCode = '';
    state.form.description = '';
    state.form.gtin = '';
    state.form.locator = '';
    updateFormViews();
    fitDescription();
    renderLocatorQr();
    renderBarcode();
    setStatus('Product code not found in workbook');
    return;
  }

  state.currentProduct = record.row;
  const derived = deriveProductFields(record.row, record.productCode);
  state.form.productCode = derived.productCode;
  state.form.genericCode = derived.genericCode;
  state.form.description = derived.description;
  state.form.gtin = derived.gtin;
  state.form.locator = derived.locator || '';
  updateFormViews();
  fitDescription();
  renderLocatorQr();
  renderBarcode();
  setStatus(`Loaded ${derived.productCode}`);
}

function copyBarcodeText() {
  const payload = buildBarcodePayload();
  if (!payload) return;
  navigator.clipboard?.writeText(payload).then(() => {
    setStatus('Barcode text copied');
  }).catch(() => {
    setStatus('Clipboard unavailable');
  });
}

async function downloadLabel() {
  const canvas = await html2canvas(el.labelPreview, {
    backgroundColor: '#ffffff',
    scale: 2,
    useCORS: true,
  });
  const link = document.createElement('a');
  const stamp = new Date().toISOString().replace(/[:.]/g, '-');
  link.download = `barcode-label-${state.form.productCode || 'label'}-${stamp}.png`;
  link.href = canvas.toDataURL('image/png');
  link.click();
}

function printLabel() {
  fitDescriptionForPrint();
  window.print();
}

function wireEvents() {
  // File input change
  el.workbookFile.addEventListener('change', () => {
    const file = el.workbookFile.files && el.workbookFile.files[0];
    if (file) loadWorkbook(file);
    // Reset so the same file can be re-uploaded if needed
    el.workbookFile.value = '';
  });

  // Drag-and-drop on the upload button label
  const uploadLabel = el.workbookFile.closest('label') || el.workbookFile.parentElement;
  if (uploadLabel) {
    uploadLabel.addEventListener('dragover', (e) => {
      e.preventDefault();
      uploadLabel.style.borderColor = 'rgba(45,212,191,0.7)';
      uploadLabel.style.background = 'rgba(45,212,191,0.18)';
    });
    uploadLabel.addEventListener('dragleave', () => {
      uploadLabel.style.borderColor = '';
      uploadLabel.style.background = '';
    });
    uploadLabel.addEventListener('drop', (e) => {
      e.preventDefault();
      uploadLabel.style.borderColor = '';
      uploadLabel.style.background = '';
      const file = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (file) loadWorkbook(file);
    });
  }

  el.productCode.addEventListener('input', () => {
    state.form.productCode = el.productCode.value.trim();
    if (state.form.productCode && state.index.length) {
      applyProductLookup(state.form.productCode);
    } else {
      renderLabel();
    }
  });

  el.batch.addEventListener('input', () => {
    state.form.batch = el.batch.value.trim();
    renderLabel();
  });

  const onExpiryChange = () => {
    state.form.expiry = el.expiry.value;
    const warning = renderExpiryWarning();
    if (warning) setStatus(warning);
    renderLabel();
  };

  el.expiry.addEventListener('input', onExpiryChange);
  el.expiry.addEventListener('change', onExpiryChange);

  el.applyBtn.addEventListener('click', () => {
    if (state.form.productCode && state.index.length) {
      applyProductLookup(state.form.productCode);
    }
    renderLabel();
  });

  el.copyBtn.addEventListener('click', copyBarcodeText);
  el.downloadBtn.addEventListener('click', () => { void downloadLabel(); });
  el.printBtn.addEventListener('click', printLabel);
}

function init() {
  el.sourceStatus.textContent = 'No workbook loaded';
  el.recordCount.textContent = 'No data';
  clearForm();
  wireEvents();
  window.addEventListener('beforeprint', fitDescriptionForPrint);
  window.addEventListener('afterprint', renderLabel);
  renderExpiryWarning();
}

init();
