// ============================================================
// Config
// ============================================================
const TEMPLATES = {
  estimate: '1CJGUEnZk2nuwxXDKXO__EGlDZ5bAuY9ErFEJqxf9Q9E',
  invoice: '1Ln2FNWsLLnTrazkgKJobwXA3Cvwpz22YmI91zj0-fxo',
};

const MASTER_SHEET = {
  spreadsheetId: '1R0sZea4OrLlplY25uxqgRK3mGHGDgHPtd-fGvqDG1D4',
  estimateSheet: '見積書',
  invoiceSheet: '請求書',
};

const ESTIMATE_CELLS = {
  clientName: 'A2', no: 'H2', date: 'H3',
  subject: 'B6', deadline: 'B7', paymentTerms: 'B8', validPeriod: 'B9',
  totalAmount: 'B11',
  itemsStartRow: 15, templateItemCount: 3,
  itemColumns: { number: 'A', description: 'B', quantity: 'E', unit: 'F', unitPrice: 'G', amount: 'H' },
};

const INVOICE_CELLS = {
  clientName: 'A2', no: 'H2', date: 'H3',
  subject: 'B6', deliveryDate: 'B7', paymentDeadline: 'B8', bankInfo: 'B9',
  totalAmount: 'B11',
  itemsStartRow: 15, templateItemCount: 3,
  itemColumns: { number: 'A', description: 'B', quantity: 'E', unit: 'F', unitPrice: 'G', amount: 'H' },
};

const CLIENT_ID = '707875244824-kuhb9drhcanafjnqrs7fk9n7l3kjkssc.apps.googleusercontent.com';
const SCOPES = 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive';

// ============================================================
// State
// ============================================================
let tokenClient;
let gapiInited = false;
let gisInited = false;

// ============================================================
// Google API init
// ============================================================
function gapiLoaded() {
  gapi.load('client', async () => {
    await gapi.client.init({
      discoveryDocs: [
        'https://sheets.googleapis.com/$discovery/rest?version=v4',
        'https://www.googleapis.com/discovery/v1/apis/drive/v3/rest',
      ],
    });
    gapiInited = true;
    maybeEnableSignIn();
  });
}

function gisLoaded() {
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CLIENT_ID,
    scope: SCOPES,
    callback: '',
  });
  gisInited = true;
  maybeEnableSignIn();
}

function maybeEnableSignIn() {
  if (gapiInited && gisInited) {
    document.getElementById('signInBtn').style.display = 'inline-flex';
    if (gapi.client.getToken()) {
      showLoggedIn();
    }
  }
}

// ============================================================
// Auth
// ============================================================
function handleSignIn() {
  tokenClient.callback = async (resp) => {
    if (resp.error) {
      console.error(resp);
      return;
    }
    showLoggedIn();
  };
  if (!gapi.client.getToken()) {
    tokenClient.requestAccessToken({ prompt: 'consent' });
  } else {
    tokenClient.requestAccessToken({ prompt: '' });
  }
}

function showLoggedIn() {
  document.getElementById('preAuth').style.display = 'none';
  document.getElementById('postAuth').style.display = 'block';
  document.getElementById('mainApp').style.display = 'block';
  document.getElementById('userName').textContent = 'Google アカウント';
  // Set today's date
  const today = new Date().toISOString().split('T')[0];
  const estDate = document.getElementById('est-date');
  const invDate = document.getElementById('inv-date');
  if (!estDate.value) estDate.value = today;
  if (!invDate.value) invDate.value = today;
  // Init item rows
  if (document.getElementById('est-items').children.length === 0) addItem('est');
  if (document.getElementById('inv-items').children.length === 0) addItem('inv');
}

function handleSignOut() {
  const token = gapi.client.getToken();
  if (token) {
    google.accounts.oauth2.revoke(token.access_token);
    gapi.client.setToken('');
  }
  document.getElementById('preAuth').style.display = 'block';
  document.getElementById('postAuth').style.display = 'none';
  document.getElementById('mainApp').style.display = 'none';
}

// ============================================================
// Tabs
// ============================================================
function switchTab(name) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
  const tabs = { estimate: 0, invoice: 1, fromEstimate: 2 };
  document.querySelectorAll('.tab')[tabs[name]].classList.add('active');
  document.getElementById('tab-' + name).classList.add('active');
}

// ============================================================
// Item rows
// ============================================================
function addItem(prefix) {
  const container = document.getElementById(prefix + '-items');
  const row = document.createElement('div');
  row.className = 'item-row';
  row.innerHTML = `
    <input type="text" placeholder="摘要">
    <input type="number" placeholder="数量" value="1" min="0" step="any">
    <input type="text" placeholder="単位" value="式">
    <input type="number" placeholder="単価" min="0">
    <button onclick="this.parentElement.remove()">×</button>
  `;
  container.appendChild(row);
}

function getItems(prefix) {
  const rows = document.querySelectorAll('#' + prefix + '-items .item-row');
  const items = [];
  rows.forEach(row => {
    const inputs = row.querySelectorAll('input');
    const desc = inputs[0].value.trim();
    if (!desc) return;
    items.push({
      description: desc,
      quantity: parseFloat(inputs[1].value) || 1,
      unit: inputs[2].value || '式',
      unitPrice: parseFloat(inputs[3].value) || 0,
    });
  });
  return items;
}

// ============================================================
// Helpers
// ============================================================
function formatDate(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}/${m}/${day}`;
}

function dateInputToFormatted(val) {
  if (!val) return formatDate(new Date());
  const d = new Date(val + 'T00:00:00');
  return formatDate(d);
}

function generateEstimateNo() {
  const now = new Date();
  return `EST${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}001`;
}

function generateInvoiceNo() {
  const now = new Date();
  return `BILL${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}001`;
}

function parseFolderId(urlOrId) {
  if (!urlOrId) return '';
  const match = urlOrId.match(/\/folders\/([^/?]+)/);
  return match ? match[1] : urlOrId;
}

function parseSpreadsheetId(urlOrId) {
  if (!urlOrId) return '';
  const match = urlOrId.match(/\/d\/([^/?]+)/);
  return match ? match[1] : urlOrId;
}

function showResult(id, html, isError) {
  const el = document.getElementById(id);
  el.innerHTML = html;
  el.className = isError ? 'result error' : 'result';
  el.style.display = 'block';
}

function setLoading(id, show) {
  document.getElementById(id).className = show ? 'loading show' : 'loading';
}

// ============================================================
// Sheets API wrappers
// ============================================================
async function copyTemplate(templateId, title, folderId) {
  const requestBody = { name: title };
  if (folderId) requestBody.parents = [folderId];
  const res = await gapi.client.drive.files.copy({
    fileId: templateId,
    resource: requestBody,
  });
  const spreadsheetId = res.result.id;
  return {
    spreadsheetId,
    spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`,
  };
}

async function getFirstSheetId(spreadsheetId) {
  const res = await gapi.client.sheets.spreadsheets.get({
    spreadsheetId,
    fields: 'sheets.properties.sheetId',
  });
  return res.result.sheets[0].properties.sheetId;
}

async function batchUpdateValues(spreadsheetId, data) {
  await gapi.client.sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    resource: {
      valueInputOption: 'USER_ENTERED',
      data,
    },
  });
}

async function clearRanges(spreadsheetId, ranges) {
  await gapi.client.sheets.spreadsheets.values.batchClear({
    spreadsheetId,
    resource: { ranges },
  });
}

async function readRange(spreadsheetId, range, valueRenderOption) {
  const params = { spreadsheetId, range };
  if (valueRenderOption) params.valueRenderOption = valueRenderOption;
  const res = await gapi.client.sheets.spreadsheets.values.get(params);
  return res.result.values || [];
}

async function insertRows(spreadsheetId, sheetId, startIndex, count) {
  await gapi.client.sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    resource: {
      requests: [{
        insertDimension: {
          range: { sheetId, dimension: 'ROWS', startIndex, endIndex: startIndex + count },
          inheritFromBefore: true,
        },
      }],
    },
  });
}

async function deleteRows(spreadsheetId, sheetId, startIndex, count) {
  await gapi.client.sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    resource: {
      requests: [{
        deleteDimension: {
          range: { sheetId, dimension: 'ROWS', startIndex, endIndex: startIndex + count },
        },
      }],
    },
  });
}

async function appendRow(spreadsheetId, sheetName, rowValues) {
  await gapi.client.sheets.spreadsheets.values.append({
    spreadsheetId,
    range: `${sheetName}!A:A`,
    valueInputOption: 'USER_ENTERED',
    resource: { values: [rowValues] },
  });
}

// ============================================================
// Create Estimate
// ============================================================
async function submitEstimate() {
  const clientName = document.getElementById('est-clientName').value.trim();
  const subject = document.getElementById('est-subject').value.trim();
  const folderUrl = document.getElementById('est-folderUrl').value.trim();
  if (!clientName || !subject) { alert('宛先と件名は必須です'); return; }
  if (!folderUrl) { alert('保存先フォルダURLは必須です'); return; }

  const items = getItems('est');
  if (items.length === 0) { alert('明細を1行以上入力してください'); return; }

  const no = document.getElementById('est-no').value.trim() || generateEstimateNo();
  const date = dateInputToFormatted(document.getElementById('est-date').value);
  const deadline = document.getElementById('est-deadline').value.trim();
  const paymentTerms = document.getElementById('est-paymentTerms').value.trim() || '別途協議の上決定';
  const validPeriod = document.getElementById('est-validPeriod').value.trim() || '御見積後1ヶ月';
  const remarks = document.getElementById('est-remarks').value.trim();
  const folderId = parseFolderId(folderUrl);

  document.getElementById('est-submit').disabled = true;
  setLoading('est-loading', true);

  try {
    const title = `見積書_${clientName}_${no}`;
    const { spreadsheetId, spreadsheetUrl } = await copyTemplate(TEMPLATES.estimate, title, folderId);

    const cfg = ESTIMATE_CELLS;
    const cols = cfg.itemColumns;
    const startRow = cfg.itemsStartRow;
    const templateCount = cfg.templateItemCount;

    const sheetId = await getFirstSheetId(spreadsheetId);
    if (items.length > templateCount) {
      await insertRows(spreadsheetId, sheetId, startRow - 1, items.length - templateCount);
    } else if (items.length < templateCount) {
      await deleteRows(spreadsheetId, sheetId, startRow - 1, templateCount - items.length);
    }

    const total = items.reduce((sum, item) => sum + item.quantity * item.unitPrice, 0);
    const totalRow = startRow + items.length;

    await clearRanges(spreadsheetId, [
      `${cols.number}${startRow}:${cols.amount}${startRow + items.length - 1}`,
    ]);

    const itemData = items.map((item, i) => {
      const row = startRow + i;
      return {
        range: `${cols.number}${row}:${cols.amount}${row}`,
        values: [[i + 1, item.description, '', '', item.quantity, item.unit, item.unitPrice, item.quantity * item.unitPrice]],
      };
    });

    const remarksContentRow = totalRow + 3;
    const updateData = [
      { range: cfg.clientName, values: [[`${clientName} 御中`]] },
      { range: cfg.no, values: [[`No: ${no}`]] },
      { range: cfg.date, values: [[`見積日: ${date}`]] },
      { range: cfg.subject, values: [[subject]] },
      { range: cfg.deadline, values: [[deadline]] },
      { range: cfg.paymentTerms, values: [[paymentTerms]] },
      { range: cfg.validPeriod, values: [[validPeriod]] },
      { range: cfg.totalAmount, values: [[`${total.toLocaleString()} 円 (税込)`]] },
      { range: `${cols.amount}${totalRow}`, values: [[total]] },
      ...itemData,
    ];
    if (remarks) {
      updateData.push({ range: `A${remarksContentRow}`, values: [[remarks]] });
    }

    await batchUpdateValues(spreadsheetId, updateData);

    // マスター管理シートに追記
    const today = formatDate(new Date());
    await appendRow(MASTER_SHEET.spreadsheetId, MASTER_SHEET.estimateSheet, [
      today, today, no, subject, clientName, '', spreadsheetUrl,
    ]);

    showResult('est-result', `見積書を作成しました: <a href="${spreadsheetUrl}" target="_blank">${spreadsheetUrl}</a>`);
  } catch (err) {
    console.error(err);
    showResult('est-result', `エラー: ${err.result?.error?.message || err.message || err}`, true);
  } finally {
    document.getElementById('est-submit').disabled = false;
    setLoading('est-loading', false);
  }
}

// ============================================================
// Create Invoice
// ============================================================
async function submitInvoice() {
  const clientName = document.getElementById('inv-clientName').value.trim();
  const subject = document.getElementById('inv-subject').value.trim();
  const folderUrl = document.getElementById('inv-folderUrl').value.trim();
  if (!clientName || !subject) { alert('宛先と件名は必須です'); return; }
  if (!folderUrl) { alert('保存先フォルダURLは必須です'); return; }

  const items = getItems('inv');
  if (items.length === 0) { alert('明細を1行以上入力してください'); return; }

  const no = document.getElementById('inv-no').value.trim() || generateInvoiceNo();
  const date = dateInputToFormatted(document.getElementById('inv-date').value);
  const deliveryDate = document.getElementById('inv-deliveryDate').value.trim();
  const paymentDeadline = document.getElementById('inv-paymentDeadline').value.trim();
  const estimateNo = document.getElementById('inv-estimateNo').value.trim();
  const folderId = parseFolderId(folderUrl);

  document.getElementById('inv-submit').disabled = true;
  setLoading('inv-loading', true);

  try {
    const title = `請求書_${clientName}_${no}`;
    const { spreadsheetId, spreadsheetUrl } = await copyTemplate(TEMPLATES.invoice, title, folderId);

    const cfg = INVOICE_CELLS;
    const cols = cfg.itemColumns;
    const startRow = cfg.itemsStartRow;
    const templateCount = cfg.templateItemCount;

    const sheetId = await getFirstSheetId(spreadsheetId);
    if (items.length > templateCount) {
      await insertRows(spreadsheetId, sheetId, startRow - 1, items.length - templateCount);
    } else if (items.length < templateCount) {
      await deleteRows(spreadsheetId, sheetId, startRow - 1, templateCount - items.length);
    }

    const total = items.reduce((sum, item) => sum + item.quantity * item.unitPrice, 0);
    const totalRow = startRow + items.length;

    await clearRanges(spreadsheetId, [
      `${cols.number}${startRow}:${cols.amount}${startRow + items.length - 1}`,
    ]);

    const itemData = items.map((item, i) => {
      const row = startRow + i;
      return {
        range: `${cols.number}${row}:${cols.amount}${row}`,
        values: [[i + 1, item.description, '', '', item.quantity, item.unit, item.unitPrice, item.quantity * item.unitPrice]],
      };
    });

    const remarksRow = totalRow + 3;
    const remarksValue = estimateNo ? `1. 見積書: ${estimateNo}` : '';

    const updateData = [
      { range: cfg.clientName, values: [[`${clientName} 御中`]] },
      { range: cfg.no, values: [[`No: ${no}`]] },
      { range: cfg.date, values: [[`請求日: ${date}`]] },
      { range: cfg.subject, values: [[subject]] },
      { range: cfg.deliveryDate, values: [[deliveryDate]] },
      { range: cfg.paymentDeadline, values: [[paymentDeadline]] },
      { range: cfg.totalAmount, values: [[`${total.toLocaleString()} 円 (税込)`]] },
      { range: `${cols.amount}${totalRow}`, values: [[total]] },
      { range: `A${remarksRow}`, values: [[remarksValue]] },
      ...itemData,
    ];

    await batchUpdateValues(spreadsheetId, updateData);

    // マスター管理シートに追記
    const today = formatDate(new Date());
    await appendRow(MASTER_SHEET.spreadsheetId, MASTER_SHEET.invoiceSheet, [
      today, today, no, subject, clientName, '', '送', spreadsheetUrl,
    ]);

    showResult('inv-result', `請求書を作成しました: <a href="${spreadsheetUrl}" target="_blank">${spreadsheetUrl}</a>`);
  } catch (err) {
    console.error(err);
    showResult('inv-result', `エラー: ${err.result?.error?.message || err.message || err}`, true);
  } finally {
    document.getElementById('inv-submit').disabled = false;
    setLoading('inv-loading', false);
  }
}

// ============================================================
// Create Invoice from Estimate
// ============================================================
async function submitInvoiceFromEstimate() {
  const estimateUrl = document.getElementById('fe-estimateUrl').value.trim();
  const folderUrl = document.getElementById('fe-folderUrl').value.trim();
  if (!estimateUrl) { alert('見積書のURLは必須です'); return; }
  if (!folderUrl) { alert('保存先フォルダURLは必須です'); return; }

  const estimateSpreadsheetId = parseSpreadsheetId(estimateUrl);
  const folderId = parseFolderId(folderUrl);
  const deliveryDate = document.getElementById('fe-deliveryDate').value.trim();
  const paymentDeadline = document.getElementById('fe-paymentDeadline').value.trim();
  const noOverride = document.getElementById('fe-no').value.trim();

  document.getElementById('fe-submit').disabled = true;
  setLoading('fe-loading', true);

  try {
    // Read estimate data
    const headerData = await readRange(estimateSpreadsheetId, 'A1:H30', 'UNFORMATTED_VALUE');

    let clientName = '';
    let estimateNo = '';
    if (headerData[1]) {
      clientName = (headerData[1][0] || '').replace(/\s*御中$/, '');
      const noCell = headerData[1][7] || '';
      estimateNo = String(noCell).replace(/^No:\s*/, '');
    }

    let subject = '';
    if (headerData[5]) {
      subject = headerData[5][1] || '';
    }

    const items = [];
    for (let i = 14; i < headerData.length; i++) {
      const row = headerData[i];
      if (!row || !row[0]) break;
      const num = parseInt(row[0]);
      if (isNaN(num)) break;
      const quantity = parseFloat(row[4]) || 1;
      const rawUnitPrice = parseFloat(String(row[6]).replace(/,/g, '').trim()) || 0;
      const amount = parseFloat(String(row[7]).replace(/,/g, '').trim()) || 0;
      const unitPrice = rawUnitPrice || (quantity ? amount / quantity : 0);
      items.push({
        description: row[1] || '',
        quantity,
        unit: row[5] || '式',
        unitPrice,
      });
    }

    if (items.length === 0) {
      showResult('fe-result', 'エラー: 見積書から明細を読み取れませんでした', true);
      return;
    }

    // Create invoice
    const no = noOverride || generateInvoiceNo();
    const date = formatDate(new Date());
    const title = `請求書_${clientName}_${no}`;
    const { spreadsheetId, spreadsheetUrl } = await copyTemplate(TEMPLATES.invoice, title, folderId);

    const cfg = INVOICE_CELLS;
    const cols = cfg.itemColumns;
    const startRow = cfg.itemsStartRow;
    const templateCount = cfg.templateItemCount;

    const sheetId = await getFirstSheetId(spreadsheetId);
    if (items.length > templateCount) {
      await insertRows(spreadsheetId, sheetId, startRow - 1, items.length - templateCount);
    } else if (items.length < templateCount) {
      await deleteRows(spreadsheetId, sheetId, startRow - 1, templateCount - items.length);
    }

    const total = items.reduce((sum, item) => sum + item.quantity * item.unitPrice, 0);
    const totalRow = startRow + items.length;

    await clearRanges(spreadsheetId, [
      `${cols.number}${startRow}:${cols.amount}${startRow + items.length - 1}`,
    ]);

    const itemData = items.map((item, i) => {
      const row = startRow + i;
      return {
        range: `${cols.number}${row}:${cols.amount}${row}`,
        values: [[i + 1, item.description, '', '', item.quantity, item.unit, item.unitPrice, item.quantity * item.unitPrice]],
      };
    });

    const remarksRow = totalRow + 3;

    const updateData = [
      { range: cfg.clientName, values: [[`${clientName} 御中`]] },
      { range: cfg.no, values: [[`No: ${no}`]] },
      { range: cfg.date, values: [[`請求日: ${date}`]] },
      { range: cfg.subject, values: [[subject]] },
      { range: cfg.deliveryDate, values: [[deliveryDate]] },
      { range: cfg.paymentDeadline, values: [[paymentDeadline]] },
      { range: cfg.totalAmount, values: [[`${total.toLocaleString()} 円 (税込)`]] },
      { range: `${cols.amount}${totalRow}`, values: [[total]] },
      { range: `A${remarksRow}`, values: [[`1. 見積書: ${estimateNo}`]] },
      ...itemData,
    ];

    await batchUpdateValues(spreadsheetId, updateData);

    // マスター管理シートに追記
    const today = formatDate(new Date());
    await appendRow(MASTER_SHEET.spreadsheetId, MASTER_SHEET.invoiceSheet, [
      today, today, no, subject, clientName, '', '送', spreadsheetUrl,
    ]);

    showResult('fe-result', `請求書を作成しました: <a href="${spreadsheetUrl}" target="_blank">${spreadsheetUrl}</a><br>元の見積書: ${estimateNo}`);
  } catch (err) {
    console.error(err);
    showResult('fe-result', `エラー: ${err.result?.error?.message || err.message || err}`, true);
  } finally {
    document.getElementById('fe-submit').disabled = false;
    setLoading('fe-loading', false);
  }
}
