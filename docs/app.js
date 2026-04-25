// ============================================================
// Config
// ============================================================
const TEMPLATES = {
  estimate: '1k4TzjsW6N1w5kuflUJJLZbpAtztehzUXzVqXBIhGcvs',
  invoice: '1S2DGu4FCLDZwGZX0uS-vl4hEPtU7Y4SHnBPxiEAANTc',
};

const MASTER_SHEET = {
  spreadsheetId: '1ufhqWGyI0to5fR5y50qHILaCl46ljuuabtFbQ6tWZz0',
  estimateSheet: '見積書',
  invoiceSheet: '請求書',
};

const COMPANIES_SHEET = {
  spreadsheetId: '1jcWUiz2JxHi_fzn-1h75LOGx50DAiF5uWJPvIWEVlus',
  selfSheet: '自社情報',
  clientsSheet: '取引先企業',
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

const ROOT_FOLDER_ID = '1nX6rrInrQ2mBQK_Y-J6AO5FkWfc2-fh-';

const CLIENT_ID = '707875244824-kuhb9drhcanafjnqrs7fk9n7l3kjkssc.apps.googleusercontent.com';
const SCOPES = 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive';

// ============================================================
// State
// ============================================================
let tokenClient;
let gapiInited = false;
let gisInited = false;
let cachedClients = [];
let cachedSelfInfo = null;
let lastEstimateNo = '';
let lastInvoiceNo = '';
let folderPickerTarget = '';
let folderPickerStack = [];

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
  tokenClient.requestAccessToken({ prompt: 'select_account' });
}

function showLoggedIn() {
  document.getElementById('authSection').style.display = 'none';
  document.getElementById('mainApp').style.display = 'block';
  const today = new Date().toISOString().split('T')[0];
  const estDate = document.getElementById('est-date');
  const invDate = document.getElementById('inv-date');
  if (!estDate.value) estDate.value = today;
  if (!invDate.value) invDate.value = today;
  if (document.getElementById('est-items').children.length === 0) addItem('est');
  if (document.getElementById('inv-items').children.length === 0) addItem('inv');
  loadCompanies();
  loadLastNumbers();
}

function handleSignOut() {
  const token = gapi.client.getToken();
  if (token) {
    google.accounts.oauth2.revoke(token.access_token);
    gapi.client.setToken('');
  }
  document.getElementById('authSection').style.display = 'block';
  document.getElementById('mainApp').style.display = 'none';
}

// ============================================================
// Tabs
// ============================================================
function switchTab(name) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
  const tabs = { estimate: 0, invoice: 1, fromEstimate: 2, github: 3 };
  document.querySelectorAll('.tab')[tabs[name]].classList.add('active');
  document.getElementById('tab-' + name).classList.add('active');
}

// ============================================================
// Companies (自社情報 / 取引先企業)
// ============================================================
async function loadCompanies() {
  const statusEl = document.getElementById('companiesStatus');
  const setStatus = (text, isError) => {
    if (!statusEl) return;
    statusEl.textContent = text;
    statusEl.style.color = isError ? '#c00' : '#666';
  };
  setStatus('読込中...');
  try {
    const meta = await gapi.client.sheets.spreadsheets.get({
      spreadsheetId: COMPANIES_SHEET.spreadsheetId,
      fields: 'sheets.properties(title)',
    });
    const tabNames = meta.result.sheets.map(s => s.properties.title);
    const pickTab = (candidates) => tabNames.find(t => candidates.includes(t))
      || tabNames.find(t => candidates.some(c => t.includes(c)));
    const clientsTab = pickTab(['取引先企業', '取引先', 'クライアント', '顧客']);
    const selfTab = pickTab(['自社情報', '自社']);
    if (!clientsTab) throw new Error(`取引先タブが見つかりません。タブ一覧: [${tabNames.join(', ')}]`);

    const ranges = [`${clientsTab}!A1:Z1000`];
    if (selfTab) ranges.push(`${selfTab}!A1:Z10`);
    const batchRes = await gapi.client.sheets.spreadsheets.values.batchGet({
      spreadsheetId: COMPANIES_SHEET.spreadsheetId,
      ranges,
    });
    const [clientsValues, selfValues] = batchRes.result.valueRanges;
    cachedClients = parseCompanyRows(clientsValues.values);
    cachedSelfInfo = selfValues ? (parseCompanyRows(selfValues.values)[0] || null) : null;
    populateClientDropdowns();
    setStatus(`取引先 ${cachedClients.length} 件${selfTab ? ' / 自社情報読込済' : ''}`);
  } catch (err) {
    console.warn('会社情報の読み込みに失敗:', err);
    const code = err.status || err.result?.error?.code || '';
    const msg = err.result?.error?.message || err.message || String(err);
    setStatus(`× 読込失敗 ${code ? '(' + code + ')' : ''}: ${msg}`, true);
  }
}

function parseCompanyRows(values) {
  if (!values || values.length < 2) return [];
  const headers = values[0].map(h => String(h).trim());
  return values.slice(1)
    .filter(row => row && row.some(cell => String(cell || '').trim()))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = (row[i] || '').toString().trim(); });
      return obj;
    });
}

function populateClientDropdowns() {
  document.querySelectorAll('.client-select').forEach(sel => {
    const current = sel.value;
    const opts = ['<option value="">（手動入力）</option>']
      .concat(cachedClients.map((c, i) => `<option value="${i}">${escapeHtml(c['会社名'] || '')}</option>`));
    sel.innerHTML = opts.join('');
    sel.value = current;
  });
}

function onClientSelect(prefix, selectEl) {
  const idx = parseInt(selectEl.value, 10);
  if (isNaN(idx)) return;
  const client = cachedClients[idx];
  if (!client) return;
  const nameEl = document.getElementById(prefix + '-clientName');
  if (nameEl) nameEl.value = client['会社名'] || '';
}

async function reloadCompanies() {
  const btn = document.getElementById('reloadCompaniesBtn');
  if (btn) { btn.disabled = true; btn.textContent = '読み込み中...'; }
  await loadCompanies();
  if (btn) { btn.disabled = false; btn.textContent = '🔄 会社リスト再読込'; }
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

function buildInvoiceRemarks(remarks, estimateNo) {
  const header = estimateNo ? `見積書: ${estimateNo}` : '';
  const body = (remarks || '').trim();
  if (header && body) {
    return body.startsWith('見積書:') ? body : `${header}\n\n${body}`;
  }
  return header || body;
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
          inheritFromBefore: false,
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

async function autoResizeRemarksRow(spreadsheetId, sheetId, rowIndex) {
  await gapi.client.sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    resource: {
      requests: [
        {
          mergeCells: {
            range: { sheetId, startRowIndex: rowIndex, endRowIndex: rowIndex + 1, startColumnIndex: 0, endColumnIndex: 8 },
            mergeType: 'MERGE_ALL',
          },
        },
        {
          repeatCell: {
            range: { sheetId, startRowIndex: rowIndex, endRowIndex: rowIndex + 1, startColumnIndex: 0, endColumnIndex: 1 },
            cell: { userEnteredFormat: { wrapStrategy: 'WRAP' } },
            fields: 'userEnteredFormat.wrapStrategy',
          },
        },
        {
          autoResizeDimensions: {
            dimensions: { sheetId, dimension: 'ROWS', startIndex: rowIndex, endIndex: rowIndex + 1 },
          },
        },
      ],
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

    if (remarks) {
      await autoResizeRemarksRow(spreadsheetId, sheetId, remarksContentRow - 1);
    }

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
  const remarks = document.getElementById('inv-remarks').value.trim();
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
    const remarksValue = buildInvoiceRemarks(remarks, estimateNo);

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

    await autoResizeRemarksRow(spreadsheetId, sheetId, remarksRow - 1);

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

    const estRemarksRowIdx = 14 + items.length + 3;
    const estimateRemarks = (headerData[estRemarksRowIdx] && headerData[estRemarksRowIdx][0]) || '';

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
      { range: `A${remarksRow}`, values: [[buildInvoiceRemarks(estimateRemarks, estimateNo)]] },
      ...itemData,
    ];

    await batchUpdateValues(spreadsheetId, updateData);

    await autoResizeRemarksRow(spreadsheetId, sheetId, remarksRow - 1);

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

// ============================================================
// Settings (OpenAI & GitHub)
// ============================================================
function getSettings() {
  return {
    openaiKey: localStorage.getItem('openai_key') || '',
    openaiModel: localStorage.getItem('openai_model') || 'gpt-5-nano',
    githubToken: localStorage.getItem('github_token') || '',
  };
}

function openSettings() {
  const s = getSettings();
  document.getElementById('setting-openaiKey').value = s.openaiKey;
  document.getElementById('setting-openaiModel').value = s.openaiModel;
  document.getElementById('setting-githubToken').value = s.githubToken;
  document.getElementById('settingsModal').classList.add('show');
}

function closeSettings() {
  document.getElementById('settingsModal').classList.remove('show');
}

function saveSettings() {
  const k = document.getElementById('setting-openaiKey').value.trim();
  const m = document.getElementById('setting-openaiModel').value;
  const t = document.getElementById('setting-githubToken').value.trim();
  localStorage.setItem('openai_key', k);
  localStorage.setItem('openai_model', m);
  localStorage.setItem('github_token', t);
  closeSettings();
}

// ============================================================
// AI Fill (OpenAI API)
// ============================================================
async function aiFill(prefix) {
  const key = getSettings().openaiKey;
  if (!key) {
    alert('OpenAI APIキーを設定してください（右上の⚙設定）');
    openSettings();
    return;
  }
  const input = document.getElementById(prefix + '-aiInput').value.trim();
  if (!input) { alert('内容を入力してください'); return; }

  const statusEl = document.getElementById(prefix + '-aiStatus');
  statusEl.textContent = '解析中...';

  const isInvoice = prefix === 'inv';
  const docType = isInvoice ? '請求書' : '見積書';
  const extraFields = isInvoice
    ? '"deliveryDate" (納品日 YYYY/MM/DD), "paymentDeadline" (お支払期限 YYYY/MM/DD), "estimateNo" (関連する見積書番号), "remarks" (備考)'
    : '"deadline" (納期 自由文字列), "paymentTerms" (支払条件), "validPeriod" (有効期限), "remarks" (備考)';

  const remarksHeader = isInvoice
    ? '本請求には、以下の内容が含まれます。'
    : '本見積には、以下の機能開発および改善が含まれます。実装完了済みの内容を事後精算として計上しています。';

  const systemPrompt = `あなたは${docType}の入力を補助するAIです。ユーザーの自然言語による依頼から、以下のJSONを**必ず全フィールド埋めて**返してください。

## 必須フィールド（必ず全て含める）

- "clientName": 宛先の会社名（「御中」「株式会社」等は含めるが、「様」「殿」は除く）
- "subject": 件名（ユーザー入力から短く要約）
- "items": **必ず1件以上**。各要素は {"description": 摘要, "quantity": 数量(数値), "unit": 単位(デフォルト"式"), "unitPrice": 単価(円・数値)}
  - 入力が曖昧でも件名を分解して推測で2〜5件作る
  - 金額不明なら "unitPrice": 0 で埋める（省略しない）
  - 件名に「複数機能」「一式」系なら機能ごとに行を分ける
- "remarks": **必ず下記テンプレートに沿った複数行文字列**（\`\\n\`で改行）
- ${extraFields}

## remarks(備考)の必須フォーマット

必ず以下3セクション（【対応内容】【注意事項】【工数・単価】）を全て含める。入力情報が不足していても推測でプレースホルダーを埋める。

\`\`\`
${isInvoice ? '見積書: EST○○○  ← estimateNoがあるときのみ1行目に入れる\n\n' : ''}【対応内容】
${remarksHeader}

1. [items[0].descriptionに基づく作業タイトル]（[items[0].quantity]人日）
   - [具体的な作業内容1]
   - [具体的な作業内容2]

2. [items[1]があれば同様に]
   - ...

【注意事項】
1. [入力から読み取れる or 一般的な注意事項]
2. システム維持費は別途かかります

【工数・単価】
- 1人日あたりの単価目安: 約[items[0].unitPrice.toLocaleString()]円/人日
- 合計工数: [sum(items.quantity)]人日
\`\`\`

## 金額変換ルール
「30万」→300000, 「3.5万」→35000, 「1.5人日」→quantity:1.5

## 出力
**JSONオブジェクトのみ**。説明文・コードブロックは不要。`;

  try {
    const res = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${key}`,
      },
      body: JSON.stringify({
        model: getSettings().openaiModel,
        response_format: { type: 'json_object' },
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: input },
        ],
      }),
    });

    if (!res.ok) {
      const errText = await res.text();
      throw new Error(`OpenAI API エラー: ${res.status} ${errText}`);
    }

    const data = await res.json();
    const text = data.choices[0].message.content.trim();
    // JSON抽出（コードブロックが入っても除去）
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    if (!jsonMatch) throw new Error('JSONを抽出できませんでした: ' + text);
    const parsed = JSON.parse(jsonMatch[0]);

    applyAiResult(prefix, parsed);
    statusEl.textContent = '✓ 反映しました';
    setTimeout(() => { statusEl.textContent = ''; }, 3000);
  } catch (err) {
    console.error(err);
    statusEl.textContent = 'エラー: ' + err.message;
  }
}

function applyAiResult(prefix, data) {
  if (data.clientName) document.getElementById(prefix + '-clientName').value = data.clientName;
  if (data.subject) document.getElementById(prefix + '-subject').value = data.subject;

  if (prefix === 'est') {
    if (data.deadline) document.getElementById('est-deadline').value = data.deadline;
    if (data.paymentTerms) document.getElementById('est-paymentTerms').value = data.paymentTerms;
    if (data.validPeriod) document.getElementById('est-validPeriod').value = data.validPeriod;
    if (data.remarks) document.getElementById('est-remarks').value = data.remarks;
  } else if (prefix === 'inv') {
    if (data.deliveryDate) document.getElementById('inv-deliveryDate').value = data.deliveryDate;
    if (data.paymentDeadline) document.getElementById('inv-paymentDeadline').value = data.paymentDeadline;
    if (data.estimateNo) document.getElementById('inv-estimateNo').value = data.estimateNo;
    if (data.remarks) document.getElementById('inv-remarks').value = data.remarks;
  }

  if (Array.isArray(data.items)) {
    const container = document.getElementById(prefix + '-items');
    container.innerHTML = '';
    data.items.forEach(item => {
      addItem(prefix);
      const last = container.lastElementChild;
      const inputs = last.querySelectorAll('input');
      inputs[0].value = item.description || '';
      inputs[1].value = item.quantity ?? 1;
      inputs[2].value = item.unit || '式';
      inputs[3].value = item.unitPrice ?? 0;
    });
    if (data.items.length === 0) addItem(prefix);
  }
}

// ============================================================
// GitHub Integration
// ============================================================
let selectedRepo = null;
let cachedWorkItems = [];

async function ghFetch(path, params) {
  const token = getSettings().githubToken;
  if (!token) throw new Error('GitHub Personal Access Tokenを設定してください');
  const url = new URL('https://api.github.com' + path);
  if (params) Object.entries(params).forEach(([k, v]) => v != null && url.searchParams.set(k, v));
  const res = await fetch(url.toString(), {
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
    },
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`GitHub API エラー: ${res.status} ${text}`);
  }
  return res.json();
}

async function loadGitHubRepos() {
  const btn = document.getElementById('gh-loadReposBtn');
  btn.disabled = true;
  btn.textContent = '読み込み中...';
  try {
    const repos = await ghFetch('/user/repos', { per_page: 100, sort: 'updated', affiliation: 'owner,collaborator,organization_member' });
    const list = document.getElementById('gh-repoList');
    list.innerHTML = '';
    repos.forEach(r => {
      const div = document.createElement('div');
      div.className = 'gh-repo-item';
      div.innerHTML = `<div>${r.full_name}${r.private ? ' 🔒' : ''}</div><div class="meta">最終更新: ${r.updated_at.split('T')[0]}</div>`;
      div.onclick = () => selectRepo(r, div);
      list.appendChild(div);
    });
    list.style.display = 'block';
  } catch (err) {
    alert(err.message);
  } finally {
    btn.disabled = false;
    btn.textContent = 'リポジトリ一覧を取得';
  }
}

function selectRepo(repo, el) {
  document.querySelectorAll('.gh-repo-item').forEach(i => i.classList.remove('selected'));
  el.classList.add('selected');
  selectedRepo = repo;
  document.getElementById('gh-selectedRepo').value = repo.full_name;
  document.getElementById('gh-config').style.display = 'block';
}

async function fetchGitHubCommits() {
  if (!selectedRepo) { alert('リポジトリを選択してください'); return; }

  const branch = document.getElementById('gh-branch').value.trim();
  const since = document.getElementById('gh-since').value;
  const until = document.getElementById('gh-until').value;
  const author = document.getElementById('gh-author').value.trim();

  setLoading('gh-fetchLoading', true);
  try {
    const params = { per_page: 100 };
    if (branch) params.sha = branch;
    if (since) params.since = since + 'T00:00:00Z';
    if (until) params.until = until + 'T23:59:59Z';
    if (author) params.author = author;

    const allCommits = [];
    let page = 1;
    while (page <= 5) {
      params.page = page;
      const commits = await ghFetch(`/repos/${selectedRepo.full_name}/commits`, params);
      if (commits.length === 0) break;
      allCommits.push(...commits);
      if (commits.length < 100) break;
      page++;
    }

    if (allCommits.length === 0) {
      alert('該当するコミットが見つかりませんでした');
      return;
    }

    const commits = allCommits.map(c => ({
      hash: c.sha,
      date: c.commit.author.date.split('T')[0],
      message: c.commit.message.split('\n')[0],
    }));

    cachedWorkItems = groupCommitsToWorkItems(commits);
    renderCommitsPreview(cachedWorkItems, commits.length);
    document.getElementById('gh-toEstimate').style.display = 'block';
  } catch (err) {
    alert(err.message);
  } finally {
    setLoading('gh-fetchLoading', false);
  }
}

function groupCommitsToWorkItems(commits) {
  const filtered = commits.filter(c => {
    const msg = c.message.toLowerCase();
    return !msg.startsWith('merge') && !msg.startsWith('auto-') && msg.length > 0;
  });
  const groups = new Map();
  for (const commit of filtered) {
    let description = commit.message;
    const conventionalMatch = description.match(/^(?:feat|fix|refactor|chore|docs|style|test|perf|ci|build)(?:\(.+?\))?:\s*(.+)/i);
    if (conventionalMatch) description = conventionalMatch[1];
    description = description.replace(/^\[?[A-Z]+-\d+\]?\s*/, '');
    const key = description.substring(0, 50).toLowerCase().trim();
    if (groups.has(key)) {
      groups.get(key).commits++;
    } else {
      groups.set(key, { description, commits: 1 });
    }
  }
  return Array.from(groups.values());
}

function renderCommitsPreview(workItems, totalCommits) {
  const el = document.getElementById('gh-commitsPreview');
  el.style.display = 'block';
  el.innerHTML = `<div style="margin-bottom:6px;color:#555;">合計 ${totalCommits} コミット → ${workItems.length} 件の作業項目</div>` +
    workItems.map((w, i) => `<div class="work-item">${i + 1}. ${escapeHtml(w.description)} <span style="color:#999;">(${w.commits} commits)</span></div>`).join('');
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// ============================================================
// Last Number Management
// ============================================================
async function loadLastNumbers() {
  try {
    const [estValues, invValues] = await Promise.all([
      readRange(MASTER_SHEET.spreadsheetId, `${MASTER_SHEET.estimateSheet}!C:C`),
      readRange(MASTER_SHEET.spreadsheetId, `${MASTER_SHEET.invoiceSheet}!C:C`),
    ]);
    lastEstimateNo = findLastNo(estValues, 'EST');
    lastInvoiceNo = findLastNo(invValues, 'BILL');
    updateNoDisplay('est', lastEstimateNo);
    updateNoDisplay('inv', lastInvoiceNo);
  } catch (err) {
    console.warn('番号読み込み失敗:', err);
    document.getElementById('est-lastNo').textContent = '番号の読み込みに失敗';
    document.getElementById('inv-lastNo').textContent = '番号の読み込みに失敗';
  }
}

function findLastNo(values, prefix) {
  if (!values) return '';
  for (let i = values.length - 1; i >= 0; i--) {
    const v = (values[i][0] || '').toString().trim();
    if (v && v.startsWith(prefix)) return v;
  }
  return '';
}

function incrementNo(no) {
  const match = no.match(/^([A-Z]+\d{6})(\d+)$/);
  if (!match) return no;
  const nextNum = parseInt(match[2], 10) + 1;
  return match[1] + String(nextNum).padStart(match[2].length, '0');
}

function updateNoDisplay(prefix, lastNo) {
  const lastEl = document.getElementById(prefix + '-lastNo');
  const noInput = document.getElementById(prefix + '-no');
  if (lastNo) {
    lastEl.textContent = '直前の番号: ' + lastNo;
    noInput.value = incrementNo(lastNo);
  } else {
    lastEl.textContent = '過去の番号なし';
    noInput.value = prefix === 'est' ? generateEstimateNo() : generateInvoiceNo();
  }
}

function onNoModeChange(prefix) {
  const mode = document.getElementById(prefix + '-noMode').value;
  const noInput = document.getElementById(prefix + '-no');
  const lastNo = prefix === 'est' ? lastEstimateNo : lastInvoiceNo;
  if (mode === 'next') {
    noInput.readOnly = true;
    noInput.value = lastNo ? incrementNo(lastNo) : (prefix === 'est' ? generateEstimateNo() : generateInvoiceNo());
  } else if (mode === 'same') {
    noInput.readOnly = true;
    noInput.value = lastNo || (prefix === 'est' ? generateEstimateNo() : generateInvoiceNo());
  } else {
    noInput.readOnly = false;
    noInput.value = '';
    noInput.placeholder = '番号を入力';
    noInput.focus();
  }
}

// ============================================================
// Folder Picker
// ============================================================
function openFolderPicker(targetPrefix) {
  folderPickerTarget = targetPrefix;
  folderPickerStack = [{ id: ROOT_FOLDER_ID, name: 'Root' }];
  document.getElementById('folderPickerModal').classList.add('show');
  loadFolderContents(ROOT_FOLDER_ID);
}

function closeFolderPicker() {
  document.getElementById('folderPickerModal').classList.remove('show');
}

async function loadFolderContents(folderId) {
  const listEl = document.getElementById('fp-list');
  const loadingEl = document.getElementById('fp-loading');
  listEl.innerHTML = '';
  loadingEl.className = 'loading show';
  try {
    const res = await gapi.client.drive.files.list({
      q: `'${folderId}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false`,
      fields: 'files(id, name)',
      orderBy: 'name',
      pageSize: 100,
    });
    loadingEl.className = 'loading';
    const folders = res.result.files || [];
    renderBreadcrumb();
    if (folders.length === 0) {
      listEl.innerHTML = '<div style="padding: 12px; color: #888; font-size: 0.9em;">サブフォルダなし</div>';
      return;
    }
    folders.forEach(f => {
      const div = document.createElement('div');
      div.className = 'folder-item';
      div.innerHTML = `<span style="font-size:1.1em;">📁</span> ${escapeHtml(f.name)}`;
      div.onclick = () => navigateToFolder(f.id, f.name);
      listEl.appendChild(div);
    });
  } catch (err) {
    loadingEl.className = 'loading';
    listEl.innerHTML = `<div style="padding: 12px; color: #c00; font-size: 0.9em;">エラー: ${escapeHtml(err.result?.error?.message || err.message || String(err))}</div>`;
  }
}

function navigateToFolder(folderId, folderName) {
  folderPickerStack.push({ id: folderId, name: folderName });
  loadFolderContents(folderId);
}

function navigateToBreadcrumb(index) {
  if (index >= folderPickerStack.length - 1) return;
  folderPickerStack = folderPickerStack.slice(0, index + 1);
  loadFolderContents(folderPickerStack[index].id);
}

function renderBreadcrumb() {
  const el = document.getElementById('fp-breadcrumb');
  el.innerHTML = folderPickerStack.map((item, i) => {
    if (i < folderPickerStack.length - 1) {
      return `<span onclick="navigateToBreadcrumb(${i})">${escapeHtml(item.name)}</span><span style="color:#999;cursor:default;"> / </span>`;
    }
    return `<span>${escapeHtml(item.name)}</span>`;
  }).join('');
}

function confirmFolderSelection() {
  const current = folderPickerStack[folderPickerStack.length - 1];
  const url = `https://drive.google.com/drive/folders/${current.id}`;
  document.getElementById(folderPickerTarget + '-folderUrl').value = url;
  closeFolderPicker();
}

async function submitGitHubEstimate() {
  if (cachedWorkItems.length === 0) { alert('先にコミットを取得してください'); return; }
  const clientName = document.getElementById('gh-clientName').value.trim();
  const subject = document.getElementById('gh-subject').value.trim();
  const folderUrl = document.getElementById('gh-folderUrl').value.trim();
  const unitPrice = parseFloat(document.getElementById('gh-unitPrice').value) || 24000;
  if (!clientName || !subject) { alert('宛先と件名は必須です'); return; }
  if (!folderUrl) { alert('保存先フォルダURLは必須です'); return; }

  const items = cachedWorkItems.map(w => ({
    description: w.description,
    quantity: 1,
    unit: '式',
    unitPrice,
  }));

  const folderId = parseFolderId(folderUrl);
  const no = generateEstimateNo();
  const date = formatDate(new Date());

  document.getElementById('gh-submit').disabled = true;
  setLoading('gh-loading', true);

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

    const updateData = [
      { range: cfg.clientName, values: [[`${clientName} 御中`]] },
      { range: cfg.no, values: [[`No: ${no}`]] },
      { range: cfg.date, values: [[`見積日: ${date}`]] },
      { range: cfg.subject, values: [[subject]] },
      { range: cfg.deadline, values: [['']] },
      { range: cfg.paymentTerms, values: [['別途協議の上決定']] },
      { range: cfg.validPeriod, values: [['御見積後1ヶ月']] },
      { range: cfg.totalAmount, values: [[`${total.toLocaleString()} 円 (税込)`]] },
      { range: `${cols.amount}${totalRow}`, values: [[total]] },
      ...itemData,
    ];

    await batchUpdateValues(spreadsheetId, updateData);

    const today = formatDate(new Date());
    await appendRow(MASTER_SHEET.spreadsheetId, MASTER_SHEET.estimateSheet, [
      today, today, no, subject, clientName, '', spreadsheetUrl,
    ]);

    showResult('gh-result', `見積書を作成しました: <a href="${spreadsheetUrl}" target="_blank">${spreadsheetUrl}</a>`);
  } catch (err) {
    console.error(err);
    showResult('gh-result', `エラー: ${err.result?.error?.message || err.message || err}`, true);
  } finally {
    document.getElementById('gh-submit').disabled = false;
    setLoading('gh-loading', false);
  }
}
