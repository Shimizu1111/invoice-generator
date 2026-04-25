const { TEMPLATES, parseFolderId, INVOICE_CELLS } = require('./config');
const { copyTemplate, getFirstSheetId, batchUpdate, clearRanges, readRange, insertRows, deleteRows, autoResizeRemarksRow, mergeItemDescCells } = require('./sheets');
const { formatDate } = require('./estimate');

/**
 * 請求書を作成する
 * @param {object} auth - OAuth2 クライアント
 * @param {object} params
 * @param {string} params.clientName - 宛先 (例: "株式会社〇〇")
 * @param {string} params.subject - 件名
 * @param {Array<{description: string, quantity: number, unit: string, unitPrice: number}>} params.items - 明細
 * @param {string} [params.date] - 請求日 (YYYY/MM/DD) デフォルト: 今日
 * @param {string} [params.no] - 請求番号 デフォルト: 自動生成
 * @param {string} [params.deliveryDate] - 納品日
 * @param {string} [params.paymentDeadline] - お支払期限
 * @param {string} [params.estimateNo] - 紐付ける見積書番号
 * @returns {object} { spreadsheetId, spreadsheetUrl, no }
 */
async function createInvoice(auth, params) {
  const {
    clientName,
    subject,
    items,
    date = formatDate(new Date()),
    no = generateInvoiceNo(),
    deliveryDate = '',
    paymentDeadline = '',
    estimateNo = '',
  } = params;

  const title = `請求書_${clientName}_${no}`;
  const folderId = parseFolderId(params.folderUrl);
  if (!folderId) {
    throw new Error('folderUrl は必須です。Google Drive のフォルダURLを指定してください。');
  }
  const { spreadsheetId, spreadsheetUrl } = await copyTemplate(
    auth,
    TEMPLATES.invoice,
    title,
    folderId
  );

  const cfg = INVOICE_CELLS;
  const cols = cfg.itemColumns;
  const startRow = cfg.itemsStartRow;
  const templateCount = cfg.templateItemCount;

  // 行数の調整
  const sheetId = await getFirstSheetId(auth, spreadsheetId);
  if (items.length > templateCount) {
    await insertRows(auth, spreadsheetId, sheetId, startRow, items.length - templateCount);
  } else if (items.length < templateCount) {
    await deleteRows(auth, spreadsheetId, sheetId, startRow - 1, templateCount - items.length);
  }

  const total = items.reduce((sum, item) => sum + item.quantity * item.unitPrice, 0);
  const totalRow = startRow + items.length;

  // 明細行のクリア
  await clearRanges(auth, spreadsheetId, [
    `${cols.number}${startRow}:${cols.amount}${startRow + items.length - 1}`,
  ]);

  // 明細データ
  const itemData = items.map((item, i) => {
    const row = startRow + i;
    const amount = item.quantity * item.unitPrice;
    return {
      range: `${cols.number}${row}:${cols.amount}${row}`,
      values: [[
        i + 1,
        item.description,
        '', '',
        item.quantity,
        item.unit || '式',
        item.unitPrice,
        amount,
      ]],
    };
  });

  // 備考欄（見積書番号がある場合）
  const remarksRow = totalRow + 3; // 合計行 + 空行 + 備考ラベル + 備考内容
  const remarksValue = estimateNo ? `1. 見積書: ${estimateNo}` : '';

  const updateData = [
    { range: cfg.clientName, values: [[`${clientName} 御中`]] },
    { range: cfg.no, values: [[`No: ${no}`]] },
    { range: cfg.date, values: [[`請求日: ${date}`]] },
    { range: cfg.subject, values: [[subject]] },
    { range: cfg.deliveryDate, values: [[deliveryDate]] },
    { range: cfg.paymentDeadline, values: [[paymentDeadline]] },
    // 合計金額
    { range: cfg.totalAmount, values: [[`${total.toLocaleString()} 円 (税込)`]] },
    // 合計行
    { range: `${cols.amount}${totalRow}`, values: [[total]] },
    // 備考
    { range: `A${remarksRow}`, values: [[remarksValue]] },
    // 明細
    ...itemData,
  ];

  await batchUpdate(auth, spreadsheetId, updateData);
  await mergeItemDescCells(auth, spreadsheetId, sheetId, startRow, items.length);
  await autoResizeRemarksRow(auth, spreadsheetId, sheetId, remarksRow - 1);

  console.log(`請求書を作成しました: ${spreadsheetUrl}`);
  return { spreadsheetId, spreadsheetUrl, no };
}

/**
 * 見積書スプレッドシートから明細を読み取って請求書を作成する
 * @param {object} auth - OAuth2 クライアント
 * @param {string} estimateSpreadsheetId - 見積書のスプレッドシートID
 * @param {object} overrides - 上書きパラメータ
 * @returns {object} { spreadsheetId, spreadsheetUrl, no }
 */
async function createInvoiceFromEstimate(auth, estimateSpreadsheetId, overrides = {}) {
  const headerData = await readRange(auth, estimateSpreadsheetId, 'A1:H30', 'UNFORMATTED_VALUE');

  // クライアント名 (Row 2, A列)
  let clientName = '';
  let estimateNo = '';
  if (headerData[1]) {
    clientName = (headerData[1][0] || '').replace(/\s*御中$/, '');
    const noCell = headerData[1][7] || '';
    estimateNo = noCell.replace(/^No:\s*/, '');
  }

  // 件名 (Row 6, B列)
  let subject = '';
  if (headerData[5]) {
    subject = headerData[5][1] || '';
  }

  // 明細行を読み取る (Row 14 = ヘッダー、Row 15~ = データ)
  const items = [];
  for (let i = 14; i < headerData.length; i++) {
    const row = headerData[i];
    if (!row || !row[0]) break;
    const num = parseInt(row[0]);
    if (isNaN(num)) break;
    const quantity = parseFloat(row[4]) || 1;
    const rawUnitPrice = parseFloat(String(row[6]).replace(/,/g, '').trim()) || 0;
    const amount = parseFloat(String(row[7]).replace(/,/g, '').trim()) || 0;
    // 単価列が空の場合、金額 ÷ 数量 で単価を算出する
    const unitPrice = rawUnitPrice || (quantity ? amount / quantity : 0);
    items.push({
      description: row[1] || '',
      quantity,
      unit: row[5] || '式',
      unitPrice,
    });
  }

  const result = await createInvoice(auth, {
    clientName,
    subject,
    items,
    estimateNo,
    ...overrides,
  });
  result.clientName = overrides.clientName || clientName;
  result.subject = overrides.subject || subject;
  return result;
}

function generateInvoiceNo() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  return `BILL${y}${m}001`;
}

module.exports = { createInvoice, createInvoiceFromEstimate };
