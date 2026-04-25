const { TEMPLATES, parseFolderId, ESTIMATE_CELLS } = require('./config');
const { copyTemplate, getFirstSheetId, batchUpdate, clearRanges, insertRows, deleteRows, autoResizeRemarksRow } = require('./sheets');

/**
 * 見積書を作成する
 * @param {object} auth - OAuth2 クライアント
 * @param {object} params
 * @param {string} params.clientName - 宛先 (例: "株式会社〇〇")
 * @param {string} params.subject - 件名
 * @param {Array<{description: string, quantity: number, unit: string, unitPrice: number}>} params.items - 明細
 * @param {string} [params.date] - 見積日 (YYYY/MM/DD) デフォルト: 今日
 * @param {string} [params.no] - 見積番号 デフォルト: 自動生成
 * @param {string} [params.deadline] - 納期
 * @param {string} [params.paymentTerms] - 支払条件 デフォルト: "別途協議の上決定"
 * @param {string} [params.validPeriod] - 有効期限 デフォルト: "御見積後1ヶ月"
 * @param {string} [params.remarks] - 備考
 * @returns {object} { spreadsheetId, spreadsheetUrl, no }
 */
async function createEstimate(auth, params) {
  const {
    clientName,
    subject,
    items,
    date = formatDate(new Date()),
    no = generateEstimateNo(),
    deadline = '',
    paymentTerms = '別途協議の上決定',
    validPeriod = '御見積後1ヶ月',
    remarks = '',
  } = params;

  const title = `見積書_${clientName}_${no}`;
  const folderId = parseFolderId(params.folderUrl);
  if (!folderId) {
    throw new Error('folderUrl は必須です。Google Drive のフォルダURLを指定してください。');
  }
  const { spreadsheetId, spreadsheetUrl } = await copyTemplate(
    auth,
    TEMPLATES.estimate,
    title,
    folderId
  );

  const cfg = ESTIMATE_CELLS;
  const cols = cfg.itemColumns;
  const startRow = cfg.itemsStartRow;
  const templateCount = cfg.templateItemCount;

  // テンプレートの明細行数と実際の明細行数の差分を調整
  const sheetId = await getFirstSheetId(auth, spreadsheetId);
  if (items.length > templateCount) {
    await insertRows(auth, spreadsheetId, sheetId, startRow - 1, items.length - templateCount);
  } else if (items.length < templateCount) {
    await deleteRows(auth, spreadsheetId, sheetId, startRow - 1, templateCount - items.length);
  }

  // 合計金額
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

  // 備考行の位置（合計行の2行下 = ラベル、3行下 = 内容）
  const remarksContentRow = totalRow + 3;

  const updateData = [
    // ヘッダー
    { range: cfg.clientName, values: [[`${clientName} 御中`]] },
    { range: cfg.no, values: [[`No: ${no}`]] },
    { range: cfg.date, values: [[`見積日: ${date}`]] },
    // 詳細
    { range: cfg.subject, values: [[subject]] },
    { range: cfg.deadline, values: [[deadline]] },
    { range: cfg.paymentTerms, values: [[paymentTerms]] },
    { range: cfg.validPeriod, values: [[validPeriod]] },
    // 合計金額 (Row 11 相当、行挿入/削除の影響を受けない位置)
    { range: cfg.totalAmount, values: [[`${total.toLocaleString()} 円 (税込)`]] },
    // 合計行 (明細テーブル末尾)
    { range: `${cols.amount}${totalRow}`, values: [[total]] },
    // 明細
    ...itemData,
  ];

  // 備考がある場合
  if (remarks) {
    updateData.push({ range: `A${remarksContentRow}`, values: [[remarks]] });
  }

  await batchUpdate(auth, spreadsheetId, updateData);

  if (remarks) {
    await autoResizeRemarksRow(auth, spreadsheetId, sheetId, remarksContentRow - 1);
  }

  console.log(`見積書を作成しました: ${spreadsheetUrl}`);
  return { spreadsheetId, spreadsheetUrl, no };
}

function formatDate(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}/${m}/${day}`;
}

function generateEstimateNo() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  return `EST${y}${m}001`;
}

module.exports = { createEstimate, formatDate, generateEstimateNo };
