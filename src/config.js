// テンプレート SpreadSheet ID
const TEMPLATES = {
  estimate: '1CJGUEnZk2nuwxXDKXO__EGlDZ5bAuY9ErFEJqxf9Q9E',
  invoice: '1Ln2FNWsLLnTrazkgKJobwXA3Cvwpz22YmI91zj0-fxo',
};

/**
 * Google Drive の URL からフォルダIDを抽出する
 * 例: https://drive.google.com/drive/folders/1ABC... → 1ABC...
 */
function parseFolderId(urlOrId) {
  if (!urlOrId) return '';
  const match = urlOrId.match(/\/folders\/([^/?]+)/);
  return match ? match[1] : urlOrId;
}

// 見積書テンプレートのセルマッピング
// Row  1: 御　見　積　書
// Row  2: A=宛名, H=No
// Row  3: H=見積日
// Row  4: (空)
// Row  5: A=「下記のとおり...」, H=SKテックラボ
// Row  6: A=件名ラベル, B=件名, H=清水勝紀
// Row  7: A=納期ラベル, B=納期, H=TEL
// Row  8: A=支払条件ラベル, B=支払条件, H=email
// Row  9: A=有効期限ラベル, B=有効期限, H=住所
// Row 10: (空)
// Row 11: A=合計金額ラベル, B=合計金額
// Row 12-13: (空)
// Row 14: 明細ヘッダー (番号, 摘要, 数量, 単位, 人日/円, 金額)
// Row 15-17: 明細行 (テンプレートは3行)
// Row 18: 合計行 (E=合計ラベル, H=合計金額)
// Row 19: (空)
// Row 20: 備考ラベル
// Row 21: 備考内容
const ESTIMATE_CELLS = {
  clientName: 'A2',
  no: 'H2',
  date: 'H3',
  subject: 'B6',
  deadline: 'B7',
  paymentTerms: 'B8',
  validPeriod: 'B9',
  totalAmount: 'B11',
  itemsHeaderRow: 14,
  itemsStartRow: 15,
  templateItemCount: 3,
  // 合計行 = itemsStartRow + itemCount
  // 備考ラベル = 合計行 + 2
  // 備考内容 = 合計行 + 3
  itemColumns: {
    number: 'A',
    description: 'B',
    quantity: 'E',
    unit: 'F',
    unitPrice: 'G',
    amount: 'H',
  },
};

// 請求書テンプレートのセルマッピング
// Row  1: 請　求　書
// Row  2: A=宛名, H=No
// Row  3: H=請求日
// Row  4: (空)
// Row  5: A=「下記のとおり...」, H=SKテックラボ
// Row  6: A=件名ラベル, B=件名, H=清水勝紀
// Row  7: A=納品日ラベル, B=納品日, H=住所
// Row  8: A=お支払期限ラベル, B=お支払期限, H=email
// Row  9: A=振込先ラベル, B=振込先, H=TEL
// Row 10: (空)
// Row 11: A=ご請求金額ラベル, B=ご請求金額
// Row 12-13: (空)
// Row 14: 明細ヘッダー
// Row 15-17: 明細行 (テンプレートは3行)
// Row 18: 合計行 (G=合計ラベル, H=合計金額)
// Row 19: (空)
// Row 20: 備考ラベル
// Row 21: 備考内容 (見積書番号等)
const INVOICE_CELLS = {
  clientName: 'A2',
  no: 'H2',
  date: 'H3',
  subject: 'B6',
  deliveryDate: 'B7',
  paymentDeadline: 'B8',
  bankInfo: 'B9',
  totalAmount: 'B11',
  itemsHeaderRow: 14,
  itemsStartRow: 15,
  templateItemCount: 3,
  itemColumns: {
    number: 'A',
    description: 'B',
    quantity: 'E',
    unit: 'F',
    unitPrice: 'G',
    amount: 'H',
  },
  remarks: 'A21',
};

// マスター管理シート
const MASTER_SHEET = {
  spreadsheetId: '1R0sZea4OrLlplY25uxqgRK3mGHGDgHPtd-fGvqDG1D4',
  estimateSheet: '見積書',
  invoiceSheet: '請求書',
};

module.exports = { TEMPLATES, parseFolderId, ESTIMATE_CELLS, INVOICE_CELLS, MASTER_SHEET };
