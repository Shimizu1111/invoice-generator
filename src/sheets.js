const { google } = require('googleapis');

/**
 * テンプレートのスプレッドシートをコピーして新しいスプレッドシートを作成する
 * @param {object} auth - OAuth2 クライアント
 * @param {string} templateId - テンプレートのスプレッドシートID
 * @param {string} title - 新しいスプレッドシートのタイトル
 * @param {string} [folderId] - 保存先フォルダID（省略時はマイドライブ直下）
 * @returns {object} { spreadsheetId, spreadsheetUrl }
 */
async function copyTemplate(auth, templateId, title, folderId) {
  const drive = google.drive({ version: 'v3', auth });
  const requestBody = { name: title };
  if (folderId) {
    requestBody.parents = [folderId];
  }
  const res = await drive.files.copy({
    fileId: templateId,
    requestBody,
  });
  const spreadsheetId = res.data.id;
  const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;
  return { spreadsheetId, spreadsheetUrl };
}

/**
 * スプレッドシートの最初のシートIDを取得する
 */
async function getFirstSheetId(auth, spreadsheetId) {
  const sheets = google.sheets({ version: 'v4', auth });
  const res = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: 'sheets.properties.sheetId',
  });
  return res.data.sheets[0].properties.sheetId;
}

/**
 * スプレッドシートの複数セルを一括更新する
 * @param {object} auth - OAuth2 クライアント
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {Array<{range: string, values: any[][]}>} data - 更新データの配列
 */
async function batchUpdate(auth, spreadsheetId, data) {
  const sheets = google.sheets({ version: 'v4', auth });
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data,
    },
  });
}

/**
 * スプレッドシートの指定範囲をクリアする
 * @param {object} auth - OAuth2 クライアント
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string[]} ranges - クリアする範囲の配列
 */
async function clearRanges(auth, spreadsheetId, ranges) {
  const sheets = google.sheets({ version: 'v4', auth });
  await sheets.spreadsheets.values.batchClear({
    spreadsheetId,
    requestBody: { ranges },
  });
}

/**
 * スプレッドシートの指定範囲を読み取る
 * @param {object} auth - OAuth2 クライアント
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} range - 範囲
 * @returns {any[][]} values
 */
async function readRange(auth, spreadsheetId, range, valueRenderOption) {
  const sheets = google.sheets({ version: 'v4', auth });
  const params = { spreadsheetId, range };
  if (valueRenderOption) {
    params.valueRenderOption = valueRenderOption;
  }
  const res = await sheets.spreadsheets.values.get(params);
  return res.data.values || [];
}

/**
 * 行を挿入する（明細行が足りない場合に使用）
 * @param {object} auth - OAuth2 クライアント
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {number} sheetId - シートID (通常0)
 * @param {number} startIndex - 挿入開始行 (0-based)
 * @param {number} count - 挿入行数
 */
async function insertRows(auth, spreadsheetId, sheetId, startIndex, count) {
  const sheets = google.sheets({ version: 'v4', auth });
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        {
          insertDimension: {
            range: {
              sheetId,
              dimension: 'ROWS',
              startIndex,
              endIndex: startIndex + count,
            },
            inheritFromBefore: true,
          },
        },
      ],
    },
  });
}

/**
 * 行を削除する（余分な明細行を削除する場合に使用）
 * @param {object} auth - OAuth2 クライアント
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {number} sheetId - シートID (通常0)
 * @param {number} startIndex - 削除開始行 (0-based)
 * @param {number} count - 削除行数
 */
async function deleteRows(auth, spreadsheetId, sheetId, startIndex, count) {
  const sheets = google.sheets({ version: 'v4', auth });
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        {
          deleteDimension: {
            range: {
              sheetId,
              dimension: 'ROWS',
              startIndex,
              endIndex: startIndex + count,
            },
          },
        },
      ],
    },
  });
}

/**
 * スプレッドシートの指定シートに行を追加する
 * @param {object} auth - OAuth2 クライアント
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {any[]} rowValues - 追加する行の値の配列
 */
async function appendRow(auth, spreadsheetId, sheetName, rowValues) {
  const sheets = google.sheets({ version: 'v4', auth });
  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: `${sheetName}!A:A`,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [rowValues],
    },
  });
}

/**
 * 備考セルのテキスト折り返しを有効にし、行高さを自動調整する
 */
async function autoResizeRemarksRow(auth, spreadsheetId, sheetId, rowIndex) {
  const sheets = google.sheets({ version: 'v4', auth });
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
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

module.exports = {
  copyTemplate,
  getFirstSheetId,
  batchUpdate,
  clearRanges,
  readRange,
  insertRows,
  deleteRows,
  appendRow,
  autoResizeRemarksRow,
};
