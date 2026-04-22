const { authorize } = require('./src/auth');
const { createEstimate, formatDate } = require('./src/estimate');
const { createInvoice, createInvoiceFromEstimate } = require('./src/invoice');
const { getCommitHistory, groupCommitsToWorkItems } = require('./src/gitUtils');
const { appendRow } = require('./src/sheets');
const { MASTER_SHEET } = require('./src/config');

async function main() {
  const command = process.argv[2];

  if (!command || command === '--help') {
    printUsage();
    return;
  }

  const auth = await authorize();

  switch (command) {
    case 'estimate':
      await handleEstimate(auth);
      break;
    case 'invoice':
      await handleInvoice(auth);
      break;
    case 'invoice-from-estimate':
      await handleInvoiceFromEstimate(auth);
      break;
    case 'estimate-from-git':
      await handleEstimateFromGit(auth);
      break;
    default:
      console.error(`不明なコマンド: ${command}`);
      printUsage();
      process.exit(1);
  }
}

async function handleEstimate(auth) {
  const json = process.argv[3];
  if (!json) {
    console.error('JSON パラメータを指定してください');
    console.error('例: node index.js estimate \'{"clientName":"株式会社テスト","subject":"開発作業","items":[{"description":"機能A開発","quantity":1,"unit":"式","unitPrice":50000}]}\'');
    process.exit(1);
  }
  const params = JSON.parse(json);
  const result = await createEstimate(auth, params);
  console.log(JSON.stringify(result, null, 2));

  // マスター管理シートに追記
  const today = formatDate(new Date());
  // 見積書: 作成日, 申請日付, ID, タイトル, 会社名, ファイル名, リンク
  await appendRow(auth, MASTER_SHEET.spreadsheetId, MASTER_SHEET.estimateSheet, [
    today, today, result.no, params.subject, params.clientName, '', result.spreadsheetUrl,
  ]);
}

async function handleInvoice(auth) {
  const json = process.argv[3];
  if (!json) {
    console.error('JSON パラメータを指定してください');
    process.exit(1);
  }
  const params = JSON.parse(json);
  const result = await createInvoice(auth, params);
  console.log(JSON.stringify(result, null, 2));

  // マスター管理シートに追記
  const today = formatDate(new Date());
  // 請求書: 作成日, 申請日付, ID, タイトル, 会社名, ファイル名, 送/受, リンク
  await appendRow(auth, MASTER_SHEET.spreadsheetId, MASTER_SHEET.invoiceSheet, [
    today, today, result.no, params.subject, params.clientName, '', '送', result.spreadsheetUrl,
  ]);
}

async function handleInvoiceFromEstimate(auth) {
  const estimateId = process.argv[3];
  const overridesJson = process.argv[4];
  if (!estimateId) {
    console.error('見積書のスプレッドシートIDを指定してください');
    console.error('例: node index.js invoice-from-estimate <spreadsheet-id> [overrides-json]');
    process.exit(1);
  }
  const overrides = overridesJson ? JSON.parse(overridesJson) : {};
  const result = await createInvoiceFromEstimate(auth, estimateId, overrides);
  console.log(JSON.stringify(result, null, 2));

  // マスター管理シートに追記（件名・会社名は見積書から読み取った値を使用）
  const today = formatDate(new Date());
  // 請求書: 作成日, 申請日付, ID, タイトル, 会社名, ファイル名, 送/受, リンク
  await appendRow(auth, MASTER_SHEET.spreadsheetId, MASTER_SHEET.invoiceSheet, [
    today, today, result.no, result.subject || '', result.clientName || '', '', '送', result.spreadsheetUrl,
  ]);
}

async function handleEstimateFromGit(auth) {
  const json = process.argv[3];
  if (!json) {
    console.error('JSON パラメータを指定してください');
    console.error('例: node index.js estimate-from-git \'{"repoPath":"/path/to/repo","clientName":"株式会社テスト","subject":"開発作業","since":"2024-01-01","unitPrice":24000}\'');
    process.exit(1);
  }
  const params = JSON.parse(json);
  const { repoPath, clientName, subject, since, until, author, branch, unitPrice = 24000 } = params;

  if (!repoPath || !clientName || !subject) {
    console.error('repoPath, clientName, subject は必須です');
    process.exit(1);
  }

  const commits = getCommitHistory({ repoPath, since, until, author, branch });
  if (commits.length === 0) {
    console.error('該当するコミットが見つかりませんでした');
    process.exit(1);
  }

  const workItems = groupCommitsToWorkItems(commits);
  console.log(`Git 履歴から ${workItems.length} 件の作業項目を抽出しました:`);
  workItems.forEach((item, i) => {
    console.log(`  ${i + 1}. ${item.description} (${item.commits} commits)`);
  });

  const items = workItems.map((item) => ({
    description: item.description,
    quantity: 1,
    unit: '式',
    unitPrice,
  }));

  const result = await createEstimate(auth, {
    clientName,
    subject,
    items,
    ...params,
  });
  console.log(JSON.stringify(result, null, 2));

  // マスター管理シートに追記
  const today = formatDate(new Date());
  await appendRow(auth, MASTER_SHEET.spreadsheetId, MASTER_SHEET.estimateSheet, [
    today, today, result.no, subject, clientName, '', result.spreadsheetUrl,
  ]);
}

function printUsage() {
  console.log(`
請求書・見積書自動作成ツール

使い方:
  node index.js <command> [params]

コマンド:
  estimate              見積書を作成
  invoice               請求書を作成
  invoice-from-estimate 見積書から請求書を作成
  estimate-from-git     Git履歴から見積書を作成

見積書作成の例:
  node index.js estimate '{
    "clientName": "株式会社テスト",
    "subject": "システム開発",
    "items": [
      {"description": "機能A開発", "quantity": 1, "unit": "式", "unitPrice": 50000},
      {"description": "機能B開発", "quantity": 1, "unit": "式", "unitPrice": 30000}
    ],
    "deadline": "2024/03/31"
  }'

請求書作成の例:
  node index.js invoice '{
    "clientName": "株式会社テスト",
    "subject": "システム開発",
    "items": [
      {"description": "機能A開発", "quantity": 1, "unit": "式", "unitPrice": 50000}
    ],
    "deliveryDate": "2024/03/15",
    "paymentDeadline": "2024/04/30",
    "estimateNo": "EST202403001"
  }'

見積書から請求書を作成:
  node index.js invoice-from-estimate <spreadsheet-id> '{"deliveryDate":"2024/03/15","paymentDeadline":"2024/04/30"}'

Git履歴から見積書を作成:
  node index.js estimate-from-git '{
    "repoPath": "/path/to/repo",
    "clientName": "株式会社テスト",
    "subject": "3月分開発作業",
    "since": "2024-03-01",
    "until": "2024-03-31",
    "unitPrice": 24000
  }'
  `);
}

main().catch((err) => {
  console.error('エラー:', err.message);
  process.exit(1);
});
