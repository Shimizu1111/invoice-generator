const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');
const http = require('http');
const url = require('url');

const CREDENTIALS_PATH = path.join(__dirname, '..', 'credentials.json');
const TOKEN_PATH = path.join(__dirname, '..', 'token.json');
const SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive',
];

function loadCredentials() {
  if (!fs.existsSync(CREDENTIALS_PATH)) {
    throw new Error(
      `credentials.json が見つかりません: ${CREDENTIALS_PATH}\n` +
      'Google Cloud Console から OAuth クライアント ID (デスクトップアプリ) を作成し、\n' +
      'credentials.json をプロジェクトルートに配置してください。'
    );
  }
  return JSON.parse(fs.readFileSync(CREDENTIALS_PATH, 'utf8'));
}

function getOAuth2Client() {
  const credentials = loadCredentials();
  const { client_id, client_secret, redirect_uris } =
    credentials.installed || credentials.web;
  return new google.auth.OAuth2(
    client_id,
    client_secret,
    redirect_uris ? redirect_uris[0] : 'http://localhost'
  );
}

async function authorize() {
  const oAuth2Client = getOAuth2Client();

  if (fs.existsSync(TOKEN_PATH)) {
    const token = JSON.parse(fs.readFileSync(TOKEN_PATH, 'utf8'));
    oAuth2Client.setCredentials(token);

    // トークンの有効期限チェック・自動リフレッシュ
    oAuth2Client.on('tokens', (tokens) => {
      const current = JSON.parse(fs.readFileSync(TOKEN_PATH, 'utf8'));
      const updated = { ...current, ...tokens };
      fs.writeFileSync(TOKEN_PATH, JSON.stringify(updated, null, 2));
    });

    return oAuth2Client;
  }

  return await getNewToken(oAuth2Client);
}

function getNewToken(oAuth2Client) {
  return new Promise((resolve, reject) => {
    const authUrl = oAuth2Client.generateAuthUrl({
      access_type: 'offline',
      scope: SCOPES,
      prompt: 'select_account consent',
    });

    console.log('以下のURLをブラウザで開いて認証してください:\n');
    console.log(authUrl);
    console.log('\n認証後のリダイレクトを待機中...');

    const server = http.createServer(async (req, res) => {
      const parsedUrl = url.parse(req.url, true);
      if (parsedUrl.pathname === '/oauth2callback') {
        const code = parsedUrl.query.code;
        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end('<h1>認証成功！このタブを閉じてください。</h1>');
        server.close();

        try {
          const { tokens } = await oAuth2Client.getToken(code);
          oAuth2Client.setCredentials(tokens);
          fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens, null, 2));
          console.log('トークンを保存しました:', TOKEN_PATH);
          resolve(oAuth2Client);
        } catch (err) {
          reject(new Error(`トークン取得エラー: ${err.message}`));
        }
      }
    });

    server.listen(3000, () => {
      console.log('ローカルサーバーをポート3000で起動しました');
    });

    server.on('error', (err) => {
      if (err.code === 'EADDRINUSE') {
        reject(new Error('ポート3000が使用中です。他のプロセスを停止してから再試行してください。'));
      } else {
        reject(err);
      }
    });
  });
}

module.exports = { authorize };
