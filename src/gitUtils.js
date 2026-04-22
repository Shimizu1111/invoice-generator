const { execSync } = require('child_process');
const path = require('path');

/**
 * 指定ディレクトリの Git コミット履歴から作業内容を抽出する
 * @param {object} options
 * @param {string} options.repoPath - Git リポジトリのパス
 * @param {string} [options.since] - 開始日 (YYYY-MM-DD)
 * @param {string} [options.until] - 終了日 (YYYY-MM-DD)
 * @param {string} [options.author] - 著者でフィルタ
 * @param {string} [options.branch] - ブランチ名 デフォルト: 現在のブランチ
 * @returns {Array<{hash: string, date: string, message: string}>}
 */
function getCommitHistory(options) {
  const { repoPath, since, until, author, branch } = options;

  const args = ['git', 'log', '--pretty=format:%H|%ad|%s', '--date=short'];

  if (since) args.push(`--since=${since}`);
  if (until) args.push(`--until=${until}`);
  if (author) args.push(`--author=${author}`);
  if (branch) args.push(branch);

  try {
    const output = execSync(args.join(' '), {
      cwd: path.resolve(repoPath),
      encoding: 'utf8',
    }).trim();

    if (!output) return [];

    return output.split('\n').map((line) => {
      const [hash, date, ...messageParts] = line.split('|');
      return { hash, date, message: messageParts.join('|') };
    });
  } catch (err) {
    throw new Error(`Git 履歴の取得に失敗しました: ${err.message}`);
  }
}

/**
 * コミット履歴を作業項目としてグループ化する
 * 類似のコミットメッセージをまとめて、見積書/請求書の明細に使える形にする
 * @param {Array<{hash: string, date: string, message: string}>} commits
 * @returns {Array<{description: string, commits: number}>}
 */
function groupCommitsToWorkItems(commits) {
  // merge コミットや自動コミットを除外
  const filtered = commits.filter((c) => {
    const msg = c.message.toLowerCase();
    return !msg.startsWith('merge') && !msg.startsWith('auto-') && msg.length > 0;
  });

  // コミットメッセージから作業項目を抽出
  // 同じプレフィックスのコミットをグループ化
  const groups = new Map();

  for (const commit of filtered) {
    // コミットメッセージの先頭部分をキーにする
    // 例: "feat: カレンダー機能追加" → "カレンダー機能追加"
    // 例: "fix: バグ修正" → "バグ修正"
    let description = commit.message;

    // Conventional Commits 形式のプレフィックスを除去
    const conventionalMatch = description.match(/^(?:feat|fix|refactor|chore|docs|style|test|perf|ci|build)(?:\(.+?\))?:\s*(.+)/i);
    if (conventionalMatch) {
      description = conventionalMatch[1];
    }

    // チケット番号を除去
    description = description.replace(/^\[?[A-Z]+-\d+\]?\s*/, '');

    // 最初の50文字で類似判定のキーを作る
    const key = description.substring(0, 50).toLowerCase().trim();

    if (groups.has(key)) {
      groups.get(key).commits++;
    } else {
      groups.set(key, { description, commits: 1 });
    }
  }

  return Array.from(groups.values());
}

module.exports = { getCommitHistory, groupCommitsToWorkItems };
