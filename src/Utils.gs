/**
 * シート名の定義
 */
const SHEET_NAMES = {
  CONFIG: '設定',
  SCHEDULE: '会議予定',
  TIME_LOG: '打刻ログ',
  CHAT_LOG: 'Chat投稿ログ',
  REACTIONS: 'リアクション集計',
  ANALYSIS: '参加分析',
  USER_MAP: 'ユーザーID対応表',
};

function getSs() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  const sheet = getSs().getSheetByName(name);
  if (!sheet) {
    throw new Error(`シート「${name}」が見つかりません。メニューの「① 初期化(シート作成)」を先に実行してください。`);
  }
  return sheet;
}

/**
 * ヘッダー行(配列)から { 見出し名: 列インデックス(0始まり) } のマップを作る
 */
function colIndexMap(headerRow) {
  const map = {};
  headerRow.forEach((h, i) => {
    if (h) map[h] = i;
  });
  return map;
}

/**
 * 「設定」シートからキーに対応する値を取得する
 */
function getConfig(key) {
  const sheet = getSheet(SHEET_NAMES.CONFIG);
  const data = sheet.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (data[r][0] === key) return data[r][1];
  }
  return null;
}

/**
 * Chatのユーザーリソース名(users/xxxxx)からメールアドレスを解決する。
 * Chat API / Meet API はプライバシー上メールアドレスを直接返さないため、
 * 「ユーザーID対応表」シートでの手動マッピングを正としている。
 * 未登録のIDが出てきた場合は自動で行を追加するので、管理者はメール欄を埋めるだけでよい。
 */
function resolveUserEmail(userResourceName) {
  if (!userResourceName) return '';
  const sheet = getSheet(SHEET_NAMES.USER_MAP);
  const data = sheet.getDataRange().getValues();

  for (let r = 1; r < data.length; r++) {
    if (data[r][0] === userResourceName) {
      return data[r][1] || '';
    }
  }

  // 未登録IDは追記しておく(メール欄は空のまま。管理者が後で埋める)
  sheet.appendRow([userResourceName, '']);
  return '';
}
