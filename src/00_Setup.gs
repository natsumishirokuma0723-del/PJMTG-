/**
 * スプレッドシートのメニューから最初に1回実行する初期化処理。
 * 必要なシートと見出し、デフォルト設定値を作成する。
 */
function initializeSheets() {
  const ss = getSs();

  createSheetIfMissing(ss, SHEET_NAMES.CONFIG, ['キー', '値'], [
    ['カレンダーID', 'primary'],
    ['ChatスペースID', 'spaces/xxxxxxxx'],
  ]);

  createSheetIfMissing(ss, SHEET_NAMES.SCHEDULE, [
    '会議名', '開始日時', '終了日時', '出席者(カンマ区切りメール)',
    '打刻対象メンバー(氏名カンマ区切り)',
    '会議ID', '会議コード', '会議記録ID', 'MeetURL', 'ステータス', '実績打刻済み',
  ]);

  createSheetIfMissing(ss, SHEET_NAMES.CHAT_LOG, [
    '会議ID', 'MessageID', '会議名', '議事録URL', '録画URL', '投稿日時',
  ]);

  createSheetIfMissing(ss, SHEET_NAMES.REACTIONS, [
    '会議ID', 'ユーザーID', '絵文字', '日時',
  ]);

  createSheetIfMissing(ss, SHEET_NAMES.ANALYSIS, [
    '会議ID', '会議名', '招待者数', '参加者数', '出席率',
    '既読者数', '既読率', '未参加者', '未読者', '集計日時',
  ]);

  createSheetIfMissing(ss, SHEET_NAMES.USER_MAP, [
    'ユーザーID(users/xxx)', 'メールアドレス', '氏名(活動タイマー用)',
  ]);

  SpreadsheetApp.getUi().alert(
    '初期化が完了しました。\n' +
    '「設定」シートの「ChatスペースID」と「カレンダーID」を必ず設定してから、' +
    '「② 定期実行トリガーを設定」を実行してください。\n\n' +
    '※「会議予定」シートの「打刻対象メンバー(氏名カンマ区切り)」列には、' +
    '⏱ 活動タイマー側の「メンバー_◯◯」シート名(氏名部分)と一致する名前を入力してください。\n\n' +
    '※実際の打刻は、Meetの参加ログに基づき会議終了後に自動で行われます。' +
    '「ユーザーID対応表」シートに現れた未登録のGoogleアカウントには、' +
    '氏名(活動タイマー用)の欄を入力してください。'
  );
}

function createSheetIfMissing(ss, name, headers, defaultRows) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    if (defaultRows) {
      defaultRows.forEach(row => sheet.appendRow(row));
    }
  }
}
