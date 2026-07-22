/**
 * スプレッドシートのメニューから最初に1回実行する初期化処理。
 * 必要なシートと見出し、デフォルト設定値を作成する。
 */
function initializeSheets() {
  const ss = getSs();

  createSheetIfMissing(ss, SHEET_NAMES.CONFIG, ['キー', '値'], [
    ['カレンダーID', 'primary'],
    ['ChatスペースID', 'spaces/xxxxxxxx'],
    ['Notion会議データベースID', ''],
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
    '初期化が完了しました。\n\n' +
    '■ Google側でやること\n' +
    '「設定」シートの「ChatスペースID」「カレンダーID」「Notion会議データベースID」を入力してください。\n\n' +
    '■ Notion側でやること(このスプレッドシートとは別に準備が必要)\n' +
    '1. https://www.notion.so/my-integrations で内部インテグレーションを作成し、シークレットを控える\n' +
    '2. Apps Scriptの「プロジェクトの設定」→「スクリプトプロパティ」に NOTION_TOKEN として保存する\n' +
    '3. 会議管理用のデータベースを作成し、①のインテグレーションと接続(共有)する\n' +
    '4. データベースに次のプロパティを用意する: 会議名(タイトル) / 開始日時(日付) / 終了日時(日付) / ' +
    '出席者メール(テキスト) / 打刻対象メンバー(テキスト) / ステータス(セレクト) / 会議ID(テキスト) / ' +
    '会議コード(テキスト) / 会議記録ID(テキスト) / MeetURL(URL) / 実績打刻済み(チェックボックス) / ' +
    '議事録URL(URL) / 録画URL(URL)\n' +
    '5. データベースのURLからIDを取得し、「設定」シートの「Notion会議データベースID」に入力する\n\n' +
    '※打刻対象メンバー(氏名カンマ区切り)には、⏱ 活動タイマー側の「メンバー_◯◯」シート名(氏名部分)と' +
    '一致する名前を入力してください。実際の打刻はMeetの参加ログに基づき会議終了後に自動で行われます。\n\n' +
    '設定が終わったら「② 定期実行トリガーを設定」を実行してください。'
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
