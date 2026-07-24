/**
 * スプレッドシートのメニューから最初に1回実行する初期化処理。
 * 必要なシートと見出し、デフォルト設定値を作成する。
 */
function initializeSheets() {
  const ss = getSs();

  createSheetIfMissing(ss, SHEET_NAMES.CONFIG, ['キー', '値'], [
    ['カレンダーID', 'primary'],
    ['ChatスペースID', 'spaces/xxxxxxxx'],
    ['議事録データベースID', ''],
    ['出席管理データベースID', ''],
    ['メンバーマスタデータベースID', ''],
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
    '「設定」シートの「ChatスペースID」「カレンダーID」' +
    '「議事録データベースID」「出席管理データベースID」「メンバーマスタデータベースID」を入力してください。\n\n' +
    '■ Notion側でやること(このスプレッドシートとは別に準備が必要)\n' +
    '1. https://www.notion.so/my-integrations で内部インテグレーションを作成し、シークレットを控える\n' +
    '2. Apps Scriptの「プロジェクトの設定」→「スクリプトプロパティ」に NOTION_TOKEN として保存する\n' +
    '3. 既存の「議事録DB」「出席管理DB」に、①のインテグレーションを接続(共有)する\n' +
    '4. 「議事録DB」の「日付」プロパティで「時刻を含む」「終了日を含む」をONにする\n' +
    '5. 「議事録DB」に次のプロパティを追加する(既存のプロパティは変更しない): ' +
    'ステータス(セレクト) / 会議ID(テキスト) / 会議コード(テキスト) / 会議記録ID(テキスト) / ' +
    'MeetURL(URL) / 実績打刻済み(チェックボックス) / 議事録URL(URL) / 録画URL(URL)\n' +
    '6. 新規に「メンバーマスタDB」を作成する: 氏名(タイトル) / メールアドレス(Eメール型) / ' +
    '計測ツール氏名(テキスト)。「計測ツール氏名」は活動タイマー側の「メンバー_◯◯」シート名と' +
    '完全一致させる名前で、「氏名」(表示名)とは別に管理する' +
    '(名字のみの人など表記が異なる場合があるため)。社内の人を一度だけ登録し、①のインテグレーションと接続する\n' +
    '7. 「出席管理DB」に「メンバー」(メンバーマスタDBへのリレーション)プロパティを追加し、' +
    '各会議の出席者行でこのメンバーマスタから人を選ぶ運用にする\n' +
    '8. 各データベースのURLからIDを取得し、「設定」シートの対応する項目に入力する\n\n' +
    '※出席管理DBの「ステータス」には、会議終了後にMeetの参加ログをもとに' +
    '「出席」または「欠席」が自動で書き込まれます。活動タイマーへの打刻も同じタイミングで、' +
    '実際にMeetへ参加した人(計測ツール氏名で照合)だけ行われます。\n\n' +
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
