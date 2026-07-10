/**
 * スプレッドシートのカスタムメニュー、および定期実行トリガーの一括設定。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('会議自動化')
    .addItem('① 初期化(シート作成)', 'initializeSheets')
    .addItem('② 定期実行トリガーを設定', 'setupTriggers')
    .addSeparator()
    .addItem('会議を作成', 'createScheduledMeetings')
    .addItem('打刻を実行', 'stampAllMeetingsStatus')
    .addItem('Chatへ投稿', 'postFinishedMeetingsToChat')
    .addItem('リアクションを収集', 'collectReactionsForPostedMessages')
    .addItem('参加率を分析', 'analyzeAttendance')
    .addSeparator()
    .addItem('一連の処理を今すぐ全部実行', 'runFullPipeline')
    .addToUi();
}

/**
 * 各処理を定期実行するトリガーを設定する。
 * 既存のトリガーは一旦すべて削除してから作り直す(重複登録防止)。
 */
function setupTriggers() {
  deleteAllTriggers();

  ScriptApp.newTrigger('createScheduledMeetings').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('stampAllMeetingsStatus').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('postFinishedMeetingsToChat').timeBased().everyMinutes(10).create();
  ScriptApp.newTrigger('collectReactionsForPostedMessages').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('analyzeAttendance').timeBased().everyHours(1).create();

  SpreadsheetApp.getUi().alert('定期実行トリガーを設定しました。');
}

function deleteAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
}

function runFullPipeline() {
  createScheduledMeetings();
  stampAllMeetingsStatus();
  postFinishedMeetingsToChat();
  collectReactionsForPostedMessages();
  analyzeAttendance();
}
