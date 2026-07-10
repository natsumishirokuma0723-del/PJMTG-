/**
 * 会議終了(ステータス=終了)した会議について、Meet APIで会議記録
 * (議事録=スマートノート/文字起こしDoc、録画)が揃い次第 Google Chat に投稿する。
 * まだ生成が完了していない場合は何もせず、次回のトリガー実行時に再試行する。
 *
 * トリガー: 10分おき(setupTriggers参照)
 */
const CHAT_API_BASE = 'https://chat.googleapis.com/v1';

function postFinishedMeetingsToChat() {
  const sheet = getSheet(SHEET_NAMES.SCHEDULE);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  const col = colIndexMap(data[0]);
  const chatLogSheet = getSheet(SHEET_NAMES.CHAT_LOG);

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (row[col['ステータス']] !== '終了') continue;

    const eventId = row[col['会議ID']];
    const title = row[col['会議名']];
    const meetingCode = row[col['会議コード']];

    let record;
    try {
      record = getConferenceRecord(meetingCode);
    } catch (e) {
      Logger.log(`会議記録の取得に失敗しました(${title}): ${e}`);
      continue;
    }
    if (!record) continue; // まだ会議記録が生成されていない。次回に再試行

    const recordingUrl = getRecordingUrl(record.name);
    const transcriptUrl = getTranscriptUrl(record.name);
    if (!recordingUrl && !transcriptUrl) continue; // 生成中。次回に再試行

    const message = buildChatMessage(title, transcriptUrl, recordingUrl);
    let posted;
    try {
      posted = sendChatMessage(message);
    } catch (e) {
      Logger.log(`Chatへの投稿に失敗しました(${title}): ${e}`);
      continue;
    }

    const rowNum = r + 1;
    sheet.getRange(rowNum, col['会議記録ID'] + 1).setValue(record.name);
    sheet.getRange(rowNum, col['ステータス'] + 1).setValue('投稿済み');

    chatLogSheet.appendRow([eventId, posted.name, title, transcriptUrl, recordingUrl, new Date()]);
  }
}

function buildChatMessage(title, transcriptUrl, recordingUrl) {
  return [
    `📋 *${title}* の議事録・録画が準備できました。`,
    transcriptUrl ? `議事録: ${transcriptUrl}` : '議事録: 準備中',
    recordingUrl ? `録画: ${recordingUrl}` : '録画: 準備中',
    '内容を確認したら、このメッセージにリアクション(スタンプ)を押してください。',
  ].join('\n');
}

function sendChatMessage(text) {
  const spaceId = getConfig('ChatスペースID');
  if (!spaceId) throw new Error('「設定」シートに ChatスペースID が設定されていません。');

  const url = `${CHAT_API_BASE}/${spaceId}/messages`;
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify({ text: text }),
    muteHttpExceptions: true,
  };

  const res = UrlFetchApp.fetch(url, options);
  if (res.getResponseCode() >= 300) {
    throw new Error(`Chat API エラー (${res.getResponseCode()}): ${res.getContentText()}`);
  }
  return JSON.parse(res.getContentText());
}
