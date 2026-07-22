/**
 * 会議終了(ステータス=終了)した会議について、Meet APIで会議記録
 * (議事録=スマートノート/文字起こしDoc、録画)が揃い次第、
 * Notionページ(プロパティ+本文)に保存し、Google Chatにも投稿する。
 * まだ生成が完了していない場合は何もせず、次回のトリガー実行時に再試行する。
 *
 * トリガー: 10分おき(setupTriggers参照)
 */
const CHAT_API_BASE = 'https://chat.googleapis.com/v1';

function postFinishedMeetingsToChat() {
  const pages = queryMeetingDatabase({
    property: NOTION_PROPS.STATUS,
    select: { equals: '終了' },
  });
  const chatLogSheet = getSheet(SHEET_NAMES.CHAT_LOG);

  pages.forEach(page => {
    const eventId = notionRichText(page, NOTION_PROPS.EVENT_ID);
    const title = notionTitleText(page, NOTION_PROPS.TITLE);
    const meetingCode = notionRichText(page, NOTION_PROPS.MEETING_CODE);

    let record;
    try {
      record = getConferenceRecord(meetingCode);
    } catch (e) {
      Logger.log(`会議記録の取得に失敗しました(${title}): ${e}`);
      return;
    }
    if (!record) return; // まだ会議記録が生成されていない。次回に再試行

    const recordingUrl = getRecordingUrl(record.name);
    const transcriptUrl = getTranscriptUrl(record.name);
    if (!recordingUrl && !transcriptUrl) return; // 生成中。次回に再試行

    const message = buildChatMessage(title, transcriptUrl, recordingUrl);
    let posted;
    try {
      posted = sendChatMessage(message);
    } catch (e) {
      Logger.log(`Chatへの投稿に失敗しました(${title}): ${e}`);
      return;
    }

    // Notionページ側にも議事録・録画リンクを保存する(プロパティ + 本文)
    updateMeetingPage(page.id, Object.assign(
      {},
      propRichText(NOTION_PROPS.CONFERENCE_RECORD_ID, record.name),
      propUrl(NOTION_PROPS.MINUTES_URL, transcriptUrl || ''),
      propUrl(NOTION_PROPS.RECORDING_URL, recordingUrl || ''),
      propSelect(NOTION_PROPS.STATUS, '投稿済み')
    ));
    try {
      appendParagraphs(page.id, [
        `議事録: ${transcriptUrl || '準備中'}`,
        `録画: ${recordingUrl || '準備中'}`,
      ]);
    } catch (e) {
      Logger.log(`Notionページへの本文追記に失敗しました(${title}): ${e}`);
    }

    chatLogSheet.appendRow([eventId, posted.name, title, transcriptUrl, recordingUrl, new Date()]);
  });
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
