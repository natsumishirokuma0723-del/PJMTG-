/**
 * 「Chat投稿ログ」に記録された議事録メッセージへのリアクション(スタンプ)を収集し、
 * 「リアクション集計」シートに記録する(=議事録を既読した人とみなす)。
 *
 * トリガー: 1時間おき(setupTriggers参照)
 */
function collectReactionsForPostedMessages() {
  const chatLogSheet = getSheet(SHEET_NAMES.CHAT_LOG);
  const data = chatLogSheet.getDataRange().getValues();
  if (data.length < 2) return;
  const col = colIndexMap(data[0]);
  const reactionSheet = getSheet(SHEET_NAMES.REACTIONS);

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const eventId = row[col['会議ID']];
    const messageName = row[col['MessageID']];
    if (!messageName) continue;

    let reactions;
    try {
      reactions = listReactions(messageName);
    } catch (e) {
      Logger.log(`リアクション取得に失敗しました(${eventId}): ${e}`);
      continue;
    }

    reactions.forEach(reaction => {
      const userResource = reaction.user && reaction.user.name;
      if (!userResource) return;
      recordReaction(reactionSheet, eventId, userResource, reaction.emoji, reaction.createTime);
    });
  }
}

function listReactions(messageName) {
  const url = `${CHAT_API_BASE}/${messageName}/reactions`;
  const options = {
    method: 'get',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  };

  const res = UrlFetchApp.fetch(url, options);
  if (res.getResponseCode() >= 300) {
    throw new Error(`Chat API エラー (${res.getResponseCode()}): ${res.getContentText()}`);
  }
  const json = JSON.parse(res.getContentText());
  return json.reactions || [];
}

function recordReaction(sheet, eventId, userResource, emoji, createTime) {
  const data = sheet.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (data[r][0] === eventId && data[r][1] === userResource) return; // 重複防止
  }

  const emojiStr = (emoji && (emoji.unicode || (emoji.customEmoji && emoji.customEmoji.uid))) || '';
  sheet.appendRow([eventId, userResource, emojiStr, createTime ? new Date(createTime) : new Date()]);
}
