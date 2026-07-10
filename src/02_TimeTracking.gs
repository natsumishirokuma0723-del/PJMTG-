/**
 * 開始/終了時刻になった会議を「打刻ログ」シートに打刻する。
 * カレンダーの開始/終了時刻をトリガーにするのではなく、
 * 数分おきの時間トリガーで「今、開始/終了すべき会議」をスキャンする方式。
 *
 * トリガー: 5分おき(setupTriggers参照)
 */
function stampAllMeetingsStatus() {
  stampStartForOngoingMeetings();
  stampEndForFinishedMeetings();
}

function stampStartForOngoingMeetings() {
  const sheet = getSheet(SHEET_NAMES.SCHEDULE);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  const col = colIndexMap(data[0]);
  const now = new Date();

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (row[col['ステータス']] !== '作成済み') continue;

    const start = new Date(row[col['開始日時']]);
    if (now >= start) {
      recordTimeStamp(row[col['会議ID']], row[col['会議名']], '開始');
      sheet.getRange(r + 1, col['ステータス'] + 1).setValue('開催中');
    }
  }
}

function stampEndForFinishedMeetings() {
  const sheet = getSheet(SHEET_NAMES.SCHEDULE);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  const col = colIndexMap(data[0]);
  const now = new Date();

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (row[col['ステータス']] !== '開催中') continue;

    const end = new Date(row[col['終了日時']]);
    if (now >= end) {
      recordTimeStamp(row[col['会議ID']], row[col['会議名']], '終了');
      sheet.getRange(r + 1, col['ステータス'] + 1).setValue('終了');
    }
  }
}

function recordTimeStamp(eventId, title, type) {
  const sheet = getSheet(SHEET_NAMES.TIME_LOG);
  const data = sheet.getDataRange().getValues();
  const col = colIndexMap(data[0]);

  let targetRow = -1;
  for (let r = 1; r < data.length; r++) {
    if (data[r][col['会議ID']] === eventId) {
      targetRow = r + 1;
      break;
    }
  }

  const now = new Date();

  if (type === '開始') {
    if (targetRow === -1) {
      sheet.appendRow([eventId, title, now, '', '']);
    } else {
      sheet.getRange(targetRow, col['開始打刻'] + 1).setValue(now);
    }
    return;
  }

  // type === '終了'
  if (targetRow === -1) {
    sheet.appendRow([eventId, title, '', now, '']);
    targetRow = sheet.getLastRow();
  } else {
    sheet.getRange(targetRow, col['終了打刻'] + 1).setValue(now);
  }

  const rowVals = sheet.getRange(targetRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const startVal = rowVals[col['開始打刻']];
  if (startVal) {
    const minutes = Math.round((now.getTime() - new Date(startVal).getTime()) / 60000);
    sheet.getRange(targetRow, col['実績時間(分)'] + 1).setValue(minutes);
  }
}
