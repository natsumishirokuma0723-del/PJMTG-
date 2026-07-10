/**
 * 開始/終了時刻になった会議を検知し、社内活動タイマー(ActivityTimer.gs)の
 * メンバーシートに開始/終了打刻を行う。
 *
 * 打刻対象は「会議予定」シートの「打刻対象メンバー(氏名カンマ区切り)」列で指定する。
 * ここに入れる名前は、ActivityTimer側の「メンバー_◯◯」シート名(氏名部分)と
 * 完全に一致させること。カレンダー招待用の「出席者(カンマ区切りメール)」とは
 * 別管理になる(ActivityTimerはメールアドレスではなく氏名でメンバーを管理するため)。
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
      const members = getStampTargetMembers(row, col);
      if (members.length > 0) {
        const results = stampStartForMembers(members);
        Logger.log(`打刻開始(${row[col['会議名']]}): ${results.join(' / ')}`);
      }
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
      const members = getStampTargetMembers(row, col);
      if (members.length > 0) {
        const results = stampEndForMembers(members);
        Logger.log(`打刻終了(${row[col['会議名']]}): ${results.join(' / ')}`);
      }
      sheet.getRange(r + 1, col['ステータス'] + 1).setValue('終了');
    }
  }
}

function getStampTargetMembers(row, col) {
  const raw = row[col['打刻対象メンバー(氏名カンマ区切り)']] || '';
  return String(raw).split(',').map(s => s.trim()).filter(Boolean);
}
