/**
 * 会議のステータスを時刻に応じて進める(作成済み→開催中→終了)。
 * ここでは打刻は行わない。実際に参加した人だけを実績時間で打刻する処理は
 * 07_AttendanceStamping.gs の stampActualAttendanceForFinishedMeetings() が、
 * 会議終了後にMeetの参加ログを使って別途行う。
 *
 * トリガー: 5分おき(setupTriggers参照)
 */
function advanceMeetingStatuses() {
  advanceCreatedToOngoing();
  advanceOngoingToFinished();
}

function advanceCreatedToOngoing() {
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
      sheet.getRange(r + 1, col['ステータス'] + 1).setValue('開催中');
    }
  }
}

function advanceOngoingToFinished() {
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
      sheet.getRange(r + 1, col['ステータス'] + 1).setValue('終了');
    }
  }
}

/**
 * 「会議予定」シートの行から、打刻対象メンバー(氏名)の配列を取り出す。
 * 07_AttendanceStamping.gs でも使用する。
 */
function getStampTargetMembers(row, col) {
  const raw = row[col['打刻対象メンバー(氏名カンマ区切り)']] || '';
  return String(raw).split(',').map(s => s.trim()).filter(Boolean);
}
