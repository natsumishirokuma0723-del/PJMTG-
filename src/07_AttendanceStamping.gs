/**
 * 会議終了後、Meetの実際の参加ログ(入室〜退室時刻)を使って、
 * 実際に参加したメンバーだけを実績の時間で打刻する。
 * 「打刻対象メンバー」に名前があっても、Meetに参加していなければ打刻されない。
 *
 * 参加者はGoogleアカウントのリソースID(users/xxx)でしか識別できないため、
 * 「ユーザーID対応表」シートの「氏名(活動タイマー用)」列で氏名への対応付けが必要。
 * 未登録のIDは自動で行が追加されるので、管理者は氏名欄を一度埋めればよい
 * (電話参加・匿名参加などGoogleアカウントを持たない参加者は自動では対応付けられない)。
 *
 * トリガー: 10分おき(setupTriggers参照)。会議記録がまだ準備できていない場合は
 * 何もせず、次回実行時に自動的に再試行する。
 */
function stampActualAttendanceForFinishedMeetings() {
  const sheet = getSheet(SHEET_NAMES.SCHEDULE);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  const col = colIndexMap(data[0]);

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const status = row[col['ステータス']];
    if (status !== '終了' && status !== '投稿済み' && status !== '分析済み') continue;
    if (row[col['実績打刻済み']] === true) continue;

    const title = row[col['会議名']];
    const meetingCode = row[col['会議コード']];
    const targetMembers = getStampTargetMembers(row, col);
    const rowNum = r + 1;

    if (targetMembers.length === 0) {
      sheet.getRange(rowNum, col['実績打刻済み'] + 1).setValue(true);
      continue;
    }

    let record;
    try {
      record = getConferenceRecord(meetingCode);
    } catch (e) {
      Logger.log(`会議記録の取得に失敗しました(${title}): ${e}`);
      continue;
    }
    if (!record) continue; // まだ会議記録が生成されていない。次回に再試行

    const sessions = getMeetParticipantSessions(record.name);
    const entries = sessions
      .map(s => ({ name: resolveMemberName(s.userResource), start: s.start, end: s.end }))
      .filter(s => s.name && targetMembers.includes(s.name));

    if (entries.length > 0) {
      const results = stampExactForMembers(entries);
      Logger.log(`実績打刻(${title}): ${results.join(' / ')}`);
    } else {
      Logger.log(`実績打刻(${title}): 参加ログと一致するメンバーがいませんでした`);
    }

    sheet.getRange(rowNum, col['実績打刻済み'] + 1).setValue(true);
  }
}
