/**
 * 「会議予定」シートに登録された、まだ会議IDが空の行から
 * Google カレンダーの予定(Meet会議付き)を作成する。
 * 作成と同時に自動録画・自動文字起こし・スマートノートをONにする。
 *
 * トリガー: 15分おき(setupTriggers参照)
 */
function createScheduledMeetings() {
  const sheet = getSheet(SHEET_NAMES.SCHEDULE);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  const col = colIndexMap(data[0]);

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (row[col['会議ID']]) continue; // 作成済み
    if (!row[col['開始日時']] || !row[col['終了日時']]) continue;

    const title = row[col['会議名']];
    const start = new Date(row[col['開始日時']]);
    const end = new Date(row[col['終了日時']]);
    const attendeesRaw = row[col['出席者(カンマ区切りメール)']] || '';
    const attendees = String(attendeesRaw)
      .split(',')
      .map(s => s.trim())
      .filter(Boolean);

    let event;
    try {
      event = insertCalendarEventWithMeet(title, start, end, attendees);
    } catch (e) {
      Logger.log(`会議作成に失敗しました(${title}): ${e}`);
      continue;
    }

    const meetingCode = event.conferenceData && event.conferenceData.conferenceId;
    const meetUrl = getMeetUrl(event);

    if (meetingCode) {
      try {
        configureMeetArtifacts(meetingCode);
      } catch (e) {
        Logger.log(`議事録/録画の自動設定に失敗しました(${title}): ${e}`);
      }
    }

    const rowNum = r + 1;
    sheet.getRange(rowNum, col['会議ID'] + 1).setValue(event.id);
    sheet.getRange(rowNum, col['会議コード'] + 1).setValue(meetingCode || '');
    sheet.getRange(rowNum, col['MeetURL'] + 1).setValue(meetUrl || '');
    sheet.getRange(rowNum, col['ステータス'] + 1).setValue('作成済み');
  }
}

function insertCalendarEventWithMeet(title, start, end, attendeeEmails) {
  const calendarId = getConfig('カレンダーID') || 'primary';
  const requestId = Utilities.getUuid();

  const event = {
    summary: title,
    start: { dateTime: start.toISOString(), timeZone: Session.getScriptTimeZone() },
    end: { dateTime: end.toISOString(), timeZone: Session.getScriptTimeZone() },
    attendees: attendeeEmails.map(email => ({ email })),
    conferenceData: {
      createRequest: {
        requestId: requestId,
        conferenceSolutionKey: { type: 'hangoutsMeet' },
      },
    },
  };

  // Advanced Calendar Service (appsscript.json の enabledAdvancedServices) を使用
  return Calendar.Events.insert(event, calendarId, {
    conferenceDataVersion: 1,
    sendUpdates: 'all',
  });
}

function getMeetUrl(event) {
  if (!event.conferenceData || !event.conferenceData.entryPoints) return '';
  const videoEntry = event.conferenceData.entryPoints.find(e => e.entryPointType === 'video');
  return videoEntry ? videoEntry.uri : '';
}
