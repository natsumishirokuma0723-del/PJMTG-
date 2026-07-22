/**
 * Notion会議データベースに登録された、まだ会議IDが空のページから
 * Google カレンダーの予定(Meet会議付き)を作成する。
 * 作成と同時に自動録画・自動文字起こし・スマートノートをONにする。
 *
 * トリガー: 15分おき(setupTriggers参照)
 */
function createScheduledMeetings() {
  const pages = queryMeetingDatabase({
    and: [
      { property: NOTION_PROPS.EVENT_ID, rich_text: { is_empty: true } },
      { property: NOTION_PROPS.START, date: { is_not_empty: true } },
      { property: NOTION_PROPS.END, date: { is_not_empty: true } },
    ],
  });

  pages.forEach(page => {
    const title = notionTitleText(page, NOTION_PROPS.TITLE);
    const start = notionDate(page, NOTION_PROPS.START);
    const end = notionDate(page, NOTION_PROPS.END);
    const attendees = notionRichText(page, NOTION_PROPS.ATTENDEES)
      .split(',').map(s => s.trim()).filter(Boolean);

    let event;
    try {
      event = insertCalendarEventWithMeet(title, start, end, attendees);
    } catch (e) {
      Logger.log(`会議作成に失敗しました(${title}): ${e}`);
      return;
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

    updateMeetingPage(page.id, Object.assign(
      {},
      propRichText(NOTION_PROPS.EVENT_ID, event.id),
      propRichText(NOTION_PROPS.MEETING_CODE, meetingCode || ''),
      propUrl(NOTION_PROPS.MEET_URL, meetUrl || ''),
      propSelect(NOTION_PROPS.STATUS, '作成済み')
    ));
  });
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
