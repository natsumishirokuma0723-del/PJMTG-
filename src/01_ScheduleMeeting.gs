/**
 * 議事録DBに登録された、まだ会議IDが空で日付(開始・終了とも)が入っているページから
 * Google カレンダーの予定(Meet会議付き)を作成する。
 * 出席者は、そのページに紐づく出席管理DBの行(の「メンバー」リレーション先のメールアドレス)から集める。
 * 作成と同時に自動録画・自動文字起こし・スマートノートをONにする。
 *
 * トリガー: 15分おき(setupTriggers参照)
 */
function createScheduledMeetings() {
  const pages = queryMeetingDatabase({
    and: [
      { property: MEETING_PROPS.EVENT_ID, rich_text: { is_empty: true } },
      { property: MEETING_PROPS.DATE, date: { is_not_empty: true } },
    ],
  });

  pages.forEach(page => {
    const title = notionTitleText(page, MEETING_PROPS.TITLE);
    const { start, end } = notionDateRange(page, MEETING_PROPS.DATE);
    if (!start || !end || start.getTime() === end.getTime()) {
      Logger.log(`「${title}」は開始/終了時刻が不足しているためスキップします。「日付」プロパティに終了日時まで入力してください。`);
      return;
    }

    const attendanceRows = getAttendanceRowsForMeeting(page.id);
    const attendeeEmails = attendanceRows.map(getAttendeeEmail).filter(Boolean);

    let event;
    try {
      event = insertCalendarEventWithMeet(title, start, end, attendeeEmails);
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

    updatePage(page.id, Object.assign(
      {},
      propRichText(MEETING_PROPS.EVENT_ID, event.id),
      propRichText(MEETING_PROPS.MEETING_CODE, meetingCode || ''),
      propUrl(MEETING_PROPS.MEET_URL, meetUrl || ''),
      propSelect(MEETING_PROPS.STATUS, '作成済み')
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
