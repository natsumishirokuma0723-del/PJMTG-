/**
 * 「投稿済み」の会議について、招待者(出席管理DBの行)・実参加者(Meet参加ログ)・
 * 既読者(リアクションを押した人)を突き合わせて出席率/既読率を算出し、
 * 「参加分析」シートに書き込む。
 *
 * トリガー: 1時間おき(setupTriggers参照)
 */
function analyzeAttendance() {
  const pages = queryMeetingDatabase({
    property: MEETING_PROPS.STATUS,
    select: { equals: '投稿済み' },
  });

  const reactionSheet = getSheet(SHEET_NAMES.REACTIONS);
  const reactionData = reactionSheet.getDataRange().getValues();
  const analysisSheet = getSheet(SHEET_NAMES.ANALYSIS);

  pages.forEach(page => {
    const eventId = notionRichText(page, MEETING_PROPS.EVENT_ID);
    const title = notionTitleText(page, MEETING_PROPS.TITLE);
    const conferenceRecordName = notionRichText(page, MEETING_PROPS.CONFERENCE_RECORD_ID);

    const attendanceRows = getAttendanceRowsForMeeting(page.id);
    const invitees = attendanceRows.map(getAttendeeEmail).filter(Boolean);

    const reactedEmails = new Set(
      reactionData
        .filter(rr => rr[0] === eventId)
        .map(rr => resolveUserEmail(rr[1]))
        .filter(Boolean)
    );

    let attendedEmails = new Set(getMeetParticipantEmails(conferenceRecordName));
    if (attendedEmails.size === 0) {
      // Meet参加ログが取得できない場合はカレンダーの出欠(RSVP)で代替する
      attendedEmails = new Set(getAcceptedFromCalendarEvent(eventId));
    }

    const totalInvited = invitees.length;
    const totalAttended = invitees.filter(e => attendedEmails.has(e)).length;
    const totalRead = invitees.filter(e => reactedEmails.has(e)).length;

    const notAttended = invitees.filter(e => !attendedEmails.has(e));
    const notRead = invitees.filter(e => !reactedEmails.has(e));

    const attendanceRate = totalInvited ? totalAttended / totalInvited : 0;
    const readRate = totalInvited ? totalRead / totalInvited : 0;

    analysisSheet.appendRow([
      eventId, title, totalInvited, totalAttended, attendanceRate,
      totalRead, readRate, notAttended.join(', '), notRead.join(', '), new Date(),
    ]);

    updatePage(page.id, propSelect(MEETING_PROPS.STATUS, '分析済み'));
  });
}

function getAcceptedFromCalendarEvent(eventId) {
  try {
    const event = Calendar.Events.get(getConfig('カレンダーID') || 'primary', eventId);
    const attendees = event.attendees || [];
    return attendees
      .filter(a => a.responseStatus === 'accepted')
      .map(a => a.email);
  } catch (e) {
    Logger.log(`カレンダー出欠取得に失敗しました: ${e}`);
    return [];
  }
}
