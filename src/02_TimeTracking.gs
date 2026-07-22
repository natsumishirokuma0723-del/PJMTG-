/**
 * Notion会議データベースのステータスを時刻に応じて進める(作成済み→開催中→終了)。
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
  const pages = queryMeetingDatabase({
    property: NOTION_PROPS.STATUS,
    select: { equals: '作成済み' },
  });
  const now = new Date();

  pages.forEach(page => {
    const start = notionDate(page, NOTION_PROPS.START);
    if (start && now >= start) {
      updateMeetingPage(page.id, propSelect(NOTION_PROPS.STATUS, '開催中'));
    }
  });
}

function advanceOngoingToFinished() {
  const pages = queryMeetingDatabase({
    property: NOTION_PROPS.STATUS,
    select: { equals: '開催中' },
  });
  const now = new Date();

  pages.forEach(page => {
    const end = notionDate(page, NOTION_PROPS.END);
    if (end && now >= end) {
      updateMeetingPage(page.id, propSelect(NOTION_PROPS.STATUS, '終了'));
    }
  });
}

/**
 * Notionページから、打刻対象メンバー(氏名)の配列を取り出す。
 * 07_AttendanceStamping.gs でも使用する。
 */
function getStampTargetMembers(page) {
  return notionRichText(page, NOTION_PROPS.STAMP_MEMBERS)
    .split(',').map(s => s.trim()).filter(Boolean);
}
