/**
 * 議事録DBのステータスを時刻に応じて進める(作成済み→開催中→終了)。
 * ここでは打刻は行わない。実際に参加した人だけを実績時間で打刻し、
 * 出席管理DBの出席/欠席を反映する処理は07_AttendanceStamping.gsが、
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
    property: MEETING_PROPS.STATUS,
    select: { equals: '作成済み' },
  });
  const now = new Date();

  pages.forEach(page => {
    const { start } = notionDateRange(page, MEETING_PROPS.DATE);
    if (start && now >= start) {
      updatePage(page.id, propSelect(MEETING_PROPS.STATUS, '開催中'));
    }
  });
}

function advanceOngoingToFinished() {
  const pages = queryMeetingDatabase({
    property: MEETING_PROPS.STATUS,
    select: { equals: '開催中' },
  });
  const now = new Date();

  pages.forEach(page => {
    const { end } = notionDateRange(page, MEETING_PROPS.DATE);
    if (end && now >= end) {
      updatePage(page.id, propSelect(MEETING_PROPS.STATUS, '終了'));
    }
  });
}
