/**
 * 会議終了後、Meetの実際の参加ログ(入室〜退室時刻)を使って、
 * (a) 実際に参加したメンバーだけを実績の時間で活動タイマーに打刻し、
 * (b) 出席管理DBの各行(出席者)に 出席 / 欠席 を自動反映する。
 *
 * 「出席管理DB」の行に載っていても、Meetに参加していなければ
 * 活動タイマーへの打刻は行われず、ステータスは「欠席」になる。
 *
 * 参加者はGoogleアカウントのリソースID(users/xxx)でしか識別できないため、
 * 「ユーザーID対応表」シートの「メールアドレス」「氏名(活動タイマー用)」列で
 * 対応付けが必要。未登録のIDは自動で行が追加されるので、管理者は
 * 該当欄を一度埋めればよい(電話参加・匿名参加などGoogleアカウントを
 * 持たない参加者は自動では対応付けられない)。
 *
 * トリガー: 10分おき(setupTriggers参照)。会議記録がまだ準備できていない場合は
 * 何もせず、次回実行時に自動的に再試行する。
 */
function stampActualAttendanceForFinishedMeetings() {
  const pages = queryMeetingDatabase({
    and: [
      { property: MEETING_PROPS.STAMPED, checkbox: { equals: false } },
      {
        or: [
          { property: MEETING_PROPS.STATUS, select: { equals: '終了' } },
          { property: MEETING_PROPS.STATUS, select: { equals: '投稿済み' } },
          { property: MEETING_PROPS.STATUS, select: { equals: '分析済み' } },
        ],
      },
    ],
  });

  pages.forEach(page => {
    const title = notionTitleText(page, MEETING_PROPS.TITLE);
    const meetingCode = notionRichText(page, MEETING_PROPS.MEETING_CODE);
    const attendanceRows = getAttendanceRowsForMeeting(page.id);

    if (attendanceRows.length === 0) {
      updatePage(page.id, propCheckbox(MEETING_PROPS.STAMPED, true));
      return;
    }

    let record;
    try {
      record = getConferenceRecord(meetingCode);
    } catch (e) {
      Logger.log(`会議記録の取得に失敗しました(${title}): ${e}`);
      return;
    }
    if (!record) return; // まだ会議記録が生成されていない。次回に再試行

    const sessions = getMeetParticipantSessions(record.name);
    const attendedEmails = new Set(
      sessions.map(s => resolveUserEmail(s.userResource)).filter(Boolean)
    );

    // (a) 活動タイマーへの実績打刻: 出席管理DBに載っている氏名と一致した人だけ
    const targetNames = attendanceRows.map(getAttendeeName).filter(Boolean);
    const stampEntries = sessions
      .map(s => ({ name: resolveMemberName(s.userResource), start: s.start, end: s.end }))
      .filter(s => s.name && targetNames.includes(s.name));

    if (stampEntries.length > 0) {
      const results = stampExactForMembers(stampEntries);
      Logger.log(`実績打刻(${title}): ${results.join(' / ')}`);
    }

    // (b) 出席管理DBの各行に 出席/欠席 を反映
    attendanceRows.forEach(row => {
      const email = getAttendeeEmail(row);
      const attended = !!email && attendedEmails.has(email);
      try {
        updatePage(row.id, propSelect(ATTENDANCE_PROPS.STATUS, attended ? '出席' : '欠席'));
      } catch (e) {
        Logger.log(`出席ステータスの更新に失敗しました(${title} / ${getAttendeeName(row)}): ${e}`);
      }
    });

    updatePage(page.id, propCheckbox(MEETING_PROPS.STAMPED, true));
  });
}
