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
  const pages = queryMeetingDatabase({
    and: [
      { property: NOTION_PROPS.STAMPED, checkbox: { equals: false } },
      {
        or: [
          { property: NOTION_PROPS.STATUS, select: { equals: '終了' } },
          { property: NOTION_PROPS.STATUS, select: { equals: '投稿済み' } },
          { property: NOTION_PROPS.STATUS, select: { equals: '分析済み' } },
        ],
      },
    ],
  });

  pages.forEach(page => {
    const title = notionTitleText(page, NOTION_PROPS.TITLE);
    const meetingCode = notionRichText(page, NOTION_PROPS.MEETING_CODE);
    const targetMembers = getStampTargetMembers(page);

    if (targetMembers.length === 0) {
      updateMeetingPage(page.id, propCheckbox(NOTION_PROPS.STAMPED, true));
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
    const entries = sessions
      .map(s => ({ name: resolveMemberName(s.userResource), start: s.start, end: s.end }))
      .filter(s => s.name && targetMembers.includes(s.name));

    if (entries.length > 0) {
      const results = stampExactForMembers(entries);
      Logger.log(`実績打刻(${title}): ${results.join(' / ')}`);
    } else {
      Logger.log(`実績打刻(${title}): 参加ログと一致するメンバーがいませんでした`);
    }

    updateMeetingPage(page.id, propCheckbox(NOTION_PROPS.STAMPED, true));
  });
}
