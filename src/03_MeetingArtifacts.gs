/**
 * Google Meet REST API (v2) 呼び出しをまとめたファイル。
 * 議事録(スマートノート)・録画の自動生成ON、および終了後の会議記録の取得を行う。
 *
 * 前提: Google Workspace 側で録画/文字起こし/スマートノート(Gemini)機能が
 * 有効になっている必要がある(エディションにより利用可否が異なる)。
 */
const MEET_API_BASE = 'https://meet.googleapis.com/v2';

/**
 * 指定した会議コードのSpaceに対して、自動録画・自動文字起こし・
 * スマートノート(議事録)をONにする。
 * Meet APIはSpaceの識別子として space ID の代わりに会議コードも受け付ける。
 */
function configureMeetArtifacts(meetingCode) {
  const updateMask = [
    'config.artifactConfig.recordingConfig.autoRecordingGeneration',
    'config.artifactConfig.transcriptionConfig.autoTranscriptionGeneration',
    'config.artifactConfig.smartNotesConfig.autoSmartNotesGeneration',
  ].join(',');

  const url = `${MEET_API_BASE}/spaces/${meetingCode}?updateMask=${encodeURIComponent(updateMask)}`;

  const payload = {
    config: {
      artifactConfig: {
        recordingConfig: { autoRecordingGeneration: 'ON' },
        transcriptionConfig: { autoTranscriptionGeneration: 'ON' },
        smartNotesConfig: { autoSmartNotesGeneration: 'ON' },
      },
    },
  };

  return meetApiFetch(url, 'patch', payload);
}

/**
 * 会議終了後に生成される conferenceRecord (会議記録) を取得する。
 * 終了直後は record がまだ存在しないことがあるため、呼び出し側で
 * 見つからない場合は次回実行時に再試行する設計にしている。
 */
function getConferenceRecord(meetingCode) {
  if (!meetingCode) return null;
  const filter = encodeURIComponent(`space.meeting_code="${meetingCode}"`);
  const url = `${MEET_API_BASE}/conferenceRecords?filter=${filter}`;

  const res = meetApiFetch(url, 'get');
  const records = res.conferenceRecords || [];
  if (records.length === 0) return null;

  records.sort((a, b) => new Date(b.startTime).getTime() - new Date(a.startTime).getTime());
  return records[0];
}

function getRecordingUrl(conferenceRecordName) {
  const url = `${MEET_API_BASE}/${conferenceRecordName}/recordings`;
  const res = meetApiFetch(url, 'get');
  const recordings = res.recordings || [];
  if (recordings.length === 0) return '';
  const rec = recordings[0];
  return (rec.driveDestination && rec.driveDestination.exportUri) || '';
}

function getTranscriptUrl(conferenceRecordName) {
  const url = `${MEET_API_BASE}/${conferenceRecordName}/transcripts`;
  const res = meetApiFetch(url, 'get');
  const transcripts = res.transcripts || [];
  if (transcripts.length === 0) return '';
  const tr = transcripts[0];
  return (tr.docsDestination && tr.docsDestination.exportUri) || '';
}

/**
 * conferenceRecord 配下の実参加者一覧を、入室〜退室時刻つきで取得する。
 * カレンダーの出欠(RSVP)ではなく、実際にMeetに入室した人を表す。
 * signedinUser.user はユーザーリソース名(users/xxx)なので、呼び出し側で
 * resolveUserEmail() / resolveMemberName() に変換して使う。
 * 匿名参加・電話参加(signedinUserを持たない)は対象外。
 * 入退室を複数回行った場合は、最初の入室〜最後の退室の幅になる。
 */
function getMeetParticipantSessions(conferenceRecordName) {
  if (!conferenceRecordName) return [];
  const url = `${MEET_API_BASE}/${conferenceRecordName}/participants`;

  try {
    const res = meetApiFetch(url, 'get');
    const participants = res.participants || [];
    return participants
      .filter(p => p.signedinUser && p.signedinUser.user && p.earliestStartTime && p.latestEndTime)
      .map(p => ({
        userResource: p.signedinUser.user,
        start: new Date(p.earliestStartTime),
        end: new Date(p.latestEndTime),
      }));
  } catch (e) {
    Logger.log(`Meet参加ログの取得に失敗しました: ${e}`);
    return [];
  }
}

/**
 * 実参加者のメールアドレス一覧を取得する(06_AttendanceAnalysis.gsで使用)。
 */
function getMeetParticipantEmails(conferenceRecordName) {
  return getMeetParticipantSessions(conferenceRecordName)
    .map(s => resolveUserEmail(s.userResource))
    .filter(Boolean);
}

function meetApiFetch(url, method, payload) {
  const options = {
    method: method,
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  };
  if (payload) options.payload = JSON.stringify(payload);

  const res = UrlFetchApp.fetch(url, options);
  const code = res.getResponseCode();
  if (code >= 300) {
    throw new Error(`Meet API エラー (${code}): ${res.getContentText()}`);
  }
  const text = res.getContentText();
  return text ? JSON.parse(text) : {};
}
