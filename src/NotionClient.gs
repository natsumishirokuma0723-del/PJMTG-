/**
 * Notion API 呼び出しをまとめたファイル。3つのデータベースを使う。
 *
 * ■ 議事録DB(既存。後輩が作成済みのものをそのまま使う)
 *   既存プロパティ: MTG名(タイトル) / 日付 / 議事録作成者 / 種別 / ファシリテーター
 *   - 「日付」プロパティは、Notion側で「時刻を含む」「終了日を含む」をONにしておくこと
 *     (1つのDateプロパティのstart/endを、そのまま開始日時/終了日時として使う)
 *   追加が必要なプロパティ(自動化用。既存プロパティは変更しない):
 *     ステータス(セレクト: 作成済み/開催中/終了/投稿済み/分析済み) / 会議ID(テキスト) /
 *     会議コード(テキスト) / 会議記録ID(テキスト) / MeetURL(URL) /
 *     実績打刻済み(チェックボックス) / 議事録URL(URL) / 録画URL(URL)
 *
 * ■ 出席管理DB(既存。後輩が作成済みのものをそのまま使う)
 *   既存プロパティ: MTG(議事録DBへのリレーション) / 名前(タイトル) / ステータス(セレクト) / 参加チーム
 *   追加が必要なプロパティ:
 *     メンバー(下記メンバーマスタDBへのリレーション。1人1件)
 *   「ステータス」は自動化が 出席 / 欠席 のどちらかを書き込む
 *     (「出席(音声のみ)」の判定は行わない。既存の選択肢に残っていても未使用)
 *
 * ■ メンバーマスタDB(新規作成)
 *   プロパティ: 氏名(タイトル) / メールアドレス(Eメール型) / 計測ツール氏名(テキスト)
 *   社内の人を一度だけ登録しておき、出席管理DBの「メンバー」から都度選ぶ運用にする。
 *   「氏名」は表示用のフルネーム、「計測ツール氏名」は活動タイマー側の
 *   「メンバー_◯◯」シート名と完全一致させる名前(名字のみの人がいるなど、
 *   氏名とは表記が異なる場合があるため別プロパティにしている)。
 *
 * 事前準備:
 *  1. https://www.notion.so/my-integrations で内部インテグレーションを作成
 *  2. Apps Scriptの「プロジェクトの設定」→「スクリプトプロパティ」に
 *     NOTION_TOKEN としてインテグレーションのシークレットを保存
 *  3. 上記3つのデータベースすべてを、①のインテグレーションと共有(接続)する
 *  4. 各データベースのIDを「設定」シートの
 *     「議事録データベースID」「出席管理データベースID」「メンバーマスタデータベースID」に入力
 */
const NOTION_API_BASE = 'https://api.notion.com/v1';
const NOTION_VERSION = '2022-06-28';

const MEETING_PROPS = {
  TITLE: 'MTG名',
  DATE: '日付',
  STATUS: 'ステータス',
  EVENT_ID: '会議ID',
  MEETING_CODE: '会議コード',
  CONFERENCE_RECORD_ID: '会議記録ID',
  MEET_URL: 'MeetURL',
  STAMPED: '実績打刻済み',
  MINUTES_URL: '議事録URL',
  RECORDING_URL: '録画URL',
};

const ATTENDANCE_PROPS = {
  NAME: '名前',
  MEETING_RELATION: 'MTG',
  MEMBER_RELATION: 'メンバー',
  STATUS: 'ステータス',
};

const MEMBER_PROPS = {
  NAME: '氏名',
  EMAIL: 'メールアドレス',
  STAMP_NAME: '計測ツール氏名',
};

function getNotionToken() {
  const token = PropertiesService.getScriptProperties().getProperty('NOTION_TOKEN');
  if (!token) throw new Error('スクリプトプロパティに NOTION_TOKEN が設定されていません。');
  return token;
}

function getMeetingDatabaseId() {
  const id = getConfig('議事録データベースID');
  if (!id) throw new Error('「設定」シートに 議事録データベースID が設定されていません。');
  return id;
}

function getAttendanceDatabaseId() {
  const id = getConfig('出席管理データベースID');
  if (!id) throw new Error('「設定」シートに 出席管理データベースID が設定されていません。');
  return id;
}

function getMemberDatabaseId() {
  const id = getConfig('メンバーマスタデータベースID');
  if (!id) throw new Error('「設定」シートに メンバーマスタデータベースID が設定されていません。');
  return id;
}

function notionFetch(path, method, payload) {
  const options = {
    method: method,
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + getNotionToken(),
      'Notion-Version': NOTION_VERSION,
    },
    muteHttpExceptions: true,
  };
  if (payload) options.payload = JSON.stringify(payload);

  const res = UrlFetchApp.fetch(`${NOTION_API_BASE}${path}`, options);
  const code = res.getResponseCode();
  if (code >= 300) {
    throw new Error(`Notion API エラー (${code}): ${res.getContentText()}`);
  }
  const text = res.getContentText();
  return text ? JSON.parse(text) : {};
}

/**
 * 指定したデータベースを検索する(ページング自動処理・全件配列で返す)。
 */
function queryDatabase(databaseId, filter) {
  const pages = [];
  let cursor = null;

  do {
    const payload = {};
    if (filter) payload.filter = filter;
    if (cursor) payload.start_cursor = cursor;

    const res = notionFetch(`/databases/${databaseId}/query`, 'post', payload);
    pages.push(...(res.results || []));
    cursor = res.has_more ? res.next_cursor : null;
  } while (cursor);

  return pages;
}

function queryMeetingDatabase(filter) {
  return queryDatabase(getMeetingDatabaseId(), filter);
}

function queryAttendanceDatabase(filter) {
  return queryDatabase(getAttendanceDatabaseId(), filter);
}

function getPage(pageId) {
  return notionFetch(`/pages/${pageId}`, 'get');
}

function updatePage(pageId, properties) {
  return notionFetch(`/pages/${pageId}`, 'patch', { properties });
}

/**
 * ページ本文の末尾に段落ブロックを追記する(議事録リンクの記録用)。
 */
function appendParagraphs(pageId, lines) {
  const children = lines.map(line => ({
    object: 'block',
    type: 'paragraph',
    paragraph: {
      rich_text: [{ type: 'text', text: { content: line } }],
    },
  }));
  return notionFetch(`/blocks/${pageId}/children`, 'patch', { children });
}

// ─── プロパティ読み取りヘルパー ───

function notionTitleText(page, propName) {
  const prop = page.properties[propName];
  if (!prop || !prop.title || prop.title.length === 0) return '';
  return prop.title.map(t => t.plain_text).join('');
}

function notionRichText(page, propName) {
  const prop = page.properties[propName];
  if (!prop || !prop.rich_text || prop.rich_text.length === 0) return '';
  return prop.rich_text.map(t => t.plain_text).join('');
}

function notionEmail(page, propName) {
  const prop = page.properties[propName];
  return (prop && prop.email) || '';
}

/**
 * Dateプロパティ(1つでstart/endを両方持つ設定にしたもの)から
 * { start: Date|null, end: Date|null } を取り出す。
 * endが未設定(単一日時のみ)の場合は end=start として扱う。
 */
function notionDateRange(page, propName) {
  const prop = page.properties[propName];
  if (!prop || !prop.date || !prop.date.start) return { start: null, end: null };
  const start = new Date(prop.date.start);
  const end = prop.date.end ? new Date(prop.date.end) : start;
  return { start, end };
}

function notionCheckbox(page, propName) {
  const prop = page.properties[propName];
  return !!(prop && prop.checkbox);
}

function notionRelationIds(page, propName) {
  const prop = page.properties[propName];
  if (!prop || !prop.relation) return [];
  return prop.relation.map(r => r.id);
}

// ─── プロパティ書き込みヘルパー ───
// updatePage() の properties にそのまま渡せる形(1キーのオブジェクト)を返す。
// 複数まとめて渡すときは Object.assign({}, propX(...), propY(...)) のように合成する。

function propRichText(propName, value) {
  return { [propName]: { rich_text: value ? [{ text: { content: String(value) } }] : [] } };
}

function propSelect(propName, value) {
  return { [propName]: { select: value ? { name: value } : null } };
}

function propUrl(propName, value) {
  return { [propName]: { url: value || null } };
}

function propCheckbox(propName, value) {
  return { [propName]: { checkbox: !!value } };
}

// ─── 出席管理DBまわりのヘルパー ───

/**
 * 指定した議事録DBページ(会議)に紐づく出席管理DBの行を全件取得する。
 */
function getAttendanceRowsForMeeting(meetingPageId) {
  return queryAttendanceDatabase({
    property: ATTENDANCE_PROPS.MEETING_RELATION,
    relation: { contains: meetingPageId },
  });
}

/**
 * 出席管理DBの行(1人分)における表示名(出席管理DB自体の「名前」列。ログ表示用)。
 * 活動タイマーとの氏名照合には使わない(→ getAttendeeInfo().stampName を使う)。
 */
function getAttendeeName(attendanceRow) {
  return notionTitleText(attendanceRow, ATTENDANCE_PROPS.NAME);
}

/**
 * 出席管理DBの行(1人分)から、「メンバー」リレーション先(メンバーマスタDB)を
 * たどってメールアドレス・活動タイマー用氏名(計測ツール氏名)をまとめて取得する。
 */
function getAttendeeInfo(attendanceRow) {
  const memberIds = notionRelationIds(attendanceRow, ATTENDANCE_PROPS.MEMBER_RELATION);
  if (memberIds.length === 0) return { email: '', stampName: '' };
  try {
    const memberPage = getPage(memberIds[0]);
    return {
      email: notionEmail(memberPage, MEMBER_PROPS.EMAIL),
      stampName: notionRichText(memberPage, MEMBER_PROPS.STAMP_NAME),
    };
  } catch (e) {
    Logger.log(`メンバー情報の取得に失敗しました: ${e}`);
    return { email: '', stampName: '' };
  }
}

function getAttendeeEmail(attendanceRow) {
  return getAttendeeInfo(attendanceRow).email;
}
