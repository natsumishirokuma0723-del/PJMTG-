/**
 * Notion API (会議管理データベース) 呼び出しをまとめたファイル。
 * 「会議予定」はGoogle Sheetsではなく、このNotionデータベースが正になる。
 *
 * 事前準備:
 *  1. https://www.notion.so/my-integrations で内部インテグレーションを作成
 *  2. Apps Scriptの「プロジェクトの設定」→「スクリプトプロパティ」に
 *     NOTION_TOKEN としてインテグレーションのシークレットを保存
 *  3. 会議管理用データベースを作成し、①のインテグレーションと共有(接続)する
 *  4. データベースのIDを「設定」シートの「Notion会議データベースID」に入力
 *
 * データベースに必要なプロパティ(名前・型を一致させること):
 *   会議名(タイトル) / 開始日時(日付) / 終了日時(日付) / 出席者メール(テキスト) /
 *   打刻対象メンバー(テキスト) / ステータス(セレクト) / 会議ID(テキスト) /
 *   会議コード(テキスト) / 会議記録ID(テキスト) / MeetURL(URL) /
 *   実績打刻済み(チェックボックス) / 議事録URL(URL) / 録画URL(URL)
 */
const NOTION_API_BASE = 'https://api.notion.com/v1';
const NOTION_VERSION = '2022-06-28';

const NOTION_PROPS = {
  TITLE: '会議名',
  START: '開始日時',
  END: '終了日時',
  ATTENDEES: '出席者メール',
  STAMP_MEMBERS: '打刻対象メンバー',
  EVENT_ID: '会議ID',
  MEETING_CODE: '会議コード',
  CONFERENCE_RECORD_ID: '会議記録ID',
  MEET_URL: 'MeetURL',
  STATUS: 'ステータス',
  STAMPED: '実績打刻済み',
  MINUTES_URL: '議事録URL',
  RECORDING_URL: '録画URL',
};

function getNotionToken() {
  const token = PropertiesService.getScriptProperties().getProperty('NOTION_TOKEN');
  if (!token) throw new Error('スクリプトプロパティに NOTION_TOKEN が設定されていません。');
  return token;
}

function getNotionDatabaseId() {
  const id = getConfig('Notion会議データベースID');
  if (!id) throw new Error('「設定」シートに Notion会議データベースID が設定されていません。');
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
 * 会議データベースを検索する。filterはNotion API仕様のfilterオブジェクト(省略可)。
 * ページングを内部で処理し、条件に合う全ページを配列で返す。
 */
function queryMeetingDatabase(filter) {
  const databaseId = getNotionDatabaseId();
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

function updateMeetingPage(pageId, properties) {
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

function notionDate(page, propName) {
  const prop = page.properties[propName];
  if (!prop || !prop.date || !prop.date.start) return null;
  return new Date(prop.date.start);
}

function notionCheckbox(page, propName) {
  const prop = page.properties[propName];
  return !!(prop && prop.checkbox);
}

// ─── プロパティ書き込みヘルパー ───
// それぞれ updateMeetingPage() の properties にそのまま渡せる形(1キーのオブジェクト)を返す。
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
