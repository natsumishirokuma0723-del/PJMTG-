// ============================================================
//  社内活動 打刻タイマー — Google Apps Script 完全版
//  機能一覧：
//  ・メンバー選択シート（個人 / グループ）
//  ・グループ追加・編集・削除・グループ名変更
//  ・メンバー追加（複数人対応）・名前変更・削除
//  ・赤伝票・黒伝票方式の打刻修正
//  ・メンバーシート保護 / 解除（管理者）
//  ・年度リセット（パスワード認証＋バックアップ）
//  ・サマリー（ランキング形式＋毎日変わるコンテンツ）
//  ・スマホ打刻シート（チェックボックス＋onEditトリガー）
//  ・マニュアルシート自動生成
//
//  ※ 会議自動化ワークフロー(02_TimeTracking.gs)から
//    stampStartForMembers() / stampEndForMembers() を直接呼び出して
//    会議の開始/終了に合わせて自動打刻できるようにしている。
// ============================================================

const SUMMARY_SHEET_NAME = "サマリー";
const SELECT_SHEET_NAME  = "メンバー選択";
const MEMBER_PREFIX      = "メンバー_";
const BACKUP_PREFIX      = "バックアップ_";
const COL_DATE   = 1;
const COL_START  = 2;
const COL_END    = 3;
const COL_HOURS  = 4;
const COL_MEMO   = 5;
const COL_MANUAL = 6;
const DATA_START = 2;

// ─── メニュー生成 ─────────────────────────────────────
// 会議自動化(Main.gs)のonOpen()から呼び出される。
// このファイル単体では onOpen という名前を持たない(重複定義エラー回避のため)。
function buildActivityTimerMenu_() {
  SpreadsheetApp.getUi()
    .createMenu("⏱ 活動タイマー")
    .addItem("📋 メンバー選択シートを開く", "openSelectSheet")
    .addItem("▶ 開始打刻を実行", "stampStartSelected")
    .addItem("⏹ 終了打刻を実行", "stampEndSelected")
    .addSeparator()
    .addItem("👥 グループを追加", "addGroup")
    .addItem("✏️ グループを編集", "editGroup")
    .addItem("🔤 グループ名を変更", "renameGroup")
    .addItem("🗑 グループを削除", "deleteGroup")
    .addSeparator()
    .addItem("🔧 打刻を手動修正", "manualCorrect")
    .addItem("🔒 メンバーシートを保護（管理者）", "protectMemberSheets")
    .addItem("🔓 メンバーシートの保護解除（管理者）", "unprotectMemberSheets")
    .addItem("🔄 年度リセット（管理者）", "yearlyReset")
    .addItem("📊 サマリーを更新", "refreshSummary")
    .addSeparator()
    .addItem("⚙ 初期セットアップ", "runSetup")
    .addItem("➕ メンバーを追加", "addMember")
    .addItem("✏️ メンバー名を変更（管理者）", "renameMember")
    .addItem("🗑 メンバーを削除（管理者）", "deleteMember")
    .addSeparator()
    .addItem("📱 スマホ打刻シートを作成", "setupCheckboxSheet")
    .addItem("📖 マニュアルシートを作成", "createManualSheet")
    .addToUi();
}

// ─── パスワード認証 ───────────────────────────────────
function _authenticate(title) {
  const ui       = SpreadsheetApp.getUi();
  const password = PropertiesService.getScriptProperties().getProperty("ADMIN_PASSWORD");
  if (!password) {
    ui.alert("❌ パスワードが設定されていません。\nGASの「プロジェクトの設定 → スクリプトプロパティ」でADMIN_PASSWORDを設定してください。");
    return false;
  }
  const res = ui.prompt(title, "管理者パスワードを入力してください：", ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return false;
  if (res.getResponseText().trim() !== password) {
    ui.alert("❌ パスワードが正しくありません。");
    return false;
  }
  return true;
}

// ─── 初期セットアップ ─────────────────────────────────
function runSetup() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.prompt("初期セットアップ", "メンバー名をカンマ区切りで入力：\n例）田中, 鈴木, 佐藤", ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const names = res.getResponseText().split(",").map(n => n.trim()).filter(Boolean);
  if (names.length === 0) { ui.alert("名前が入力されていません。"); return; }
  names.forEach(name => _ensureMemberSheet(name));
  _ensureSummarySheet(names);
  _buildSelectSheet();
  refreshSummary();
  ui.alert(`✅ セットアップ完了！\n登録メンバー：${names.join("、")}`);
}

// ─── メンバー追加（複数人対応）──────────────────────
function addMember() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.prompt("メンバー追加", "追加するメンバー名をカンマ区切りで入力：\n例）山田, 伊藤, 中村", ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const names = res.getResponseText().split(",").map(n => n.trim()).filter(Boolean);
  if (names.length === 0) { ui.alert("名前が入力されていません。"); return; }
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const summary = ss.getSheetByName(SUMMARY_SHEET_NAME);
  names.forEach(name => {
    _ensureMemberSheet(name);
    if (summary) {
      const lastRow = summary.getLastRow();
      summary.getRange(lastRow + 1, 1).setValue(name);
      summary.getRange(lastRow + 1, 2).setValue(0).setNumberFormat("0.00");
      _formatSummaryRow(summary, lastRow + 1);
    }
  });
  const selectSh = ss.getSheetByName(SELECT_SHEET_NAME);
  if (selectSh) _refreshMemberColumn(selectSh);
  refreshSummary();
  ui.alert(`✅ ${names.length}人を追加しました！\n登録：${names.join("、")}`);
}

// ─── メンバー名変更 ───────────────────────────────────
function renameMember() {
  const ui      = SpreadsheetApp.getUi();
  if (!_authenticate("✏️ メンバー名変更")) return;
  const members = _getMemberNames();
  if (members.length === 0) { ui.alert("メンバーが登録されていません。"); return; }
  const memberList = members.map((name, i) => `${i + 1}：${name}`).join("\n");
  const r1 = ui.prompt("メンバー名変更 ①", `名前を変更するメンバーの番号を入力：\n\n${memberList}`, ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const num = parseInt(r1.getResponseText().trim());
  if (isNaN(num) || num < 1 || num > members.length) { ui.alert(`❌ 1〜${members.length} の番号を入力してください。`); return; }
  const oldName = members[num - 1];
  const r2 = ui.prompt("メンバー名変更 ②", `「${oldName}」の新しい名前を入力：`, ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const newName = r2.getResponseText().trim();
  if (!newName) { ui.alert("名前が入力されていません。"); return; }
  if (newName === oldName) { ui.alert("同じ名前です。"); return; }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(MEMBER_PREFIX + newName)) { ui.alert(`❌ 「${newName}」はすでに存在します。`); return; }
  const memberSh = ss.getSheetByName(MEMBER_PREFIX + oldName);
  if (memberSh) memberSh.setName(MEMBER_PREFIX + newName);
  const summary = ss.getSheetByName(SUMMARY_SHEET_NAME);
  if (summary) {
    const data = summary.getDataRange().getValues();
    for (let r = 1; r < data.length; r++) {
      if (data[r][0] === oldName) { summary.getRange(r + 1, 1).setValue(newName); break; }
    }
  }
  const selectSh = ss.getSheetByName(SELECT_SHEET_NAME);
  if (selectSh) {
    const groups = _getGroups(selectSh);
    groups.forEach(g => {
      const updated = g.members.map(m => m === oldName ? newName : m);
      if (JSON.stringify(updated) !== JSON.stringify(g.members)) {
        selectSh.getRange(g.shRow, 7).setValue(updated.join(","));
      }
    });
    _refreshMemberColumn(selectSh);
  }
  ui.alert(`✅ 「${oldName}」→「${newName}」に変更しました。\nグループ内の名前も自動更新されました。`);
}

// ─── メンバー削除 ─────────────────────────────────────
function deleteMember() {
  const ui      = SpreadsheetApp.getUi();
  if (!_authenticate("🗑 メンバー削除")) return;
  const members = _getMemberNames();
  if (members.length === 0) { ui.alert("メンバーが登録されていません。"); return; }
  const memberList = members.map((name, i) => `${i + 1}：${name}`).join("\n");
  const r1 = ui.prompt("メンバー削除", `削除するメンバーの番号を入力：\n\n${memberList}`, ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const num = parseInt(r1.getResponseText().trim());
  if (isNaN(num) || num < 1 || num > members.length) { ui.alert(`❌ 1〜${members.length} の番号を入力してください。`); return; }
  const name    = members[num - 1];
  const confirm = ui.alert("削除の確認", `「${name}」を削除します。\nこの操作は元に戻せません。よろしいですか？`, ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const memberSh = ss.getSheetByName(MEMBER_PREFIX + name);
  if (memberSh) ss.deleteSheet(memberSh);
  const summary = ss.getSheetByName(SUMMARY_SHEET_NAME);
  if (summary) {
    const data = summary.getDataRange().getValues();
    for (let r = 1; r < data.length; r++) {
      if (data[r][0] === name) { summary.deleteRow(r + 1); break; }
    }
  }
  const selectSh = ss.getSheetByName(SELECT_SHEET_NAME);
  if (selectSh) _refreshMemberColumn(selectSh);
  refreshSummary();
  ui.alert(`✅ 「${name}」を削除しました。`);
}

// ─── メンバー選択シートを開く ─────────────────────────
function openSelectSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SELECT_SHEET_NAME);
  if (!sh) { _buildSelectSheet(); } else { _refreshMemberColumn(sh); }
  ss.setActiveSheet(ss.getSheetByName(SELECT_SHEET_NAME));
}

// ─── 選択シート初期構築 ───────────────────────────────
function _buildSelectSheet() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  let   sh      = ss.getSheetByName(SELECT_SHEET_NAME);
  const members = _getMemberNames();
  if (!sh) { sh = ss.insertSheet(SELECT_SHEET_NAME); } else { sh.clearContents(); sh.clearFormats(); }
  sh.getRange("B1").setValue("👤 個人").setFontWeight("bold").setBackground("#e8f0fe").setHorizontalAlignment("center");
  sh.getRange("C1").setValue("名前").setFontWeight("bold").setBackground("#e8f0fe");
  members.forEach((name, i) => {
    const row = i + 2;
    sh.getRange(row, 2).insertCheckboxes().setValue(false);
    sh.getRange(row, 3).setValue(name).setFontSize(11);
    sh.getRange(row, 2, 1, 2).setBackground(row % 2 === 0 ? "#ffffff" : "#f8f9fa");
  });
  sh.getRange(1, 4).setValue("").setBackground("#dadce0");
  sh.getRange("E1").setValue("👥 グループ").setFontWeight("bold").setBackground("#fce8e6").setHorizontalAlignment("center");
  sh.getRange("F1").setValue("グループ名").setFontWeight("bold").setBackground("#fce8e6");
  sh.getRange("G1").setValue("メンバー").setFontWeight("bold").setBackground("#fce8e6");
  sh.setColumnWidth(2, 40); sh.setColumnWidth(3, 130);
  sh.setColumnWidth(4, 15); sh.setColumnWidth(5, 40);
  sh.setColumnWidth(6, 130); sh.setColumnWidth(7, 220);
  sh.setFrozenRows(1);
}

// ─── 個人列だけ再描画（グループ定義保持）────────────────
function _refreshMemberColumn(sh) {
  const members = _getMemberNames();
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) sh.getRange(2, 2, lastRow - 1, 2).clearContent().clearFormat();
  members.forEach((name, i) => {
    const row = i + 2;
    sh.getRange(row, 2).insertCheckboxes().setValue(false);
    sh.getRange(row, 3).setValue(name).setFontSize(11);
    sh.getRange(row, 2, 1, 2).setBackground(row % 2 === 0 ? "#ffffff" : "#f8f9fa");
  });
}

// ─── グループ一覧取得 ─────────────────────────────────
function _getGroups(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const data   = sh.getRange(2, 5, lastRow - 1, 3).getValues();
  const groups = [];
  data.forEach((row, i) => {
    if (row[1] && String(row[1]).trim() !== "") {
      groups.push({
        shRow:   i + 2,
        name:    String(row[1]).trim(),
        members: String(row[2] || "").split(",").map(n => n.trim()).filter(Boolean)
      });
    }
  });
  return groups;
}

// ─── グループ追加 ─────────────────────────────────────
function addGroup() {
  const ui      = SpreadsheetApp.getUi();
  const members = _getMemberNames();
  if (members.length === 0) { ui.alert("⚠️ メンバーが登録されていません。"); return; }
  const r1 = ui.prompt("グループ追加 ①", "グループ名を入力してください：", ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const groupName = r1.getResponseText().trim();
  if (!groupName) { ui.alert("グループ名が入力されていません。"); return; }
  const selected = _promptMemberSelection(ui, members, `「${groupName}」に追加するメンバー`);
  if (!selected) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let   sh = ss.getSheetByName(SELECT_SHEET_NAME);
  if (!sh) { _buildSelectSheet(); sh = ss.getSheetByName(SELECT_SHEET_NAME); }
  const nextGRow = _nextGroupRow(sh);
  sh.getRange(nextGRow, 5).insertCheckboxes().setValue(false);
  sh.getRange(nextGRow, 6).setValue(groupName).setFontSize(11).setFontWeight("bold");
  sh.getRange(nextGRow, 7).setValue(selected.join(",")).setFontSize(10).setFontColor("#5f6368");
  sh.getRange(nextGRow, 5, 1, 3).setBackground("#fff3e0");
  ui.alert(`✅ グループ「${groupName}」を登録しました！\nメンバー：${selected.join("、")}`);
  ss.setActiveSheet(sh);
}

// ─── グループ編集 ─────────────────────────────────────
function editGroup() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SELECT_SHEET_NAME);
  if (!sh) { ui.alert("メンバー選択シートが見つかりません。"); return; }
  const groups = _getGroups(sh);
  if (groups.length === 0) { ui.alert("登録されているグループがありません。"); return; }
  const groupList = groups.map((g, i) => `${i + 1}：${g.name}（${g.members.join("、")}）`).join("\n");
  const r1 = ui.prompt("グループ編集 ①", `編集するグループの番号を入力：\n\n${groupList}`, ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const gNum = parseInt(r1.getResponseText().trim());
  if (isNaN(gNum) || gNum < 1 || gNum > groups.length) { ui.alert(`❌ 1〜${groups.length} の番号を入力してください。`); return; }
  const target = groups[gNum - 1];
  const r2 = ui.prompt(`グループ編集 ②「${target.name}」`, `現在のメンバー：${target.members.join("、")}\n\n操作を選択：\n1：メンバーを追加\n2：メンバーを削除`, ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const opType     = r2.getResponseText().trim();
  const allMembers = _getMemberNames();
  if (opType === "1") {
    const notInGroup = allMembers.filter(m => !target.members.includes(m));
    if (notInGroup.length === 0) { ui.alert("追加できるメンバーがいません。"); return; }
    const selected = _promptMemberSelection(ui, notInGroup, `「${target.name}」に追加するメンバー`);
    if (!selected) return;
    const newMembers = [...target.members, ...selected];
    sh.getRange(target.shRow, 7).setValue(newMembers.join(",")).setFontSize(10).setFontColor("#5f6368");
    ui.alert(`✅ 追加しました！\n「${target.name}」のメンバー：${newMembers.join("、")}`);
  } else if (opType === "2") {
    if (target.members.length === 0) { ui.alert("グループにメンバーがいません。"); return; }
    const selected = _promptMemberSelection(ui, target.members, `「${target.name}」から削除するメンバー`);
    if (!selected) return;
    const newMembers = target.members.filter(m => !selected.includes(m));
    sh.getRange(target.shRow, 7).setValue(newMembers.join(",")).setFontSize(10).setFontColor("#5f6368");
    ui.alert(`✅ 削除しました！\n「${target.name}」のメンバー：${newMembers.length > 0 ? newMembers.join("、") : "（なし）"}`);
  } else {
    ui.alert("❌ 1 か 2 を入力してください。");
  }
}

// ─── グループ名変更 ───────────────────────────────────
function renameGroup() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SELECT_SHEET_NAME);
  if (!sh) { ui.alert("メンバー選択シートが見つかりません。"); return; }
  const groups = _getGroups(sh);
  if (groups.length === 0) { ui.alert("登録されているグループがありません。"); return; }
  const groupList = groups.map((g, i) => `${i + 1}：${g.name}`).join("\n");
  const r1 = ui.prompt("グループ名変更 ①", `名前を変更するグループの番号を入力：\n\n${groupList}`, ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const num = parseInt(r1.getResponseText().trim());
  if (isNaN(num) || num < 1 || num > groups.length) { ui.alert(`❌ 1〜${groups.length} の番号を入力してください。`); return; }
  const target = groups[num - 1];
  const r2 = ui.prompt("グループ名変更 ②", `「${target.name}」の新しいグループ名を入力：`, ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const newName = r2.getResponseText().trim();
  if (!newName) { ui.alert("グループ名が入力されていません。"); return; }
  if (newName === target.name) { ui.alert("同じ名前です。"); return; }
  sh.getRange(target.shRow, 6).setValue(newName).setFontSize(11).setFontWeight("bold");
  ui.alert(`✅ 「${target.name}」→「${newName}」に変更しました。`);
}

// ─── グループ削除 ─────────────────────────────────────
function deleteGroup() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SELECT_SHEET_NAME);
  if (!sh) { ui.alert("メンバー選択シートが見つかりません。"); return; }
  const groups = _getGroups(sh);
  if (groups.length === 0) { ui.alert("登録されているグループがありません。"); return; }
  const groupList = groups.map((g, i) => `${i + 1}：${g.name}`).join("\n");
  const r = ui.prompt("グループ削除", `削除するグループの番号を入力：\n\n${groupList}`, ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const num = parseInt(r.getResponseText().trim());
  if (isNaN(num) || num < 1 || num > groups.length) { ui.alert(`❌ 1〜${groups.length} の番号を入力してください。`); return; }
  const target = groups[num - 1];
  sh.getRange(target.shRow, 5, 1, 3).clearContent().clearFormat();
  ui.alert(`🗑 グループ「${target.name}」を削除しました。`);
}

// ─── メンバー選択ダイアログ（共通）──────────────────────
function _promptMemberSelection(ui, memberList, label) {
  const list = memberList.map((name, i) => `${i + 1}：${name}`).join("\n");
  const res  = ui.prompt(label, `番号をカンマ区切りで入力：\n\n${list}\n\n例）1,3`, ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return null;
  const nums = res.getResponseText().split(",").map(n => parseInt(n.trim())).filter(n => !isNaN(n));
  if (nums.length === 0) { ui.alert("番号が入力されていません。"); return null; }
  const invalid = nums.filter(n => n < 1 || n > memberList.length);
  if (invalid.length > 0) { ui.alert(`❌ 無効な番号：${invalid.join(", ")}`); return null; }
  return nums.map(n => memberList[n - 1]);
}

// ─── グループの次の空き行 ─────────────────────────────
function _nextGroupRow(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 2;
  const eVals = sh.getRange(2, 5, lastRow - 1, 1).getValues();
  for (let i = eVals.length - 1; i >= 0; i--) {
    if (eVals[i][0] !== "" && eVals[i][0] !== null) return i + 3;
  }
  return 2;
}

// ─── チェック済みメンバーを収集 ──────────────────────
function _collectSelectedMembers() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sh      = ss.getSheetByName(SELECT_SHEET_NAME);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const data     = sh.getRange(2, 1, lastRow - 1, 7).getValues();
  const selected = new Set();
  data.forEach(row => {
    if (row[1] === true && row[2]) selected.add(String(row[2]).trim());
    if (row[4] === true && row[6]) String(row[6]).split(",").map(n => n.trim()).filter(Boolean).forEach(n => selected.add(n));
  });
  return [...selected];
}

// ─── 開始打刻のコア処理（メンバー名の配列を受け取る）───────
// UIに依存しないため、会議自動化のトリガーなど自動実行からも呼び出せる。
function stampStartForMembers(names) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const results = [];
  names.forEach(name => {
    const memberSh = ss.getSheetByName(MEMBER_PREFIX + name);
    if (!memberSh) { results.push(`❌ ${name}：シートが見つかりません`); return; }
    const data    = memberSh.getDataRange().getValues();
    const hasOpen = data.slice(DATA_START - 1).some(r => r[COL_START - 1] && !r[COL_END - 1]);
    if (hasOpen) { results.push(`⚠️ ${name}：未終了の打刻あり（スキップ）`); return; }
    const nextRow = memberSh.getLastRow() + 1;
    memberSh.getRange(nextRow, COL_DATE).setValue(now).setNumberFormat("yyyy/MM/dd");
    memberSh.getRange(nextRow, COL_START).setValue(now).setNumberFormat("HH:mm:ss");
    memberSh.getRange(nextRow, COL_END).setValue("");
    memberSh.getRange(nextRow, COL_HOURS).setValue("").setNumberFormat("0.00");
    memberSh.getRange(nextRow, COL_MEMO).setValue("");
    memberSh.getRange(nextRow, COL_MANUAL).setValue("");
    memberSh.getRange(nextRow, 1, 1, 6).setBackground("#e6f4ea");
    results.push(`✅ ${name}`);
  });
  return results;
}

// ─── 終了打刻のコア処理（メンバー名の配列を受け取る）───────
function stampEndForMembers(names) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const results = [];
  names.forEach(name => {
    const memberSh = ss.getSheetByName(MEMBER_PREFIX + name);
    if (!memberSh) { results.push(`❌ ${name}：シートが見つかりません`); return; }
    const lastRow = memberSh.getLastRow();
    let targetRow = -1;
    for (let r = DATA_START; r <= lastRow; r++) {
      if (memberSh.getRange(r, COL_START).getValue() && !memberSh.getRange(r, COL_END).getValue()) { targetRow = r; break; }
    }
    if (targetRow === -1) { results.push(`⚠️ ${name}：開始打刻なし（スキップ）`); return; }
    const start = memberSh.getRange(targetRow, COL_START).getValue();
    const hours = (now - start) / 3600000;
    memberSh.getRange(targetRow, COL_END).setValue(now).setNumberFormat("HH:mm:ss");
    memberSh.getRange(targetRow, COL_HOURS).setValue(hours).setNumberFormat("0.00");
    memberSh.getRange(targetRow, 1, 1, 6).setBackground("#f8f9fa");
    results.push(`✅ ${name}：${hours.toFixed(2)}h`);
  });
  refreshSummary();
  return results;
}

// ─── 実績時間での打刻（開始・終了が両方確定しているケース）─────
// Meetの参加ログ(入室〜退室時刻)など、開始・終了が同時に分かっている場合に
// 1行にまとめて記録する。stampStartForMembers/stampEndForMembersと違い、
// 「今」ではなく指定した開始・終了時刻をそのまま書き込む。
function stampExactForMembers(entries) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const results = [];
  entries.forEach(({ name, start, end }) => {
    const memberSh = ss.getSheetByName(MEMBER_PREFIX + name);
    if (!memberSh) { results.push(`❌ ${name}：シートが見つかりません`); return; }
    const hours = (end - start) / 3600000;
    const nextRow = memberSh.getLastRow() + 1;
    memberSh.getRange(nextRow, COL_DATE).setValue(start).setNumberFormat("yyyy/MM/dd");
    memberSh.getRange(nextRow, COL_START).setValue(start).setNumberFormat("HH:mm:ss");
    memberSh.getRange(nextRow, COL_END).setValue(end).setNumberFormat("HH:mm:ss");
    memberSh.getRange(nextRow, COL_HOURS).setValue(hours).setNumberFormat("0.00");
    memberSh.getRange(nextRow, COL_MEMO).setValue("Meet参加ログ");
    memberSh.getRange(nextRow, COL_MANUAL).setValue("");
    memberSh.getRange(nextRow, 1, 1, 6).setBackground("#f8f9fa");
    results.push(`✅ ${name}：${hours.toFixed(2)}h`);
  });
  refreshSummary();
  return results;
}

// ─── 一括開始打刻（メニュー用・チェックボックス選択を使用）───
function stampStartSelected() {
  const ui       = SpreadsheetApp.getUi();
  const selected = _collectSelectedMembers();
  if (selected.length === 0) { ui.alert("⚠️ メンバーが選択されていません。"); return; }
  const results = stampStartForMembers(selected);
  _resetAllCheckboxes();
  ui.alert(`▶ 開始打刻 完了\n\n${results.join("\n")}`);
}

// ─── 一括終了打刻（メニュー用・チェックボックス選択を使用）───
function stampEndSelected() {
  const ui       = SpreadsheetApp.getUi();
  const selected = _collectSelectedMembers();
  if (selected.length === 0) { ui.alert("⚠️ メンバーが選択されていません。"); return; }
  const results = stampEndForMembers(selected);
  _resetAllCheckboxes();
  ui.alert(`⏹ 終了打刻 完了\n\n${results.join("\n")}`);
}

// ─── 打刻手動修正（赤伝票・黒伝票方式）─────────────────
function manualCorrect() {
  const ui      = SpreadsheetApp.getUi();
  const members = _getMemberNames();
  if (members.length === 0) { ui.alert("メンバーが登録されていません。"); return; }

  const memberList = members.map((name, i) => `${i + 1}：${name}`).join("\n");
  const r1 = ui.prompt("打刻手動修正 ①", `修正するメンバーの番号を入力：\n\n${memberList}`, ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const memberNum = parseInt(r1.getResponseText().trim());
  if (isNaN(memberNum) || memberNum < 1 || memberNum > members.length) {
    ui.alert(`❌ 1〜${members.length} の番号を入力してください。`); return;
  }
  const name     = members[memberNum - 1];
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const memberSh = ss.getSheetByName(MEMBER_PREFIX + name);
  if (!memberSh) { ui.alert(`❌ 「${name}」のシートが見つかりません。`); return; }

  const r2 = ui.prompt(
    "打刻手動修正 ②",
    "修正の種類を番号で入力：\n" +
    "1：開始打刻を追加（開始忘れ）\n" +
    "2：終了打刻を修正（終了忘れ・赤伝黒伝）\n" +
    "3：既存の行を修正（赤伝・黒伝方式）",
    ui.ButtonSet.OK_CANCEL
  );
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const type = r2.getResponseText().trim();

  if (type === "1") {
    // ── 開始忘れ：新行追加（修正前なし） ──
    const r3 = ui.prompt("開始時刻を入力", "例）2026/06/18 09:30", ui.ButtonSet.OK_CANCEL);
    if (r3.getSelectedButton() !== ui.Button.OK) return;
    const startDate = new Date(r3.getResponseText().trim());
    if (isNaN(startDate)) { ui.alert("❌ 日時の形式が正しくありません。"); return; }

    const nextRow = memberSh.getLastRow() + 1;
    memberSh.getRange(nextRow, COL_DATE).setValue(startDate).setNumberFormat("yyyy/MM/dd");
    memberSh.getRange(nextRow, COL_START).setValue(startDate).setNumberFormat("HH:mm:ss");
    memberSh.getRange(nextRow, COL_MANUAL).setValue("手動追加");
    memberSh.getRange(nextRow, 1, 1, 6).setBackground("#fff8e1");
    ui.alert(`✅ ${name} の開始打刻を追加しました（手動）`);

  } else if (type === "2") {
    // ── 終了忘れ：赤伝（元行）＋黒伝（修正後新行）──
    const lastRow = memberSh.getLastRow();
    let targetRow = -1;
    for (let r = DATA_START; r <= lastRow; r++) {
      if (memberSh.getRange(r, COL_START).getValue() && !memberSh.getRange(r, COL_END).getValue()) {
        targetRow = r; break;
      }
    }
    if (targetRow === -1) { ui.alert("⚠️ 未終了の打刻が見つかりません。"); return; }

    const r3 = ui.prompt("終了時刻を入力", "例）2026/06/18 18:00", ui.ButtonSet.OK_CANCEL);
    if (r3.getSelectedButton() !== ui.Button.OK) return;
    const endDate = new Date(r3.getResponseText().trim());
    if (isNaN(endDate)) { ui.alert("❌ 日時の形式が正しくありません。"); return; }

    // 元の値を取得
    const oldDate  = memberSh.getRange(targetRow, COL_DATE).getValue();
    const oldStart = memberSh.getRange(targetRow, COL_START).getValue();
    const oldMemo  = memberSh.getRange(targetRow, COL_MEMO).getValue();
    const hours    = (endDate - oldStart) / 3600000;

    // 赤伝：元の未終了行を赤背景＋打消し線
    memberSh.getRange(targetRow, COL_MANUAL).setValue("修正前");
    memberSh.getRange(targetRow, 1, 1, 6)
      .setBackground("#fce8e6").setFontLine("line-through").setFontColor("#999999");

    // 黒伝：修正後を末尾に追記
    const nextRow = memberSh.getLastRow() + 1;
    memberSh.getRange(nextRow, COL_DATE).setValue(oldDate).setNumberFormat("yyyy/MM/dd");
    memberSh.getRange(nextRow, COL_START).setValue(oldStart).setNumberFormat("HH:mm:ss");
    memberSh.getRange(nextRow, COL_END).setValue(endDate).setNumberFormat("HH:mm:ss");
    memberSh.getRange(nextRow, COL_HOURS).setValue(hours).setNumberFormat("0.00");
    memberSh.getRange(nextRow, COL_MEMO).setValue(oldMemo);
    memberSh.getRange(nextRow, COL_MANUAL).setValue("修正後");
    memberSh.getRange(nextRow, 1, 1, 6).setBackground("#e8f5e9").setFontLine("none").setFontColor("#000000");

    refreshSummary();
    ui.alert(
      `✅ ${name} の終了打刻を修正しました。\n\n` +
      `【修正前】赤行（${targetRow}行目）で確認できます\n` +
      `【修正後】緑行（${nextRow}行目）に追記\n` +
      `活動時間：${hours.toFixed(2)}h`
    );

  } else if (type === "3") {
    // ── 既存行修正：赤伝（元行）＋黒伝（修正後新行）──
    const r3 = ui.prompt("修正する行番号を入力", `${name}のシートを確認して行番号（数字）を入力：`, ui.ButtonSet.OK_CANCEL);
    if (r3.getSelectedButton() !== ui.Button.OK) return;
    const rowNum = parseInt(r3.getResponseText().trim());
    if (isNaN(rowNum) || rowNum < DATA_START) { ui.alert("❌ 正しい行番号を入力してください。"); return; }

    const oldDate  = memberSh.getRange(rowNum, COL_DATE).getValue();
    const oldStart = memberSh.getRange(rowNum, COL_START).getValue();
    const oldEnd   = memberSh.getRange(rowNum, COL_END).getValue();
    const oldHours = memberSh.getRange(rowNum, COL_HOURS).getValue();
    const oldMemo  = memberSh.getRange(rowNum, COL_MEMO).getValue();

    const r4 = ui.prompt("新しい開始時刻", "例）2026/06/18 09:00（変更しない場合は空白）", ui.ButtonSet.OK_CANCEL);
    if (r4.getSelectedButton() !== ui.Button.OK) return;
    const r5 = ui.prompt("新しい終了時刻", "例）2026/06/18 17:00（変更しない場合は空白）", ui.ButtonSet.OK_CANCEL);
    if (r5.getSelectedButton() !== ui.Button.OK) return;

    const newStartStr = r4.getResponseText().trim();
    const newEndStr   = r5.getResponseText().trim();
    const newStart    = newStartStr ? new Date(newStartStr) : oldStart;
    const newEnd      = newEndStr   ? new Date(newEndStr)   : oldEnd;

    if (newStartStr && isNaN(newStart)) { ui.alert("❌ 開始時刻の形式が正しくありません。"); return; }
    if (newEndStr   && isNaN(newEnd))   { ui.alert("❌ 終了時刻の形式が正しくありません。"); return; }

    const newHours = (newStart && newEnd) ? (newEnd - newStart) / 3600000 : oldHours;

    // 赤伝：元行を赤背景＋打消し線
    memberSh.getRange(rowNum, COL_MANUAL).setValue("修正前");
    memberSh.getRange(rowNum, 1, 1, 6)
      .setBackground("#fce8e6").setFontLine("line-through").setFontColor("#999999");

    // 黒伝：修正後を末尾に追記
    const nextRow = memberSh.getLastRow() + 1;
    memberSh.getRange(nextRow, COL_DATE).setValue(oldDate).setNumberFormat("yyyy/MM/dd");
    memberSh.getRange(nextRow, COL_START).setValue(newStart).setNumberFormat("HH:mm:ss");
    memberSh.getRange(nextRow, COL_END).setValue(newEnd).setNumberFormat("HH:mm:ss");
    memberSh.getRange(nextRow, COL_HOURS).setValue(newHours).setNumberFormat("0.00");
    memberSh.getRange(nextRow, COL_MEMO).setValue(oldMemo);
    memberSh.getRange(nextRow, COL_MANUAL).setValue("修正後");
    memberSh.getRange(nextRow, 1, 1, 6).setBackground("#e8f5e9").setFontLine("none").setFontColor("#000000");

    refreshSummary();
    ui.alert(
      `✅ ${name} の ${rowNum} 行目を修正しました。\n\n` +
      `【修正前】赤行（${rowNum}行目）で確認できます\n` +
      `【修正後】緑行（${nextRow}行目）に追記\n` +
      `活動時間：${newHours.toFixed(2)}h`
    );

  } else {
    ui.alert("❌ 1〜3の番号を入力してください。");
  }
}
// ─── シート保護 ───────────────────────────────────────
function protectMemberSheets() {
  const ui = SpreadsheetApp.getUi();
  if (!_authenticate("🔒 シート保護")) return;

  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const members = _getMemberNames();
  let count = 0;

  members.forEach(name => {
    const sh = ss.getSheetByName(MEMBER_PREFIX + name);
    if (!sh) return;

    // 既存の保護を一旦削除
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());

    // 警告モードで保護（編集しようとすると警告が出るが強制はしない）
    sh.protect()
      .setDescription("打刻データ保護 ※誤編集防止のため警告が出ます")
      .setWarningOnly(true); // ← 警告モード
    count++;
  });

  ui.alert(
    `✅ ${count}枚のメンバーシートに保護をかけました。\n\n` +
    "【警告モード】\n" +
    "セルを編集しようとすると警告ダイアログが表示されます。\n" +
    "「OK」を押せば編集できますが、誤編集の抑止になります。\n\n" +
    "※Googleスプレッドシートの仕様上、編集権限を持つユーザーの\n" +
    "完全なブロックはできません。"
  );
}

// ─── シート保護解除 ───────────────────────────────────
function unprotectMemberSheets() {
  const ui = SpreadsheetApp.getUi();
  if (!_authenticate("🔓 シート保護解除")) return;
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const members = _getMemberNames();
  let count = 0;
  members.forEach(name => {
    const sh = ss.getSheetByName(MEMBER_PREFIX + name);
    if (!sh) return;
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => { p.remove(); count++; });
  });
  ui.alert(`✅ ${count}枚のメンバーシートの保護を解除しました。`);
}

// ─── 年度リセット ─────────────────────────────────────
function yearlyReset() {
  const ui = SpreadsheetApp.getUi();
  if (!_authenticate("🔄 年度リセット")) return;
  const now        = new Date();
  const fiscalYear = now.getMonth() >= 3 ? now.getFullYear() : now.getFullYear() - 1;
  const confirm    = ui.alert("年度リセット確認", `${fiscalYear}年度のデータをバックアップしてリセットします。\nこの操作は元に戻せません。実行しますか？`, ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const members = _getMemberNames();
  const backupName = `${BACKUP_PREFIX}${fiscalYear}年度`;
  let backupSh = ss.getSheetByName(backupName);
  if (backupSh) ss.deleteSheet(backupSh);
  backupSh = ss.insertSheet(backupName);
  backupSh.getRange(1, 1, 1, 7).setValues([["メンバー", "日付", "開始時刻", "終了時刻", "活動時間(h)", "メモ", "修正"]]);
  backupSh.getRange(1, 1, 1, 7).setBackground("#424242").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
  let backupRow = 2;
  let totalBackupRows = 0;
  members.forEach(name => {
    const memberSh = ss.getSheetByName(MEMBER_PREFIX + name);
    if (!memberSh) return;
    const lastRow = memberSh.getLastRow();
    if (lastRow >= DATA_START) {
      const data = memberSh.getRange(DATA_START, 1, lastRow - DATA_START + 1, 6).getValues();
      data.forEach(row => {
        backupSh.getRange(backupRow, 1).setValue(name);
        backupSh.getRange(backupRow, 2, 1, 6).setValues([row]);
        backupRow++; totalBackupRows++;
      });
      memberSh.getRange(DATA_START, 1, lastRow - DATA_START + 1, 6).clearContent().clearFormat();
    }
  });
  backupSh.setColumnWidths(1, 7, 120);
  backupSh.setFrozenRows(1);
  refreshSummary();
  ui.alert(`✅ 年度リセット完了！\nバックアップ：${totalBackupRows}件 → 「${backupName}」シート`);
}

// ─── チェックをリセット ───────────────────────────────
function _resetAllCheckboxes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SELECT_SHEET_NAME);
  if (!sh) return;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  sh.getRange(2, 2, lastRow - 1, 1).setValue(false);
  sh.getRange(2, 5, lastRow - 1, 1).setValue(false);
}

// ─── メンバー名一覧取得 ───────────────────────────────
function _getMemberNames() {
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .filter(s => s.getName().startsWith(MEMBER_PREFIX))
    .map(s => s.getName().replace(MEMBER_PREFIX, ""));
}

// ─── サマリー更新（ランキング形式）──────────────────────
function refreshSummary() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const summary = ss.getSheetByName(SUMMARY_SHEET_NAME);
  if (!summary) return;
  const sheets  = ss.getSheets().filter(s => s.getName().startsWith(MEMBER_PREFIX));
  let totalAll  = 0;
  const ranking = [];
  sheets.forEach(sh => {
    const memberName = sh.getName().replace(MEMBER_PREFIX, "");
    const lastRow    = sh.getLastRow();
    let total        = 0;
    if (lastRow >= DATA_START) {
      // 「修正前」タグの行は集計から除外
      const rows = sh.getRange(DATA_START, 1, lastRow - DATA_START + 1, 6).getValues();
      rows.forEach(row => {
        if (row[COL_MANUAL - 1] !== "修正前" && typeof row[COL_HOURS - 1] === "number") {
          total += row[COL_HOURS - 1];
        }
      });
    }
    totalAll += total;
    ranking.push({ name: memberName, total });
  });
  ranking.sort((a, b) => b.total - a.total);
  const lastRow = summary.getLastRow();
  if (lastRow >= 4) summary.getRange(4, 1, lastRow - 3, 2).clearContent().clearFormat();
  ranking.forEach((member, i) => {
    const row   = 4 + i;
    const medal = i === 0 ? "🥇" : i === 1 ? "🥈" : i === 2 ? "🥉" : `${i + 1}位`;
    summary.getRange(row, 1).setValue(`${medal}  ${member.name}`);
    summary.getRange(row, 2).setValue(member.total).setNumberFormat("0.00").setHorizontalAlignment("center");
    _formatSummaryRow(summary, row);
  });
  summary.getRange(2, 2).setValue(totalAll).setNumberFormat("0.00");
  _updateDailyContent(summary);
}

// ─── メンバーシート作成 ───────────────────────────────
function _ensureMemberSheet(name) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const shName = MEMBER_PREFIX + name;
  let   sh     = ss.getSheetByName(shName);
  if (!sh) {
    sh = ss.insertSheet(shName);
    sh.getRange(1, 1, 1, 6).setValues([["日付", "開始時刻", "終了時刻", "活動時間(h)", "メモ", "修正"]]);
    sh.getRange(1, 1, 1, 6).setBackground("#1a73e8").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
    sh.setColumnWidth(1, 110); sh.setColumnWidth(2, 100);
    sh.setColumnWidth(3, 100); sh.setColumnWidth(4, 120);
    sh.setColumnWidth(5, 180); sh.setColumnWidth(6, 80);
    sh.setFrozenRows(1);
  }
  return sh;
}

// ─── サマリーシート作成 ───────────────────────────────
function _ensureSummarySheet(names) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let   sh = ss.getSheetByName(SUMMARY_SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SUMMARY_SHEET_NAME, 0); } else { sh.clearContents(); }
  sh.getRange(1, 1, 1, 2).merge().setValue("📊 活動時間サマリー").setFontSize(14).setFontWeight("bold")
    .setBackground("#1a73e8").setFontColor("#ffffff").setHorizontalAlignment("center");
  sh.getRange(2, 1).setValue("🏆 全体合計").setFontWeight("bold");
  sh.getRange(2, 2).setValue(0).setNumberFormat("0.00").setFontWeight("bold");
  sh.getRange(2, 1, 1, 2).setBackground("#fce8e6");
  sh.getRange(3, 1).setValue("メンバー").setFontWeight("bold").setBackground("#f1f3f4");
  sh.getRange(3, 2).setValue("合計時間(h)").setFontWeight("bold").setBackground("#f1f3f4").setHorizontalAlignment("center");
  sh.setFrozenRows(3);
  names.forEach((name, i) => {
    const row = 4 + i;
    sh.getRange(row, 1).setValue(name);
    sh.getRange(row, 2).setValue(0).setNumberFormat("0.00").setHorizontalAlignment("center");
    _formatSummaryRow(sh, row);
  });
  sh.setColumnWidth(1, 150); sh.setColumnWidth(2, 130); sh.setColumnWidth(3, 20);
  _updateDailyContent(sh);
}

// ─── サマリー行フォーマット ───────────────────────────
function _formatSummaryRow(sh, row) {
  sh.getRange(row, 1, 1, 2)
    .setBackground(row % 2 === 0 ? "#ffffff" : "#f8f9fa")
    .setBorder(true, true, true, true, null, null, "#dadce0", SpreadsheetApp.BorderStyle.SOLID);
}

// ─── スマホ打刻シート作成 ─────────────────────────────
function setupCheckboxSheet() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const shName  = "📱 スマホ打刻";
  let   sh      = ss.getSheetByName(shName);
  const members = _getMemberNames();

  if (members.length === 0) {
    SpreadsheetApp.getUi().alert("⚠️ メンバーが登録されていません。先に初期セットアップを実行してください。");
    return;
  }

  if (!sh) { sh = ss.insertSheet(shName); } else { sh.clearContents(); sh.clearFormats(); }

  // ── タイトル・説明 ──
  sh.getRange("A1:D1").merge();
  sh.getRange("A1").setValue("📱 スマホ打刻シート")
    .setFontSize(14).setFontWeight("bold")
    .setBackground("#1a73e8").setFontColor("#ffffff").setHorizontalAlignment("center");
  sh.getRange("A2:D2").merge();
  sh.getRange("A2").setValue("チェックを入れると自動で打刻されます（チェックは自動でOFFに戻ります）")
    .setFontSize(10).setFontColor("#5f6368").setBackground("#f8f9fa");

  // ── 列ヘッダー ──
  sh.getRange("A3").setValue("名前 / グループ名").setFontWeight("bold").setBackground("#e8f0fe").setHorizontalAlignment("center");
  sh.getRange("B3").setValue("▶ 開始").setFontWeight("bold").setBackground("#34a853").setFontColor("#ffffff").setHorizontalAlignment("center");
  sh.getRange("C3").setValue("⏹ 終了").setFontWeight("bold").setBackground("#ea4335").setFontColor("#ffffff").setHorizontalAlignment("center");
  sh.getRange("D3").setValue("状態").setFontWeight("bold").setBackground("#e8f0fe").setHorizontalAlignment("center");

  let currentRow = 4;

  // ── 個人セクション ──
  sh.getRange(currentRow, 1, 1, 4).merge();
  sh.getRange(currentRow, 1).setValue("👤 個人")
    .setFontWeight("bold").setBackground("#e8f0fe").setFontSize(11);
  currentRow++;

  members.forEach(name => {
    sh.getRange(currentRow, 1).setValue(name).setFontSize(11);
    sh.getRange(currentRow, 2).insertCheckboxes().setValue(false);
    sh.getRange(currentRow, 3).insertCheckboxes().setValue(false);
    sh.getRange(currentRow, 4).setValue("待機中").setFontColor("#5f6368").setHorizontalAlignment("center");
    sh.getRange(currentRow, 1, 1, 4).setBackground(currentRow % 2 === 0 ? "#ffffff" : "#f8f9fa");
    currentRow++;
  });

  // ── グループセクション ──
  const selectSh = ss.getSheetByName(SELECT_SHEET_NAME);
  const groups   = selectSh ? _getGroups(selectSh) : [];

  if (groups.length > 0) {
    // 区切り行
    sh.getRange(currentRow, 1, 1, 4).merge();
    sh.getRange(currentRow, 1).setValue("👥 グループ")
      .setFontWeight("bold").setBackground("#fce8e6").setFontSize(11);
    currentRow++;

    groups.forEach(g => {
      sh.getRange(currentRow, 1).setValue(g.name).setFontSize(11).setFontWeight("bold");
      sh.getRange(currentRow, 2).insertCheckboxes().setValue(false);
      sh.getRange(currentRow, 3).insertCheckboxes().setValue(false);
      sh.getRange(currentRow, 4).setValue(`(${g.members.join("・")})`).setFontColor("#5f6368").setFontSize(9).setHorizontalAlignment("left");
      sh.getRange(currentRow, 1, 1, 4).setBackground("#fff3e0");
      currentRow++;
    });
  }

  // 列幅・行高
  sh.setColumnWidth(1, 160);
  sh.setColumnWidth(2, 65);
  sh.setColumnWidth(3, 65);
  sh.setColumnWidth(4, 160);
  sh.setFrozenRows(3);

  // セクション行の種別をメモに記録（stampFromCheckboxが参照）
  // A列が「👤 個人」または「👥 グループ」のセクション行はスキップする

  ss.setActiveSheet(sh);
  SpreadsheetApp.getUi().alert(
    "✅ スマホ打刻シートを作成しました！\n\n" +
    "⚠️ onEditトリガーが未設定の場合は設定してください：\n" +
    "GASエディタ →「⏰ トリガー」→「＋ トリガーを追加」\n" +
    "・関数：stampFromCheckbox\n" +
    "・イベント：スプレッドシートから → 編集時"
  );
}

// ─── チェックボックス打刻（onEditトリガー）────────────
function stampFromCheckbox(e) {
  if (!e) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== "📱 スマホ打刻") return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < 4) return;
  if (col !== 2 && col !== 3) return;
  if (e.value !== "TRUE") return;

  const cellVal = sh.getRange(row, 1).getValue();
  if (!cellVal) { e.range.setValue(false); return; }

  // セクションヘッダー行（「👤 個人」「👥 グループ」）はスキップ
  if (String(cellVal).includes("👤") || String(cellVal).includes("👥")) {
    e.range.setValue(false); return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const selectSh = ss.getSheetByName(SELECT_SHEET_NAME);
  const groups   = selectSh ? _getGroups(selectSh) : [];

  // グループかどうかを判定
  const matchedGroup = groups.find(g => g.name === String(cellVal).trim());
  const targets      = matchedGroup ? matchedGroup.members : [String(cellVal).trim()];

  const now     = new Date();
  const results = [];

  targets.forEach(name => {
    const memberSh = ss.getSheetByName(MEMBER_PREFIX + name);
    if (!memberSh) { results.push(`❌ ${name}：シートなし`); return; }

    if (col === 2) {
      // ── 開始打刻 ──
      const data    = memberSh.getDataRange().getValues();
      const hasOpen = data.slice(DATA_START - 1).some(r => r[COL_START - 1] && !r[COL_END - 1]);
      if (hasOpen) { results.push(`⚠️ ${name}：打刻中`); return; }

      const nextRow = memberSh.getLastRow() + 1;
      memberSh.getRange(nextRow, COL_DATE).setValue(now).setNumberFormat("yyyy/MM/dd");
      memberSh.getRange(nextRow, COL_START).setValue(now).setNumberFormat("HH:mm:ss");
      memberSh.getRange(nextRow, COL_END).setValue("");
      memberSh.getRange(nextRow, COL_HOURS).setValue("").setNumberFormat("0.00");
      memberSh.getRange(nextRow, COL_MEMO).setValue("");
      memberSh.getRange(nextRow, COL_MANUAL).setValue("");
      memberSh.getRange(nextRow, 1, 1, 6).setBackground("#e6f4ea");
      results.push(`✅ ${name}`);

    } else if (col === 3) {
      // ── 終了打刻 ──
      const lastRow = memberSh.getLastRow();
      let targetRow = -1;
      for (let r = DATA_START; r <= lastRow; r++) {
        if (memberSh.getRange(r, COL_START).getValue() && !memberSh.getRange(r, COL_END).getValue()) {
          targetRow = r; break;
        }
      }
      if (targetRow === -1) { results.push(`⚠️ ${name}：未開始`); return; }

      const start = memberSh.getRange(targetRow, COL_START).getValue();
      const hours = (now - start) / 3600000;
      memberSh.getRange(targetRow, COL_END).setValue(now).setNumberFormat("HH:mm:ss");
      memberSh.getRange(targetRow, COL_HOURS).setValue(hours).setNumberFormat("0.00");
      memberSh.getRange(targetRow, 1, 1, 6).setBackground("#f8f9fa");
      results.push(`✅ ${name}：${hours.toFixed(1)}h`);
    }
  });

  // 状態列を更新
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm");
  if (col === 2) {
    sh.getRange(row, 4).setValue(`▶ ${timeStr}〜`).setFontColor("#34a853").setFontWeight("bold");
    sh.getRange(row, 1, 1, 4).setBackground(matchedGroup ? "#e6f9f0" : "#e6f4ea");
  } else {
    const okCount = results.filter(r => r.startsWith("✅")).length;
    sh.getRange(row, 4).setValue(`⏹ ${okCount}人完了`).setFontColor("#5f6368").setFontWeight("normal");
    sh.getRange(row, 1, 1, 4).setBackground(matchedGroup ? "#fff3e0" : (row % 2 === 0 ? "#ffffff" : "#f8f9fa"));
    refreshSummary();
  }

  e.range.setValue(false);
}



// ─── マニュアルシート作成 ────────────────────────────
function createManualSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let   sh = ss.getSheetByName("📖 使い方");
  if (!sh) { sh = ss.insertSheet("📖 使い方", 0); } else { sh.clearContents(); sh.clearFormats(); }
  const sections = [
    [1,  "📖 活動タイマー 使い方マニュアル", "h1"],
    [2,  "このスプレッドシートは社内活動の作業時間を記録・集計するツールです。", "body"],
    [3,  "", "blank"],
    [4,  "▶ 動画マニュアル", "h2"],
    [5,  "📹 基本操作編（YouTubeリンクをここに貼る）", "link"],
    [6,  "📹 グループ設定編（YouTubeリンクをここに貼る）", "link"],
    [7,  "", "blank"],
    [8,  "① 初回セットアップ", "h2"],
    [9,  "メニュー「⏱ 活動タイマー」→「⚙ 初期セットアップ」を開き、メンバー名をカンマ区切りで入力します。\n例）田中, 鈴木, 佐藤", "body"],
    [10, "", "blank"],
    [11, "② 打刻の基本操作（PC）", "h2"],
    [12, "1. メニュー「📋 メンバー選択シートを開く」をクリック\n2. 参加するメンバーまたはグループにチェックを入れる\n3. メニュー「▶ 開始打刻を実行」をクリック\n4. 活動終了後、同じようにチェックを入れて「⏹ 終了打刻を実行」をクリック", "body"],
    [13, "", "blank"],
    [14, "③ スマホからの打刻", "h2"],
    [15, "「📱 スマホ打刻」シートを開き、自分の名前の横にある「▶ 開始」チェックボックスをONにするだけで打刻できます。終了時は「⏹ 終了」をONにします。チェックは自動でOFFに戻ります。", "body"],
    [16, "", "blank"],
    [17, "④ グループの追加・編集", "h2"],
    [18, "メニュー「👥 グループを追加」からグループ名とメンバーを番号で選択して登録できます。\n「✏️ グループを編集」からメンバーの追加・削除、「🔤 グループ名を変更」で名称変更が可能です。", "body"],
    [19, "", "blank"],
    [20, "⑤ 打刻忘れの修正", "h2"],
    [21, "メニュー「🔧 打刻を手動修正」から修正できます。\n・開始忘れ：さかのぼって開始時刻を登録\n・終了忘れ：終了時刻を入力すると時間が自動計算\n・既存行の修正：修正前（赤行）と修正後（緑行）が両方残ります", "body"],
    [22, "", "blank"],
    [23, "⑥ 管理者向け操作", "h2"],
    [24, "以下の操作にはパスワードが必要です：\n・メンバーシートの保護 / 解除（誤編集防止）\n・メンバー名の変更・削除\n・年度リセット（過去データはバックアップシートに保存）", "body"],
    [25, "", "blank"],
    [26, "⑦ サマリーの見方", "h2"],
    [27, "「サマリー」シートに全メンバーの合計時間がランキング形式で表示されます。\n🥇🥈🥉で上位3名がひと目でわかります。打刻のたびに自動更新されます。", "body"],
    [28, "", "blank"],
    [29, "🎨 色の意味", "h2"],
    [30, "🟩 緑：進行中（開始打刻済み・未終了）\n⬜ グレー：完了した打刻\n🟨 黄：手動で追加した打刻\n🟥 赤（打消し線）：修正前の打刻\n🟢 緑（薄）：修正後の打刻", "body"],
  ];
  sections.forEach(([row, text, type]) => {
    const cell = sh.getRange(row, 1);
    cell.setValue(text);
    if (type === "h1") {
      sh.getRange(row, 1, 1, 3).merge();
      cell.setFontSize(16).setFontWeight("bold").setBackground("#1a73e8").setFontColor("#ffffff").setHorizontalAlignment("center");
    } else if (type === "h2") {
      sh.getRange(row, 1, 1, 3).merge();
      cell.setFontSize(12).setFontWeight("bold").setBackground("#e8f0fe").setFontColor("#1a73e8");
    } else if (type === "body") {
      sh.getRange(row, 1, 1, 3).merge();
      cell.setFontSize(11).setWrap(true);
    } else if (type === "link") {
      sh.getRange(row, 1, 1, 3).merge();
      cell.setFontSize(11).setFontColor("#1a73e8");
    }
  });
  sh.setColumnWidth(1, 600);
  sh.setRowHeights(1, 33, 28);
  sh.setFrozenRows(1);
  SpreadsheetApp.getUi().alert("✅ マニュアルシートを作成しました！\n「📖 使い方」シートを確認してください。");
}

// ─── 今日のコンテンツ表示 ────────────────────────────
function _updateDailyContent(sh) {
  const today   = new Date();
  const month   = today.getMonth() + 1;
  const day     = today.getDate();
  const dateStr = `${month}月${day}日`;

  const whatsDays = [
    ["1/1","元日。グレゴリオ暦の年始。日本では初詣や初日の出を楽しむ日。"],["1/2","初夢の日。その年最初に見る夢「初夢」を楽しむ日。"],["1/3","瞳の日。「ひ（1）とみ（3）」の語呂合わせから制定。"],["1/4","石の日。「い（1）し（4）」の語呂合わせから制定。"],["1/5","いちごの日。「いち（1）ご（5）」の語呂合わせから制定。"],["1/6","色の日。「い（1）ろ（6）」の語呂合わせから制定。"],["1/7","七草の日。春の七草を入れたお粥を食べ、一年の無病息災を願う日。"],["1/8","勝負の日。「一か八か」にちなんで制定。"],["1/9","とんちの日。一休さんの「一（いち）休（く）」の語呂合わせから制定。"],["1/10","明太子の日。1949年のこの日、ふくやが初めて辛子明太子を販売した。"],["1/11","鏡開きの日。お正月に供えた鏡餅を割って食べ、無病息災を願う日。"],["1/12","スキーの日。1911年のこの日、日本初のスキー指導が行われた。"],["1/13","ピースの日。「ピー（P）ス（S）」の語呂合わせから制定。"],["1/14","尖閣諸島の日。1895年のこの日、日本が尖閣諸島を沖縄県に編入。"],["1/15","小正月。松飾りを外してどんど焼きを行う日。"],["1/16","禁酒の日。1920年のこの日、アメリカで禁酒法が発効された。"],["1/17","防災とボランティアの日。1995年の阪神淡路大震災が起きた日。"],["1/18","都バスの日。1924年のこの日、東京市営バスが運行を開始した。"],["1/19","家庭消火器点検の日。「良い（1）消火（19）」の語呂合わせから制定。"],["1/20","大寒。一年で最も寒い時期とされる二十四節気の一つ。"],["1/21","ライバルが手を結ぶ日。1866年のこの日、薩長同盟が結ばれた。"],["1/22","カレーの日。1982年のこの日、全国の小中学校の給食でカレーが出された。"],["1/23","電子メールの日。「1（いい）23（ふみ）」の語呂合わせから制定。"],["1/24","ゴールドラッシュの日。1848年のこの日、カリフォルニアで金が発見された。"],["1/25","中華まんの日。寒い日に中華まんを食べる風習から制定。"],["1/26","文化財防火デー。1949年のこの日、法隆寺金堂の壁画が焼損した。"],["1/27","国旗制定記念日。1870年のこの日、太政官布告で日の丸が国旗と定められた。"],["1/28","コードの日。「コー（1）ド（28）」の語呂合わせから制定。"],["1/29","タウン情報の日。1973年のこの日、日本初のタウン情報誌が創刊された。"],["1/30","3分間電話の日。1970年のこの日、公衆電話の通話料が3分10円になった。"],["1/31","愛妻家の日。「1（あい）31（さい）」の語呂合わせから制定。"],
    ["2/1","テレビ放送記念日。1953年のこの日、NHKが日本初のテレビ放送を開始した。"],["2/2","世界湿地の日。1971年のこの日、ラムサール条約が締結された。"],["2/3","節分。豆まきで鬼を追い払い、福を呼び込む日。"],["2/4","立春。暦の上では春が始まる日。二十四節気の一つ。"],["2/5","プロ野球の日。1936年のこの日、日本初のプロ野球試合が行われた。"],["2/6","海苔の日。702年のこの日、大宝律令で海苔が年貢として認められた。"],["2/7","北方領土の日。1855年のこの日、日露和親条約が締結された。"],["2/8","ニッポン放送開局記念日。1954年のこの日、ニッポン放送が開局した。"],["2/9","漫画の日。語呂合わせから制定。"],["2/10","ニットの日。「ニッ（2）ト（10）」の語呂合わせから制定。"],["2/11","建国記念の日。日本の建国を祝う国民の祝日。"],["2/12","ダーウィンの日。1809年のこの日、チャールズ・ダーウィンが誕生した。"],["2/13","苗字制定記念日。1875年のこの日、平民も苗字を名乗ることが義務付けられた。"],["2/14","バレンタインデー。愛を伝えるためにチョコレートを贈る日。"],["2/15","涅槃会。お釈迦様が亡くなった日とされ、各地の寺院で法要が行われる。"],["2/16","天気図記念日。1883年のこの日、日本初の天気図が作成された。"],["2/17","天使のささやきの日。1978年のこの日、北海道幌加内町で気温-41.2℃を記録した。"],["2/18","エアメールの日。1911年のこの日、世界初の公式航空郵便が飛んだ。"],["2/19","万国郵便連合加盟記念日。1877年のこの日、日本が万国郵便連合に加盟した。"],["2/20","旅券の日。1878年のこの日、外務省布達によりパスポートの規則が定められた。"],["2/21","国際母語デー。世界の言語的・文化的多様性を守るためのユネスコ制定の日。"],["2/22","猫の日。「ニャン（2）ニャン（2）ニャン（2）」の語呂合わせから制定。"],["2/23","天皇誕生日。日本の国民の祝日。"],["2/24","鉄道ストの日。1898年のこの日、日本初の鉄道ストライキが行われた。"],["2/25","夕刊紙の日。1969年のこの日、日本初の駅売り専門夕刊紙が創刊された。"],["2/26","二・二六事件の日。1936年のこの日、陸軍青年将校がクーデターを起こした。"],["2/27","新撰組の日。1863年のこの日、新撰組の前身・浪士組が江戸を出発した。"],["2/28","ビスケットの日。1855年のこの日、水戸藩士がビスケットの製法を学んだ。"],["2/29","うるう日。4年に1度だけ存在する特別な日。"],
    ["3/1","マーチの日。3月（March）にちなんで制定。"],["3/2","遠山の金さんの日。遠山金四郎が江戸北町奉行に任命された日。"],["3/3","ひな祭り。女の子の健やかな成長を願って雛人形を飾る日。"],["3/4","バウムクーヘンの日。1919年のこの日、日本初のバウムクーヘンが作られた。"],["3/5","ミスコンの日。1908年のこの日、日本初の美人コンテストが開催された。"],["3/6","スポーツ新聞の日。1946年のこの日、日本初のスポーツ新聞が創刊された。"],["3/7","消防記念日。1948年のこの日、消防組織法が施行された。"],["3/8","国際女性デー。女性の権利と社会参加を讃える国際的な記念日。"],["3/9","感謝の日。「サン（3）キュー（9）」の語呂合わせから制定。"],["3/10","砂糖の日。「さ（3）とう（10）」の語呂合わせから制定。"],["3/11","東日本大震災の日。2011年のこの日、M9.0の大地震と津波が発生した。"],["3/12","財布の日。「さ（3）いふ（12）」の語呂合わせから制定。"],["3/13","新撰組忌。1868年のこの日、近藤勇が捕縛された。"],["3/14","ホワイトデー。バレンタインデーのお返しにお菓子を贈る日。"],["3/15","靴の記念日。1870年のこの日、日本初の西洋靴の工場が開業した。"],["3/16","国立公園指定記念日。1934年のこの日、初の国立公園が指定された。"],["3/17","漫画週刊誌の日。1959年のこの日、週刊少年マガジンと週刊少年サンデーが創刊。"],["3/18","点字ブロックの日。1967年のこの日、世界初の点字ブロックが敷設された。"],["3/19","ミュージックの日。語呂合わせから制定。"],["3/20","春分の日。昼と夜の長さがほぼ等しくなる日。自然をたたえ生物を慈しむ祝日。"],["3/21","国際人種差別撤廃デー。1960年のシャープビル虐殺事件に由来する国連制定の日。"],["3/22","放送記念日。1925年のこの日、日本初のラジオ放送が行われた。"],["3/23","世界気象デー。1950年のこの日、世界気象機関条約が発効した。"],["3/24","壇ノ浦の戦いの日。1185年のこの日、源氏が平氏を滅ぼした。"],["3/25","電気記念日。1878年のこの日、日本で初めて電灯が点灯された。"],["3/26","カチューシャの歌の日。1914年のこの日、カチューシャの歌が初めて披露された。"],["3/27","さくらの日。語呂合わせから制定。"],["3/28","シルクロードの日。1900年のこの日、スウェーデンの探検家がシルクロードを命名した。"],["3/29","マリモの日。1952年のこの日、マリモが国の特別天然記念物に指定された。"],["3/30","国立競技場落成記念日。1958年のこの日、国立霞ヶ丘陸上競技場が完成した。"],["3/31","オーケストラの日。語呂合わせから制定。"],
    ["4/1","エイプリルフール。嘘をついても許されるとされる日。新年度スタートの日でもある。"],["4/2","国際子どもの本の日。アンデルセンの誕生日に合わせて制定された日。"],["4/3","シーサーの日。「シー（4）サー（3）」の語呂合わせから制定。"],["4/4","獅子の日。「し（4）し（4）」の語呂合わせから制定。"],["4/5","ヘアカットの日。1872年のこの日、兵部省が散髪を許可した。"],["4/6","城の日。「し（4）ろ（6）」の語呂合わせから制定。"],["4/7","世界保健デー。1948年のこの日、世界保健機関（WHO）が発足した。"],["4/8","花祭り。お釈迦様の誕生日。各地のお寺で甘茶をかける法要が行われる日。"],["4/9","大仏の日。752年のこの日、東大寺の大仏開眼供養が行われた。"],["4/10","女性の日。1946年のこの日、初めて女性が選挙権を行使した。"],["4/11","メートル法の日。1921年のこの日、メートル法が公布された。"],["4/12","世界宇宙飛行の日。1961年のこの日、ガガーリンが人類初の宇宙飛行を達成した。"],["4/13","喫茶店の日。1888年のこの日、日本初の喫茶店が東京・上野に開業した。"],["4/14","タイタニック号の日。1912年のこの日、タイタニック号が氷山に衝突した。"],["4/15","ヘリコプターの日。語呂合わせから制定。"],["4/16","チャップリンの日。1889年のこの日、チャーリー・チャップリンが誕生した。"],["4/17","ハイビジョンの日。語呂合わせから制定。"],["4/18","発明の日。1885年のこの日、専売特許条例が公布された。"],["4/19","地図の日。1800年のこの日、伊能忠敬が蝦夷地の測量を開始した。"],["4/20","郵政記念日。1871年のこの日、郵便制度が始まった。"],["4/21","民放の日。1951年のこの日、日本で初めて民間放送局が免許を受けた。"],["4/22","アースデー（地球の日）。地球環境について考える日。"],["4/23","子ども読書の日。子どもたちに読書の楽しさを伝えるために制定された日。"],["4/24","植物学の日。1862年のこの日、植物学者・牧野富太郎が誕生した。"],["4/25","国連記念日。1945年のこの日、国連設立のサンフランシスコ会議が開幕した。"],["4/26","よい風呂の日。語呂合わせから制定。"],["4/27","哲学の日。紀元前399年のこの日、ソクラテスが毒を飲んで亡くなった。"],["4/28","象の日。1729年のこの日、ベトナムから来た象が徳川吉宗に披露された。"],["4/29","昭和の日。激動の時代を経て復興を遂げた昭和の時代を顧みる祝日。"],["4/30","図書館記念日。1950年のこの日、図書館法が公布された。"],
    ["5/1","メーデー。労働者の祭典として世界各地でデモや集会が行われる日。"],["5/2","交通広告の日。語呂合わせから制定。"],["5/3","憲法記念日。1947年のこの日、日本国憲法が施行された。"],["5/4","みどりの日。自然に親しみその恩恵に感謝する祝日。"],["5/5","こどもの日。子どもの人格を重んじ幸福を祈る祝日。端午の節句でもある。"],["5/6","コロコロの日。語呂合わせから制定。"],["5/7","コナモンの日。語呂合わせから制定。"],["5/8","世界赤十字デー。赤十字の創設者アンリ・デュナンの誕生日。"],["5/9","アイスクリームの日。1964年のこの日、アイスクリームの宣伝活動が始まった。"],["5/10","日本気象協会創立記念日。1950年のこの日、日本気象協会が設立された。"],["5/11","ご当地キャラの日。語呂合わせから制定。"],["5/12","看護の日。近代看護の祖フローレンス・ナイチンゲールの誕生日。"],["5/13","カクテルの日。1806年のこの日、カクテルの語源となる記事が初めて掲載された。"],["5/14","温度計の日。1686年のこの日、ファーレンハイトが誕生した。"],["5/15","沖縄本土復帰記念日。1972年のこの日、沖縄が日本に返還された。"],["5/16","旅の日。1689年のこの日、松尾芭蕉が奥の細道の旅に出発した。"],["5/17","世界通信の日。1865年のこの日、万国電信連合が創設された。"],["5/18","国際博物館の日。博物館の社会における役割を普及するための国際的な記念日。"],["5/19","ボクシングの日。語呂合わせから制定。"],["5/20","ローマ字の日。1954年のこの日、ローマ字教育の振興を目的として制定された。"],["5/21","小満。草木が茂り自然が満ち足りてくる時期を表す二十四節気の一つ。"],["5/22","国際生物多様性の日。地球上の生物多様性の保全を考える日。"],["5/23","世界亀の日。亀の保護と知識の普及を目的として制定された日。"],["5/24","伊達巻の日。伊達政宗の命日にちなんで制定。"],["5/25","アフリカデー。1963年のこの日、アフリカ統一機構が設立された。"],["5/26","ル・マンの日。1923年のこの日、第1回ル・マン24時間耐久レースが開催された。"],["5/27","百人一首の日。1235年のこの日、藤原定家が百人一首を完成させた。"],["5/28","花火の日。1733年のこの日、隅田川で初めて花火大会が開催された。"],["5/29","エベレスト登頂記念日。1953年のこの日、ヒラリーとテンジンが初登頂に成功した。"],["5/30","ごみゼロの日。「ご（5）み（3）ゼロ（0）」の語呂合わせから制定。"],["5/31","世界禁煙デー。WHO（世界保健機関）が制定した禁煙を推進する日。"],
    ["6/1","電波の日。1950年のこの日、電波法が施行された。"],["6/2","横浜港開港記念日。1859年のこの日、横浜港が開港した。"],["6/3","測量の日。1949年のこの日、測量法が公布された。"],["6/4","虫歯予防デー。「む（6）し（4）」の語呂合わせから制定。"],["6/5","環境の日。世界環境デー。地球環境保全について考える日。"],["6/6","おけいこの日。6歳の6月6日からおけいこを始めると上達が早いという言い伝えから。"],["6/7","緑内障を考える日。語呂合わせから制定。"],["6/8","成層圏の日。1923年のこの日、成層圏の存在が確認された。"],["6/9","ロックの日。「ロ（6）ック（9）」の語呂合わせから制定。"],["6/10","時の記念日。671年のこの日、日本初の時計（水時計）が時を刻んだとされる日。"],["6/11","傘の日。梅雨入りの時期に合わせて制定された日。"],["6/12","日記の日。1942年のこの日、アンネ・フランクが日記を書き始めた。"],["6/13","小さな親切の日。1963年のこの日、「小さな親切」運動が始まった。"],["6/14","世界献血者デー。血液の輸血・血液製剤の重要性を啓発する日。"],["6/15","暑中見舞いの日。梅雨明けが近づくこの時期に制定された。"],["6/16","和菓子の日。848年のこの日、和菓子で疫病退散を祈願したことに由来。"],["6/17","砂漠化と干ばつと闘う世界デー。国連が制定した環境問題を考える日。"],["6/18","おにぎりの日。石川県で弥生時代のおにぎりが発見されたことから制定。"],["6/19","ベースボールの日。1846年のこの日、アメリカで初の野球試合が行われた。"],["6/20","ペパーミントの日。北海道北見市のハッカ産業に由来して制定。"],["6/21","夏至。一年で最も昼が長い日。二十四節気の一つ。"],["6/22","ボウリングの日。1861年のこの日、日本初のボウリング場が長崎に開業した。"],["6/23","沖縄慰霊の日。1945年のこの日、沖縄戦が終結した。"],["6/24","UFOの日。1947年のこの日、アメリカでUFO目撃情報が報告された。"],["6/25","住宅デー。語呂合わせから制定。"],["6/26","国際薬物乱用・不正取引防止デー。国連が制定した薬物問題を考える日。"],["6/27","ちらし寿司の日。江戸時代にちらし寿司が広まったことに由来。"],["6/28","パフェの日。1950年のこの日、プロ野球で初のパーフェクトゲームが達成された。"],["6/29","ビートルズ記念日。1966年のこの日、ビートルズが来日した。"],["6/30","ハーフタイムデー。一年365日のちょうど折り返しにあたる日。"],
    ["7/1","国民安全の日。交通事故や労働災害などの防止を呼びかける日。"],["7/2","たわしの日。1915年のこの日、亀の子たわしが特許登録された。"],["7/3","ソフトクリームの日。1951年のこの日、日本で初めてソフトクリームが販売された。"],["7/4","アメリカ独立記念日。1776年のこの日、アメリカ独立宣言が採択された。"],["7/5","江戸切子の日。江戸切子の美しいカットガラスを広める日として制定。"],["7/6","公認会計士の日。1948年のこの日、公認会計士法が制定された。"],["7/7","七夕。織姫と彦星が天の川を渡って年に一度会えるとされる日。"],["7/8","なわとびの日。語呂合わせから制定。"],["7/9","ジェットコースターの日。1955年のこの日、後楽園遊園地にジェットコースターが登場した。"],["7/10","納豆の日。「な（7）っとう（10）」の語呂合わせから制定。"],["7/11","セブンイレブンの日。セブンイレブンの創業日にちなんで制定。"],["7/12","ラジオ本放送の日。1925年のこの日、東京放送局がラジオの本放送を開始した。"],["7/13","盆迎え火。祖先の霊を迎えるために火を焚く日。"],["7/14","パリ祭。1789年のこの日、フランス革命が始まった。"],["7/15","お盆。祖先の霊を迎え供養する日。"],["7/16","虹の日。「な（7）な（7）い（1）ろ（6）」の語呂合わせから制定。"],["7/17","東京の日。1868年のこの日、明治天皇が江戸を東京と改めた。"],["7/18","光化学スモッグの日。1970年のこの日、東京で光化学スモッグが発生した。"],["7/19","サイボーグの日。漫画「サイボーグ009」の連載が1964年に始まったことから制定。"],["7/20","月面着陸の日。1969年のこの日、アポロ11号が月面に着陸した。"],["7/21","神前結婚記念日。1900年のこの日、初めての神前結婚式が行われた。"],["7/22","下駄の日。語呂合わせから制定。"],["7/23","米騒動記念日。1918年のこの日、富山県で米騒動が起きた。"],["7/24","劇画の日。1964年のこの日、劇画誌「ガロ」が創刊された。"],["7/25","かき氷の日。「な（7）つ（2）ご（5）おり」の語呂合わせから制定。"],["7/26","幽霊の日。1825年のこの日、東海道四谷怪談が初演された。"],["7/27","スイカの日。縞模様を綱に見立てて語呂合わせから制定。"],["7/28","世界肝炎デー。肝炎ウイルス感染の予防と治療を啓発する国際的な記念日。"],["7/29","福井県民の日。1871年のこの日、廃藩置県で福井県が誕生した。"],["7/30","梅干しの日。語呂合わせから制定。"],["7/31","こだまの日。語呂合わせから制定。"],
    ["8/1","水の日。日本の水資源の大切さについて考える日として制定。"],["8/2","ハーブの日。「ハ（8）ーブ（2）」の語呂合わせから制定。"],["8/3","ハチミツの日。語呂合わせから制定。"],["8/4","橋の日。「は（8）し（4）」の語呂合わせから制定。"],["8/5","はしの日。語呂合わせから制定。"],["8/6","広島平和記念日。1945年のこの日、広島に原子爆弾が投下された。"],["8/7","花の日。「は（8）な（7）」の語呂合わせから制定。"],["8/8","そろばんの日。「パチ（8）パチ（8）」の語呂合わせから制定。"],["8/9","長崎原爆の日。1945年のこの日、長崎に原子爆弾が投下された。"],["8/10","道の日。1920年のこの日、日本初の近代的な道路整備計画が発足した。"],["8/11","山の日。山に親しむ機会を得て山の恩恵に感謝する祝日。"],["8/12","航空安全の日。1985年のこの日、日航機墜落事故が起きた。"],["8/13","盆迎え火。祖先の霊を迎えるために各家庭で火を焚く日。"],["8/14","専売特許の日。1888年のこの日、日本初の特許が交付された。"],["8/15","終戦記念日。1945年のこの日、日本がポツダム宣言を受諾し戦争が終わった。"],["8/16","盆送り火。祖先の霊を送り出すために火を焚く日。"],["8/17","パイナップルの日。語呂合わせから制定。"],["8/18","高校野球記念日。1915年のこの日、第1回全国中等学校優勝野球大会が開催された。"],["8/19","俳句の日。「は（8）い（1）く（9）」の語呂合わせから制定。"],["8/20","蚊の日。1897年のこの日、ロスが蚊がマラリアを媒介することを発見した。"],["8/21","噴水の日。1877年のこの日、日本初の噴水が上野公園に完成した。"],["8/22","チンチン電車の日。1903年のこの日、日本初の電車が大阪で走った。"],["8/23","処暑。暑さが峠を越えて落ち着いてくる時期を示す二十四節気の一つ。"],["8/24","ポンペイ最後の日。79年のこの日、ベスビオ火山が噴火しポンペイが埋没した。"],["8/25","東京国際空港開港記念日。1931年のこの日、羽田空港が開港した。"],["8/26","人権宣言記念日。1789年のこの日、フランスで人権宣言が採択された。"],["8/27","男はつらいよの日。1969年のこの日、映画「男はつらいよ」が公開された。"],["8/28","民放テレビスタートの日。1953年のこの日、日本初の民放テレビが放送を開始した。"],["8/29","焼き肉の日。「や（8）き（2）にく（9）」の語呂合わせから制定。"],["8/30","冒険家の日。多くの冒険家がこの日に偉業を達成したことから制定。"],["8/31","野菜の日。「や（8）さ（3）い（1）」の語呂合わせから制定。"],
    ["9/1","防災の日。1923年のこの日、関東大震災が発生した。防災意識を高める日。"],["9/2","宝くじの日。「く（9）じ（2）」の語呂合わせから制定。"],["9/3","グミの日。「グ（9）ミ（3）」の語呂合わせから制定。"],["9/4","串の日。「く（9）し（4）」の語呂合わせから制定。"],["9/5","石炭の日。語呂合わせから制定。"],["9/6","黒の日。「く（9）ろ（6）」の語呂合わせから制定。"],["9/7","クリーナーの日。語呂合わせから制定。"],["9/8","国際識字デー。ユネスコが制定した読み書き能力の大切さを考える日。"],["9/9","重陽の節句。菊の花を飾り長寿と繁栄を願う日。"],["9/10","下水道の日。1961年のこの日、下水道の大切さを訴えるために制定された。"],["9/11","公衆電話の日。1900年のこの日、日本初の公衆電話が東京に設置された。"],["9/12","マラソンの日。紀元前490年のこの日、マラソンの戦いが起きた。"],["9/13","世界の法の日。1965年のこの日、法の支配に関する国際会議が宣言を採択した。"],["9/14","コスモスの日。秋の花コスモスが咲き誇る時期に制定された日。"],["9/15","老人の日。敬老の日の前身として高齢者を敬う日。"],["9/16","マッチの日。1948年のこの日、マッチの配給制が廃止された。"],["9/17","モノレールの日。1964年のこの日、東京モノレールが開業した。"],["9/18","かいわれ大根の日。語呂合わせから制定。"],["9/19","苗字の日。1870年のこの日、平民が苗字を持つことを許された。"],["9/20","空の日。1911年のこの日、日本で初めて飛行機が飛んだ。"],["9/21","国際平和デー。国連が制定した世界の平和を考える日。"],["9/22","カーフリーデー。車のない生活を体験して環境について考える日。"],["9/23","秋分の日。昼と夜の長さがほぼ等しくなる日。祖先を敬い故人を偲ぶ祝日。"],["9/24","畳の日。語呂合わせから制定。"],["9/25","主婦休みの日。日頃頑張る主婦がリフレッシュする日として制定。"],["9/26","ワープロの日。1978年のこの日、日本初のワードプロセッサが発売された。"],["9/27","世界観光の日。国連世界観光機関が制定した観光を通じた交流を促進する日。"],["9/28","パソコンの日。1979年のこの日、日本初のパソコンが発売された。"],["9/29","招き猫の日。「来る（9）福（29→ふく）」の語呂合わせから制定。"],["9/30","クレーンの日。1952年のこの日、クレーン等安全規則が施行された。"],
    ["10/1","国際音楽の日。音楽の素晴らしさを世界に広めるために制定された日。"],["10/2","豆腐の日。語呂合わせから制定。"],["10/3","登山の日。「と（10）ざん（3→さん）」の語呂合わせから制定。"],["10/4","世界動物の日。動物の権利と福祉を守るための国際デー。"],["10/5","時刻表記念日。1894年のこの日、日本初の時刻表が発行された。"],["10/6","国際協力の日。1954年のこの日、日本がコロンボ計画に加盟した。"],["10/7","ミステリー記念日。1849年のこの日、推理小説の父エドガー・アラン・ポーが亡くなった。"],["10/8","入れ歯の日。語呂合わせから制定。"],["10/9","世界郵便デー。1874年のこの日、万国郵便連合が設立された。"],["10/10","目の愛護デー。「1010」が目に見えることから制定。"],["10/11","安全・安心なまちづくりの日。地域の安全を考える日として制定。"],["10/12","コロンブスデー。1492年のこの日、コロンブスがアメリカ大陸を発見した。"],["10/13","サツマイモの日。語呂合わせから制定。"],["10/14","鉄道の日。1872年のこの日、日本初の鉄道が新橋〜横浜間で開業した。"],["10/15","世界手洗いの日。石けんを使った手洗いの普及を促進する国際デー。"],["10/16","世界食料デー。飢餓のない世界を目指してFAOが制定した日。"],["10/17","貯蓄の日。貯蓄意識を高める日として制定。"],["10/18","冷凍食品の日。語呂合わせから制定。"],["10/19","バーゲンの日。1895年のこの日、日本初のバーゲンセールが開催された。"],["10/20","新聞広告の日。新聞広告の役割と重要性を知る日として制定。"],["10/21","あかりの日。1879年のこの日、エジソンが白熱電球の実験に成功した。"],["10/22","平安遷都の日。794年のこの日、桓武天皇が平安京に遷都した。"],["10/23","電信電話記念日。1869年のこの日、日本初の電信が開通した。"],["10/24","国際連合デー。1945年のこの日、国際連合が発足した。"],["10/25","世界パスタデー。1995年のこの日、世界初のパスタ会議が開催された。"],["10/26","サーカスの日。1871年のこの日、日本初の西洋サーカスが来日した。"],["10/27","文字・活字文化の日。読書週間の最初の日として制定された。"],["10/28","速記記念日。1882年のこの日、日本初の速記法が発表された。"],["10/29","おしぼりの日。語呂合わせから制定。"],["10/30","初恋の日。1896年のこの日、島崎藤村が初恋の詩を発表した。"],["10/31","ハロウィン。仮装やお菓子のイベントが楽しい西洋の秋祭りの日。"],
    ["11/1","計量記念日。1993年のこの日、新計量法が施行された。"],["11/2","キッチン・バスの日。語呂合わせから制定。"],["11/3","文化の日。自由と平和を愛し文化をすすめる国民の祝日。"],["11/4","ユネスコ憲章記念日。1946年のこの日、ユネスコ憲章が発効した。"],["11/5","津波防災の日。世界津波の日として国連が制定した防災を考える日。"],["11/6","お見合い記念日。1947年のこの日、集団お見合いが初めて開催された。"],["11/7","立冬。暦の上で冬が始まる日。二十四節気の一つ。"],["11/8","いい歯の日。「い（1）い（1）は（8）」の語呂合わせから制定。"],["11/9","119番の日。火災・救急を呼ぶ日として制定。"],["11/10","エレベーターの日。1890年のこの日、日本初の電動エレベーターが運転を開始した。"],["11/11","ポッキーの日。「1111」の形がポッキーに見えることから制定。"],["11/12","洋服記念日。1872年のこの日、太政官布告で洋服着用が正式に認められた。"],["11/13","うるしの日。弘法大師が中国からうるしの技術を伝えたとされる日。"],["11/14","世界糖尿病デー。糖尿病の予防と治療の啓発を促進する国際的な記念日。"],["11/15","七五三。3歳・5歳・7歳の子どもの成長を祝い、神社に参拝する日。"],["11/16","国際寛容デー。ユネスコが制定した寛容と多様性を推進する日。"],["11/17","蓮根の日。茨城県土浦市のれんこん祭りに合わせて制定。"],["11/18","土木の日。語呂合わせから制定。"],["11/19","世界トイレの日。トイレと衛生問題について考える国連制定の日。"],["11/20","世界子どもの日。1959年のこの日、国連が「子どもの権利宣言」を採択した。"],["11/21","世界テレビデー。1996年のこの日、国連で世界テレビ・フォーラムが開催された。"],["11/22","いい夫婦の日。「い（1）い（1）ふ（2）うふ（2）」の語呂合わせから制定。"],["11/23","勤労感謝の日。勤労をたっとび生産を祝い、国民がたがいに感謝しあう祝日。"],["11/24","オペラ記念日。1894年のこの日、日本初のオペラが上演された。"],["11/25","OLの日。「OL」の語源となった出来事に由来して制定。"],["11/26","いい風呂の日。「い（1）い（1）ふ（2）ろ（6）」の語呂合わせから制定。"],["11/27","ノーベル賞制定記念日。1895年のこの日、アルフレッド・ノーベルが遺言状に署名した。"],["11/28","税関記念日。1872年のこの日、現在の税関の基礎が定まった。"],["11/29","いい肉の日。「い（1）い（1）に（2）く（9）」の語呂合わせから制定。"],["11/30","本みりんの日。語呂合わせから制定。"],
    ["12/1","映画の日。1896年のこの日、日本初の映画が公開された。"],["12/2","日本アニメの日。日本のアニメ文化を讃えて制定された日。"],["12/3","国際障害者デー。1992年に国連が制定した障害者への理解を深める日。"],["12/4","E.T.の日。1982年のこの日、映画「E.T.」が日本で公開された。"],["12/5","バミューダトライアングルの日。1945年のこの日、バミューダ海域で飛行機が消息を絶った。"],["12/6","姉の日。「あ（12）ね（6）」の語呂合わせから制定。"],["12/7","大雪。雪が激しく降り始める頃を表す二十四節気の一つ。"],["12/8","太平洋戦争開戦記念日。1941年のこの日、日本がアメリカに宣戦布告した。"],["12/9","障害者の日。1975年のこの日、国連が障害者の権利宣言を採択した。"],["12/10","ノーベル賞授賞式。アルフレッド・ノーベルの命日であるこの日に毎年授賞式が行われる。"],["12/11","胃腸の日。語呂合わせから制定。"],["12/12","漢字の日。語呂合わせから制定。"],["12/13","煤払いの日。江戸時代からの慣習で、年末の大掃除をする日。"],["12/14","赤穂浪士討ち入りの日。1702年のこの日、四十七士が吉良邸に討ち入りした。"],["12/15","観光バス記念日。1925年のこの日、日本初の観光バスが運行された。"],["12/16","電話創業の日。1890年のこの日、日本で電話が開通した。"],["12/17","飛行機の日。1903年のこの日、ライト兄弟が人類初の動力飛行に成功した。"],["12/18","国連加盟記念日。1956年のこの日、日本が国連に加盟した。"],["12/19","日本人初飛行の日。1910年のこの日、日本人初の飛行機による飛行が行われた。"],["12/20","道路交通法施行記念日。1960年のこの日、道路交通法が施行された。"],["12/21","冬至。一年で最も夜が長い日。かぼちゃを食べゆず湯に入る風習がある。"],["12/22","労働組合法制定記念日。1945年のこの日、労働組合法が公布された。"],["12/23","テレコムの日。1933年のこの日、国際電気通信連合が設立された。"],["12/24","クリスマスイブ。サンタクロースがプレゼントを届けるとされる特別な夜。"],["12/25","クリスマス。イエス・キリストの誕生を祝うキリスト教の祭日。"],["12/26","プロ野球誕生の日。1934年のこの日、日本初のプロ野球チームが誕生した。"],["12/27","ピーターパンの日。1904年のこの日、舞台「ピーターパン」が初演された。"],["12/28","シネマの日。映画の歴史を振り返り映画文化を讃える日。"],["12/29","清水トンネル貫通記念日。1929年のこの日、上越線の清水トンネルが貫通した。"],["12/30","地下鉄記念日。1927年のこの日、日本初の地下鉄が上野〜浅草間で開業した。"],["12/31","大晦日。一年の最後の日。年越し蕎麦を食べ新年を待つ日。"],
  ];

  const trivia = [
    "🐙 タコは心臓を3つ持っている。2つは鰓心臓、1つは全身に血液を送る体心臓。",
    "🍯 はちみつは腐らない。古代エジプトの墓から3000年前のはちみつが発見されている。",
    "🌍 地球上の砂浜の砂の数より、宇宙の星の数のほうが多いと言われている。",
    "🐘 ゾウは唯一ジャンプできない哺乳類とされている。",
    "🦷 人間の歯のエナメル質は体の中で最も硬い組織。",
    "🍌 バナナは植物学的には木ではなく草本植物（ハーブ）の一種。",
    "💤 人間は人生の約3分の1を睡眠に費やす。",
    "🐬 イルカは眠るとき、脳の半分ずつ交互に休ませる「半球睡眠」を行う。",
    "🌙 月は地球から毎年約3.8cm遠ざかっている。",
    "🦋 チョウの味覚センサーは足についている。花に止まるだけで甘さがわかる。",
    "🧠 人間の脳は体重の約2%しかないが、消費カロリーの約20%を使う。",
    "🐟 金魚の記憶力は3秒ではなく、実際には数ヶ月間記憶できるとされている。",
    "🌿 1本の成木は1年間に約22kgのCO2を吸収すると言われている。",
    "🎵 音楽を聴くと脳内でドーパミンが分泌され、気分が向上する。",
    "☕ コーヒーの香りだけでストレスが軽減されるという研究がある。",
    "🐝 ハチは地球上で最も重要な生物の一つ。食料の約3分の1はハチの受粉に依存。",
    "📚 読書は認知症予防に効果があるとされている。",
    "🌈 虹は実際には円形。地平線があるため半円にしか見えない。",
    "🦈 サメは軟骨魚類で骨がない。骨格はすべて軟骨でできている。",
    "🍎 リンゴを水に入れると浮く。約25%が空気でできているため。",
    "🦒 キリンの舌は約45cmで黒っぽい色をしている。紫外線から守るためとされる。",
    "🐦 フラミンゴがピンク色なのは食べるエビや藻に含まれるカロチノイドのため。",
    "🌡️ 金属は冷えると縮む。エッフェル塔は冬に約15cm低くなる。",
    "🐌 カタツムリは歯を持っており、その数は約14000本以上にもなる。",
    "🦑 イカの血液は青い。銅を含むヘモシアニンという成分のため。",
    "🌺 チューリップはユリ科の植物で、玉ねぎと同じ仲間。",
    "🐧 ペンギンは空を飛べないが、水中では時速25km以上で泳げる。",
    "🔥 炎は重力がないと球形になる。宇宙空間での炎は丸い。",
    "🦅 ワシは人間の約4〜8倍の視力を持つとされている。",
    "🌊 海の最深部マリアナ海溝は約11000m。エベレストよりも深い。",
    "🐢 亀は5億年以上前から存在する。恐竜より古い生き物。",
    "🍕 イタリアのピザは世界無形文化遺産に登録されている。",
    "📱 スマートフォンのガラスは砂から作られている。",
    "🦁 ライオンのタテガミは成熟度を示す。黒いほど強い個体が多い。",
    "🌴 ヤシの木は熱帯のイメージだが、実は草本植物に分類される。",
    "🐊 ワニは舌が口の底に固定されており、舌を動かせない。",
    "💎 ダイヤモンドは炭素の固まり。鉛筆の芯（グラファイト）と同じ元素でできている。",
    "🎸 ギターの弦を弾いたときの振動数によって音の高さが決まる。",
    "🧊 氷は水よりも軽い。そのため氷は水に浮く。",
    "🦠 人体には人間の細胞の数より多い細菌（腸内細菌など）が存在する。",
    "🐻 クマは冬眠中も体温はほとんど下がらず、脈拍が下がるだけ。",
    "🌏 地球の中心部の温度は太陽の表面温度とほぼ同じ約6000℃。",
    "🐸 カエルは水を飲まず、皮膚から水分を吸収する。",
    "🎯 ダーツのボードの数字の合計は501。",
    "🦔 ハリネズミのトゲは中が空洞になっており、軽くて丈夫な構造をしている。",
    "🌸 桜の花びらは5枚が基本だが、八重桜は100枚以上のものもある。",
    "🐠 金魚は元々フナを品種改良したもの。野生には存在しない。",
    "⚡ 落雷は同じ場所に何度も落ちることがある。エンパイアステートビルには年平均20〜25回落ちる。",
    "🐝 ハチの羽は1秒間に約200回羽ばたく。それが「ブーン」という音の正体。",
    "🧲 地球自体が巨大な磁石。北極が磁石のS極にあたる。",
    "🐡 フグの毒テトロドトキシンはフグ自身が作るのではなく、エサとなる細菌から蓄積する。",
    "🌰 栗は果物ではなく種子。私たちが食べる部分は種にあたる。",
    "🐕 犬の鼻紋（鼻のしわ）は人間の指紋と同様、個体を識別できる。",
    "🎻 バイオリンの弦は昔は羊の腸で作られていた。今は主にスチール製。",
    "🌱 植物は音楽を聴くと成長が促進されるという研究がある。",
    "🐰 ウサギは自分の糞を食べる。栄養を二度吸収するための行動。",
    "🔭 光が太陽から地球に届くまで約8分20秒かかる。",
    "🐾 ネコの前足の指は5本、後ろ足は4本。",
    "🌵 サボテンは幹に水を蓄えているのではなく、細胞に水分を貯めている。",
    "🎈 ヘリウムを吸うと声が高くなるのは、音速がヘリウム中では速いため。",
    "🦩 フラミンゴは片足で立つと疲れにくい。脚の構造が自然にロックされるため。",
    "🍫 チョコレートは犬に有毒。テオブロミンという成分が犬には分解できない。",
    "🐙 タコは無脊椎動物の中で最も知能が高いとされ、ビンの蓋を開けることもできる。",
    "🌑 月には大気がないため、宇宙飛行士が残した足跡は何百万年も消えない。",
    "🐃 闘牛で赤い布に反応するのは色ではなく動きのため。牛は赤色も認識できる。",
    "🎃 ハロウィンのカボチャは元々カブで作られていた。アメリカでカボチャに変わった。",
    "🦞 ロブスターは老化しないとされている。年を取るほど繁殖力が増す。",
    "🏔️ エベレストは地球の中心から最も遠い山ではない。エクアドルのチンボラソ山が最遠。",
    "🌊 海底の山脈（中央海嶺）の総延長は約65000km。地球一周の約1.6倍。",
    "🐝 女王蜂は1日に最大2000個の卵を産む。",
    "🍵 お茶の起源は中国。伝説では紀元前2700年頃に発見されたとされている。",
    "🔬 人体の全DNA分子を1本につなぐと地球から太陽まで往復できる長さになる。",
    "🐝 蜂蜜1瓶（500g）を作るために、ミツバチは約200万回の往復をする。",
    "🌿 竹は世界で最も成長が早い植物。1日に最大91cm成長するものもある。",
    "🧬 チンパンジーのDNAは人間と約98.7%一致する。",
    "🐠 魚の種類は哺乳類・爬虫類・鳥類・両生類の合計より多い。",
    "🎭 能面は微妙な角度の変化で喜怒哀楽を表現できるよう設計されている。",
    "🌍 アフリカは世界の陸地の約20%を占めるが、地図では実際より小さく見える。",
    "🐋 クジラの歌は数百km先まで届く。特にザトウクジラの歌は複雑で学習される。",
    "🧊 南極の氷床に閉じ込められた空気の泡には、数十万年前の大気が含まれている。",
    "🌸 日本のソメイヨシノはすべてクローン。一本の木から接ぎ木で増やされた。",
    "🐜 アリは自分の体重の50倍以上の重さを運ぶことができる。",
    "🌙 月の重力は地球の約6分の1。月面では体重が6分の1になる。",
    "🐦 ハチドリは唯一後ろ向きに飛べる鳥。羽を1秒間に80回ばたつかせる。",
    "💧 地球上の水の約97%は海水。淡水はわずか3%しかない。",
    "🐓 鶏は色を識別できる。人間より多くの種類の色を見分けられる。",
    "🌺 蓮の種は数千年後も発芽できる可能性がある。古代の蓮の種が発芽した例がある。",
    "🦭 アザラシは陸上では不格好だが、水中では時速40km以上で泳げる。",
    "🎨 ピカソは生涯に約15万点以上の作品を制作したとされる。",
    "🍄 キノコは植物でも動物でもなく、菌類という独立した生物の仲間。",
    "⭐ 夜空に見える星のほとんどは、実際にはすでに存在しない可能性がある。光が届くのに数百万年かかるため。",
    "🦎 カメレオンの色変化は擬態だけでなく、体温調節やコミュニケーションのためでもある。",
    "🌊 tsunamiは日本語から来た言葉。世界共通語として使われている。",
    "🐺 オオカミは犬の祖先。約1万5000年前に人間に家畜化されたとされる。",
    "🍙 日本のコンビニおにぎりの包装は、海苔がパリッとした状態を保てるよう特殊設計されている。",
    "🔑 最古の鍵は約4000年前の古代エジプトで発明された木製の鍵とピン錠。",
    "🧸 テディベアは米国大統領セオドア・ルーズベルトの愛称「テディ」に由来する。",
  ];

  const key        = `${month}/${day}`;
  const todayEntry = whatsDays.find(w => w[0] === key);
  const whatsDay   = todayEntry
    ? `📅 ${dateStr}は…\n${todayEntry[1]}`
    : `📅 ${dateStr}は…\n記録には残っていませんが、今日もきっと特別な一日！`;
  const seed        = month * 100 + day;
  const todayTrivia = `💡 今日の豆知識\n${trivia[seed % trivia.length]}`;

  sh.getRange(1, 4, 1, 2).merge();
  sh.getRange(1, 4).setValue(`✨ Today ${dateStr}`)
    .setFontWeight("bold").setFontSize(12).setBackground("#fce8e6").setHorizontalAlignment("center");
  sh.getRange(2, 4, 1, 2).merge();
  sh.getRange(2, 4).setValue(whatsDay).setWrap(true).setFontSize(11).setBackground("#fff8e1").setVerticalAlignment("top");
  sh.setRowHeight(2, 80);
  sh.getRange(4, 4, 1, 2).merge();
  sh.getRange(4, 4).setValue(todayTrivia).setWrap(true).setFontSize(11).setBackground("#e8f5e9").setVerticalAlignment("top");
  sh.setRowHeight(4, 120);
  sh.setColumnWidth(4, 200);
  sh.setColumnWidth(5, 10);
}
