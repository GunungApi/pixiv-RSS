/**
 * スケジューラーMVP Apps Script
 * 親用フォームからの候補日時を承認し、Googleカレンダーに登録します。
 * 設定値はスプレッドシートの「設定」シートに保存されます。
 */

const SETTINGS_SHEET_NAME = '設定';
const FORM_RESPONSES_SHEET_NAME = 'フォーム回答';
const AUDIT_LOG_SHEET_NAME = '監査ログ';

/**
 * 初期セットアップ：フォームとシート、トリガーを作成
 */
function setup() {
  const ss = SpreadsheetApp.getActive();

  // 設定シート作成
  let settings = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!settings) {
    settings = ss.insertSheet(SETTINGS_SHEET_NAME);
    settings.getRange('A1:B10').setValues([
      ['カレンダーID', ''],
      ['件名フォーマット', '[生徒名] 科目 / 対面'],
      ['移動バッファ(分)', '0'],
      ['リマインド24h前', 'true'],
      ['リマインド2h前', 'true'],
      ['受付期間', '当月'],
    ]);
  }

  // 監査ログシート
  if (!ss.getSheetByName(AUDIT_LOG_SHEET_NAME)) {
    const logSheet = ss.insertSheet(AUDIT_LOG_SHEET_NAME);
    logSheet.appendRow(['タイムスタンプ','操作','トークン','候補日','承認者']);
  }

  // フォーム作成
  let form = FormApp.create('面談申込フォーム');
  form.setDescription('希望日時を選択し、備考があれば入力してください。');

  // トークン（家ごと）
  form.addTextItem().setTitle('トークン').setRequired(true);
  // 生徒名
  form.addTextItem().setTitle('生徒名').setRequired(true);
  // 同月候補複数選択
  form.addDateTimeItem().setTitle('候補日時').setRequired(true).setHelpText('同じ月の候補を複数送信してください。');
  // 備考
  form.addParagraphTextItem().setTitle('備考');

  form.setAllowResponseEdits(false);
  form.setCollectEmail(true);

  const formUrl = form.getPublishedUrl();
  const sheet = ss.getSheetByName(FORM_RESPONSES_SHEET_NAME) || ss.insertSheet(FORM_RESPONSES_SHEET_NAME);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  Logger.log('フォームURL: ' + formUrl);
}

/**
 * 承認画面用 Web アプリ
 */
function doGet() {
  const template = HtmlService.createTemplateFromFile('approve');
  return template.evaluate().setTitle('承認画面').setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

/**
 * 承認対象の一覧を取得
 */
function getPendingRequests() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(FORM_RESPONSES_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const tokenIndex = headers.indexOf('トークン');
  const nameIndex = headers.indexOf('生徒名');
  const dateIndex = headers.indexOf('候補日時');
  const noteIndex = headers.indexOf('備考');
  const statusIndex = headers.indexOf('ステータス');

  const results = [];
  data.forEach((row, i) => {
    if (row[statusIndex] !== '承認済み') {
      results.push({
        row: i + 2,
        token: row[tokenIndex],
        student: row[nameIndex],
        datetime: row[dateIndex],
        note: row[noteIndex]
      });
    }
  });
  return results;
}

/**
 * 承認処理
 */
function approveRequests(rows) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(FORM_RESPONSES_SHEET_NAME);
  const settings = getSettings();
  const cal = CalendarApp.getCalendarById(settings.calendarId);

  rows.forEach(r => {
    const row = sheet.getRange(r,1,1,sheet.getLastColumn()).getValues()[0];
    const email = row[1];
    const student = row[2];
    const datetime = new Date(row[3]);
    const note = row[4];

    // 競合チェック
    const start = new Date(datetime.getTime() - settings.bufferMinutes*60000);
    const end = new Date(datetime.getTime() + settings.bufferMinutes*60000);
    const conflicts = cal.getEvents(start, end);
    if (conflicts.length) {
      throw new Error('既存の予定と競合しています: ' + student + ' ' + datetime);
    }

    const title = settings.titleFormat.replace('[生徒名]', student);
    const event = cal.createEvent(title, datetime, new Date(datetime.getTime()+60*60*1000));
    event.addGuest(email);
    if (settings.reminder24h) event.addPopupReminder(24*60);
    if (settings.reminder2h) event.addPopupReminder(2*60);

    sheet.getRange(r, sheet.getLastColumn()+1).setValue('承認済み');
    logAudit('承認', row[0], datetime, Session.getActiveUser().getEmail());
    GmailApp.sendEmail(email, '面談が確定しました', '以下の日程で確定しました:\n'+datetime+'\n変更・キャンセルは次のリンクから:\n'+scriptUrl());
  });
}

/**
 * 設定取得
 */
function getSettings() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SETTINGS_SHEET_NAME);
  const values = sheet.getRange('A1:B10').getValues();
  const obj = {};
  values.forEach(r => obj[r[0]] = r[1]);
  return {
    calendarId: obj['カレンダーID'],
    titleFormat: obj['件名フォーマット'],
    bufferMinutes: Number(obj['移動バッファ(分)'] || 0),
    reminder24h: obj['リマインド24h前'] === 'true',
    reminder2h: obj['リマインド2h前'] === 'true',
    receptionSpan: obj['受付期間']
  };
}

/**
 * 監査ログ
 */
function logAudit(action, token, date, user) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(AUDIT_LOG_SHEET_NAME);
  sheet.appendRow([new Date(), action, token, date, user]);
}

function scriptUrl() {
  return ScriptApp.getService().getUrl();
}

