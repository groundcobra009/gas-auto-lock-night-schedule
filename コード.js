/**
 * スプレッドシートを開いたときにメニューを追加
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('予定自動ブロック')
    .addItem('設定シートを作成', 'initializeSettings')
    .addItem('予定自動ブロック実行', 'autoBlockEvening')
    .addSeparator()
    .addItem('3時間ごと自動ブロックON', 'setBlockTrigger')
    .addItem('3時間ごと自動ブロックOFF', 'deleteBlockTrigger')
    .addSeparator()
    .addItem('朝5時ロック削除ON', 'setDeleteTrigger')
    .addItem('朝5時ロック削除OFF', 'deleteDeleteTrigger')
    .addSeparator()
    .addItem('前日の予定を手動削除', 'manualDeletePreviousDayBlocks')
    .addToUi();
}

/**
 * 設定シートを自動生成する
 */
function initializeSettings() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName('設定');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('設定');
  
  // ヘッダー
  sheet.getRange('A1:B1').setValues([[
    '設定項目', '値'
  ]]);
  // デフォルト値
  const settings = [
    ['カレンダーID', 'ここにカレンダーIDを入力'],
    ['キーワード', '飲,懇親,宴,パーティ,会食,交流,親睦,打ち上げ'],
    ['ブロック開始', '18:30'],
    ['ブロック終了', '21:00'],
    ['検索日数（未来）', '30']
  ];
  sheet.getRange(2, 1, settings.length, 2).setValues(settings);
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 300);
  sheet.getRange('A1:B1').setFontWeight('bold').setBackground('#f3f3f3');
  SpreadsheetApp.getUi().alert('設定シートを作成しました。カレンダーID等を入力してください。');
}

function getTimeOrDefault(val, def) {
  return (typeof val === 'string' && val.trim()) ? val.trim() : def;
}

/**
 * 予定自動ブロック本体
 */
function autoBlockEvening() {
  const ui = SpreadsheetApp.getUi();
  try {
    // 設定取得
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName('設定');
    if (!sheet) throw new Error('設定シートがありません。メニューから作成してください。');
    const values = sheet.getRange(2, 1, 5, 2).getValues();
    const config = {};
    values.forEach(function(row) { config[row[0]] = row[1]; });
    const calendarId = config['カレンダーID'];
    if (!calendarId || calendarId === 'ここにカレンダーIDを入力') throw new Error('カレンダーIDを設定してください。');
    const keywords = config['キーワード'].split(',').map(function(s){return s.trim();}).filter(String);
    const blockStart = getTimeOrDefault(config['ブロック開始'], '18:30');
    const blockEnd = getTimeOrDefault(config['ブロック終了'], '21:00');
    const days = parseInt(config['検索日数（未来）'] || '30', 10);

    // カレンダー取得
    const cal = CalendarApp.getCalendarById(calendarId);
    if (!cal) throw new Error('カレンダーが見つかりません。IDを確認してください。');

    // 未来N日分の予定取得
    const today = new Date();
    const until = new Date();
    until.setDate(today.getDate() + days);
    const events = cal.getEvents(today, until);

    // 日付ごとに該当イベントがあるか判定
    const blockDates = {};
    events.forEach(function(ev) {
      const title = ev.getTitle();
      if (keywords.some(function(k){return title.indexOf(k) !== -1;})) {
        const ymd = ev.getStartTime().getFullYear() + '-' + (ev.getStartTime().getMonth()+1) + '-' + ev.getStartTime().getDate();
        blockDates[ymd] = true;
      }
    });
    
    // ブロック実行
    let addCount = 0;
    Object.keys(blockDates).forEach(function(ymd) {
      const d = new Date(ymd);
      const s = blockStart.split(':');
      const e = blockEnd.split(':');
      const start = new Date(d); start.setHours(parseInt(s[0],10), parseInt(s[1],10), 0, 0);
      const end = new Date(d); end.setHours(parseInt(e[0],10), parseInt(e[1],10), 0, 0);
      // 既存の同名イベントがなければ追加
      const exists = cal.getEvents(start, end).some(function(ev){return ev.getTitle()==='予定あり'});
      if (!exists) {
        cal.createEvent('予定あり', start, end);
        addCount++;
      }
    });
    ui.alert('完了', addCount + '件の夜予定を自動ブロックしました。', ui.ButtonSet.OK);
  } catch(e) {
    ui.alert('エラー', e.message, ui.ButtonSet.OK);
  }
}

/**
 * 前日分の「予定あり」「予定を確保」イベントを削除
 */
function deletePreviousDayBlocks() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName('設定');
    if (!sheet) throw new Error('設定シートがありません。');
    const values = sheet.getRange(2, 1, 5, 2).getValues();
    const config = {};
    values.forEach(function(row) { config[row[0]] = row[1]; });
    const calendarId = config['カレンダーID'];
    if (!calendarId || calendarId === 'ここにカレンダーIDを入力') throw new Error('カレンダーIDを設定してください。');
    const cal = CalendarApp.getCalendarById(calendarId);
    if (!cal) throw new Error('カレンダーが見つかりません。IDを確認してください。');
    const today = new Date();
    const yesterday = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1);
    const start = new Date(yesterday); start.setHours(0,0,0,0);
    const end = new Date(yesterday); end.setHours(23,59,59,999);
    const events = cal.getEvents(start, end);
    let delCount = 0;
    events.forEach(function(ev) {
      if (ev.getTitle() === '予定あり' || ev.getTitle() === '予定を確保') {
        ev.deleteEvent();
        delCount++;
      }
    });
    Logger.log('前日分の予定あり/予定を確保イベントを' + delCount + '件削除しました');
  } catch(e) {
    Logger.log('ロック削除エラー: ' + e.message);
  }
}

/**
 * 3時間ごと自動ブロックトリガーON
 */
function setBlockTrigger() {
  deleteBlockTrigger();
  ScriptApp.newTrigger('autoBlockEvening').timeBased().everyHours(3).create();
  SpreadsheetApp.getUi().alert('3時間ごと自動ブロックトリガーを設定しました');
}

/**
 * 3時間ごと自動ブロックトリガーOFF
 */
function deleteBlockTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t){
    if (t.getHandlerFunction() === 'autoBlockEvening') ScriptApp.deleteTrigger(t);
  });
  SpreadsheetApp.getUi().alert('3時間ごと自動ブロックトリガーを解除しました');
}

/**
 * 朝5時ロック削除トリガーON
 */
function setDeleteTrigger() {
  deleteDeleteTrigger();
  ScriptApp.newTrigger('deletePreviousDayBlocks').timeBased().atHour(5).everyDays(1).create();
  SpreadsheetApp.getUi().alert('朝5時ロック削除トリガーを設定しました');
}

/**
 * 朝5時ロック削除トリガーOFF
 */
function deleteDeleteTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t){
    if (t.getHandlerFunction() === 'deletePreviousDayBlocks') ScriptApp.deleteTrigger(t);
  });
  SpreadsheetApp.getUi().alert('朝5時ロック削除トリガーを解除しました');
}

/**
 * 前日分の「予定あり」「予定を確保」イベントを手動で削除
 */
function manualDeletePreviousDayBlocks() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName('設定');
    if (!sheet) throw new Error('設定シートがありません。');
    const values = sheet.getRange(2, 1, 5, 2).getValues();
    const config = {};
    values.forEach(function(row) { config[row[0]] = row[1]; });
    const calendarId = config['カレンダーID'];
    if (!calendarId || calendarId === 'ここにカレンダーIDを入力') throw new Error('カレンダーIDを設定してください。');
    const cal = CalendarApp.getCalendarById(calendarId);
    if (!cal) throw new Error('カレンダーが見つかりません。IDを確認してください。');
    const today = new Date();
    const yesterday = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1);
    const start = new Date(yesterday); start.setHours(0,0,0,0);
    const end = new Date(yesterday); end.setHours(23,59,59,999);
    const events = cal.getEvents(start, end);
    let delCount = 0;
    events.forEach(function(ev) {
      if (ev.getTitle() === '予定あり' || ev.getTitle() === '予定を確保') {
        ev.deleteEvent();
        delCount++;
      }
    });
    SpreadsheetApp.getUi().alert('前日分の予定あり/予定を確保イベントを' + delCount + '件削除しました');
  } catch(e) {
    SpreadsheetApp.getUi().alert('エラー: ' + e.message);
  }
}
