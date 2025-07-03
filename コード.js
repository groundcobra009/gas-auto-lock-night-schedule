/**
 * 夜の予定自動ブロックスクリプト
 * バージョン: 1.1.0
 * 最終更新: 20250703
 * https://docs.google.com/spreadsheets/d/1XlpKdYyes_iUmdODLDeRnTBE4MnNo3eiscw24zdNdEg/edit?gid=475079626#gid=475079626
 * 
 * 機能:
 * - 指定したキーワードを含む予定がある日の夜の時間帯を自動でブロック
 * - 20時半以降に予定がある日の夜の時間帯を自動でブロック
 * - 前日分のブロック予定を自動削除
 * - 手動でのブロック設定と削除
 */

/**
 * スプレッドシートを開いたときにメニューを追加
 * メニュー項目:
 * - 設定シート作成
 * - 予定自動ブロック実行
 * - 3時間ごと自動ブロックON/OFF
 * - 朝5時ロック削除ON/OFF
 * - 前日の予定を手動削除
 * - トリガー初期設定
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
    .addSeparator()
    .addItem('トリガー初期設定', 'initializeTriggers')
    .addToUi();
}

/**
 * 設定シートを自動生成する
 * 設定項目:
 * - カレンダーID: 対象のGoogleカレンダーID
 * - キーワード: ブロック対象となる予定のキーワード（カンマ区切り）
 * - ブロック開始: ブロック開始時間（HH:MM形式）
 * - ブロック終了: ブロック終了時間（HH:MM形式）
 * - 検索日数（未来）: 未来何日分の予定を検索するか
 * - 夜予定ブロック開始: 20:30以降の予定がある場合のブロック開始時間
 * - 夜予定ブロック終了: 20:30以降の予定がある場合のブロック終了時間
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
    ['検索日数（未来）', '30'],
    ['夜予定ブロック開始', '18:30'],
    ['夜予定ブロック終了', '20:00']
  ];
  sheet.getRange(2, 1, settings.length, 2).setValues(settings);
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 300);
  sheet.getRange('A1:B1').setFontWeight('bold').setBackground('#f3f3f3');
  
  // 説明を追加
  sheet.getRange('A9').setValue('説明');
  sheet.getRange('B9').setValue('キーワードに該当する予定または20:30以降に予定がある日の夜をブロック');
  sheet.getRange('A9:B9').setFontWeight('bold').setBackground('#e6f3ff');
  
  SpreadsheetApp.getUi().alert('設定シートを作成しました。カレンダーID等を入力してください。');
}

/**
 * 時間文字列のデフォルト値処理
 * @param {string} val - 入力値
 * @param {string} def - デフォルト値
 * @return {string} 処理後の時間文字列
 */
function getTimeOrDefault(val, def) {
  return (typeof val === 'string' && val.trim()) ? val.trim() : def;
}

/**
 * 予定自動ブロック本体
 * 処理の流れ:
 * 1. 設定の取得
 * 2. カレンダーから予定を取得
 * 3. キーワードに一致する予定がある日を特定
 * 4. 20時半以降に予定がある日を特定
 * 5. 該当日の夜の時間帯をブロック
 */
function autoBlockEvening() {
  try {
    // 設定取得
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName('設定');
    if (!sheet) throw new Error('設定シートがありません。メニューから作成してください。');
    const values = sheet.getRange(2, 1, 7, 2).getValues();
    const config = {};
    values.forEach(function(row) { config[row[0]] = row[1]; });
    const calendarId = config['カレンダーID'];
    if (!calendarId || calendarId === 'ここにカレンダーIDを入力') throw new Error('カレンダーIDを設定してください。');
    const keywords = config['キーワード'].split(',').map(function(s){return s.trim();}).filter(String);
    const blockStart = getTimeOrDefault(config['ブロック開始'], '18:30');
    const blockEnd = getTimeOrDefault(config['ブロック終了'], '21:00');
    const nightBlockStart = getTimeOrDefault(config['夜予定ブロック開始'], '18:30');
    const nightBlockEnd = getTimeOrDefault(config['夜予定ブロック終了'], '20:00');
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
    const keywordBlockDates = {};  // キーワードベースのブロック
    const nightBlockDates = {};    // 20:30以降の予定ベースのブロック
    events.forEach(function(ev) {
      const title = ev.getTitle();
      const startTime = ev.getStartTime();
      const eventDate = new Date(startTime.getFullYear(), startTime.getMonth(), startTime.getDate());
      const ymd = eventDate.getFullYear() + '-' + (eventDate.getMonth()+1) + '-' + eventDate.getDate();
      
      // キーワードに一致する予定をチェック
      if (keywords.some(function(k){return title.indexOf(k) !== -1;})) {
        keywordBlockDates[ymd] = true;
      }
      
      // 20時半以降の予定をチェック
      const eventHour = startTime.getHours();
      const eventMinute = startTime.getMinutes();
      if (eventHour > 20 || (eventHour === 20 && eventMinute >= 30)) {
        nightBlockDates[ymd] = true;
      }
    });
    
    // ブロック実行
    let addCount = 0;
    
    // キーワードベースのブロック
    Object.keys(keywordBlockDates).forEach(function(ymd) {
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
    
    // 夜予定ベースのブロック（キーワードブロックと重複しない日のみ）
    Object.keys(nightBlockDates).forEach(function(ymd) {
      if (!keywordBlockDates[ymd]) {  // キーワードベースのブロックと重複しない場合のみ
        const d = new Date(ymd);
        const s = nightBlockStart.split(':');
        const e = nightBlockEnd.split(':');
        const start = new Date(d); start.setHours(parseInt(s[0],10), parseInt(s[1],10), 0, 0);
        const end = new Date(d); end.setHours(parseInt(e[0],10), parseInt(e[1],10), 0, 0);
        // 既存の同名イベントがなければ追加
        const exists = cal.getEvents(start, end).some(function(ev){return ev.getTitle()==='予定あり'});
        if (!exists) {
          cal.createEvent('予定あり', start, end);
          addCount++;
        }
      }
    });

    // トリガーから実行された場合はログに記録、手動実行の場合はダイアログを表示
    if (ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'autoBlockEvening')) {
      Logger.log(addCount + '件の夜予定を自動ブロックしました。');
    } else {
      SpreadsheetApp.getUi().alert('完了', addCount + '件の夜予定を自動ブロックしました。', SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch(e) {
    // トリガーから実行された場合はログに記録、手動実行の場合はダイアログを表示
    if (ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'autoBlockEvening')) {
      Logger.log('エラー: ' + e.message);
    } else {
      SpreadsheetApp.getUi().alert('エラー', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}

/**
 * 前日分の「予定あり」「予定を確保」イベントを削除
 * 処理の流れ:
 * 1. 設定の取得
 * 2. 前日分のイベントを取得
 * 3. 「予定あり」「予定を確保」のイベントを削除
 */
function deletePreviousDayBlocks() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName('設定');
    if (!sheet) throw new Error('設定シートがありません。');
    const values = sheet.getRange(2, 1, 7, 2).getValues();
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
 * 既存のトリガーを削除してから新規作成
 */
function setBlockTrigger() {
  deleteBlockTrigger();
  ScriptApp.newTrigger('autoBlockEvening').timeBased().everyHours(3).create();
  SpreadsheetApp.getUi().alert('3時間ごと自動ブロックトリガーを設定しました');
}

/**
 * 3時間ごと自動ブロックトリガーOFF
 * 該当するトリガーをすべて削除
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
 * 既存のトリガーを削除してから新規作成
 */
function setDeleteTrigger() {
  deleteDeleteTrigger();
  ScriptApp.newTrigger('deletePreviousDayBlocks').timeBased().atHour(5).everyDays(1).create();
  SpreadsheetApp.getUi().alert('朝5時ロック削除トリガーを設定しました');
}

/**
 * 朝5時ロック削除トリガーOFF
 * 該当するトリガーをすべて削除
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
 * 処理の流れ:
 * 1. 設定の取得
 * 2. 前日分のイベントを取得
 * 3. 「予定あり」「予定を確保」のイベントを削除
 * 4. 結果をダイアログで表示
 */
function manualDeletePreviousDayBlocks() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName('設定');
    if (!sheet) throw new Error('設定シートがありません。');
    const values = sheet.getRange(2, 1, 7, 2).getValues();
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

/**
 * トリガーの初期設定
 * 1. 既存のトリガーをすべて削除
 * 2. 3時間ごとの自動ブロックトリガーを設定
 * 3. 朝5時のロック削除トリガーを設定
 */
function initializeTriggers() {
  const ui = SpreadsheetApp.getUi();
  try {
    // 既存のトリガーをすべて削除
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(function(t) {
      ScriptApp.deleteTrigger(t);
    });
    
    // 新しいトリガーを設定
    ScriptApp.newTrigger('autoBlockEvening').timeBased().everyHours(3).create();
    ScriptApp.newTrigger('deletePreviousDayBlocks').timeBased().atHour(5).everyDays(1).create();
    
    ui.alert('トリガー初期設定完了', '以下のトリガーを設定しました：\n・3時間ごとの自動ブロック\n・朝5時のロック削除', ui.ButtonSet.OK);
  } catch(e) {
    ui.alert('エラー', 'トリガー初期設定中にエラーが発生しました：\n' + e.message, ui.ButtonSet.OK);
  }
}
