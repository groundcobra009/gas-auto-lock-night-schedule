/**
 * コンテナバインドされたスプレッドシートのURL取得とログ出力
 * バージョン: 1.0.0
 * 作成日: 2025-01-17
 * 
 * 機能:
 * - 現在のGASプロジェクトにバインドされたスプレッドシートのURLを取得
 * - URLをログに出力
 * - URLをコンソールに表示（手動実行時）
 */

/**
 * コンテナバインドされたスプレッドシートのURLを取得してログに出力する
 * @return {string} スプレッドシートのURL
 */
function getBoundSpreadsheetUrl() {
  try {
    // コンテナバインドされたスプレッドシートを取得
    const spreadsheet = SpreadsheetApp.getActive();
    
    if (!spreadsheet) {
      const errorMsg = 'このスクリプトはスプレッドシートにバインドされていません。';
      Logger.log('エラー: ' + errorMsg);
      console.log('エラー: ' + errorMsg);
      return null;
    }
    
    // スプレッドシートのURLを取得
    const url = spreadsheet.getUrl();
    const name = spreadsheet.getName();
    const id = spreadsheet.getId();
    
    // ログに出力
    Logger.log('=== コンテナバインドされたスプレッドシート情報 ===');
    Logger.log('スプレッドシート名: ' + name);
    Logger.log('スプレッドシートID: ' + id);
    Logger.log('スプレッドシートURL: ' + url);
    Logger.log('=======================================');
    
    // コンソールにも出力（手動実行時に確認しやすくするため）
    console.log('=== コンテナバインドされたスプレッドシート情報 ===');
    console.log('スプレッドシート名: ' + name);
    console.log('スプレッドシートID: ' + id);
    console.log('スプレッドシートURL: ' + url);
    console.log('=======================================');
    
    return url;
    
  } catch (error) {
    const errorMsg = 'スプレッドシートURL取得中にエラーが発生しました: ' + error.toString();
    Logger.log('エラー: ' + errorMsg);
    console.log('エラー: ' + errorMsg);
    return null;
  }
}

/**
 * スプレッドシートの詳細情報を取得してログに出力する
 * @return {Object} スプレッドシートの詳細情報
 */
function getBoundSpreadsheetDetails() {
  try {
    const spreadsheet = SpreadsheetApp.getActive();
    
    if (!spreadsheet) {
      Logger.log('エラー: このスクリプトはスプレッドシートにバインドされていません。');
      return null;
    }
    
    // 詳細情報を取得
    const details = {
      name: spreadsheet.getName(),
      id: spreadsheet.getId(),
      url: spreadsheet.getUrl(),
      locale: spreadsheet.getSpreadsheetLocale(),
      timeZone: spreadsheet.getSpreadsheetTimeZone(),
      sheetCount: spreadsheet.getSheets().length,
      sheetNames: spreadsheet.getSheets().map(sheet => sheet.getName())
    };
    
    // ログに出力
    Logger.log('=== スプレッドシート詳細情報 ===');
    Logger.log('名前: ' + details.name);
    Logger.log('ID: ' + details.id);
    Logger.log('URL: ' + details.url);
    Logger.log('ロケール: ' + details.locale);
    Logger.log('タイムゾーン: ' + details.timeZone);
    Logger.log('シート数: ' + details.sheetCount);
    Logger.log('シート名: ' + details.sheetNames.join(', '));
    Logger.log('============================');
    
    return details;
    
  } catch (error) {
    Logger.log('エラー: ' + error.toString());
    return null;
  }
}

/**
 * 現在のGASプロジェクト情報も含めてログに出力する
 */
function getProjectAndSpreadsheetInfo() {
  try {
    // GASプロジェクト情報
    const scriptId = ScriptApp.getScriptId();
    
    Logger.log('=== GASプロジェクト情報 ===');
    Logger.log('スクリプトID: ' + scriptId);
    Logger.log('スクリプトURL: https://script.google.com/d/' + scriptId + '/edit');
    
    // スプレッドシート情報
    const spreadsheet = SpreadsheetApp.getActive();
    if (spreadsheet) {
      Logger.log('=== バインドされたスプレッドシート ===');
      Logger.log('スプレッドシート名: ' + spreadsheet.getName());
      Logger.log('スプレッドシートURL: ' + spreadsheet.getUrl());
    } else {
      Logger.log('このスクリプトはスプレッドシートにバインドされていません。');
    }
    
    Logger.log('=========================');
    
  } catch (error) {
    Logger.log('エラー: ' + error.toString());
  }
} 