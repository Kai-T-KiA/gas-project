// スプレッドシートが開かれた時に自動実行
function onOpen() {
　// 現在開いているスプレッドシートのユーザーインターフェース（UI）を操作するためのオブジェクトを取得
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('PDF管理')
      .addItem('PDFをスキャン', 'scan_papers_folder')
      .addToUi();
}

// スプレッドシートのIDを設定
const SPREADSHEET_ID = '使用したいスプレッドシートのID';
// 論文フォルダのIDを設定
const PAPERS_FOLDER_ID = '使用したいフォルダのID';

// PDFをフォルダからスキャン
function scan_papers_folder() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
  const folder = DriveApp.getFolderById(PAPERS_FOLDER_ID);
  
  // ヘッダーを設定（既存のヘッダーがある場合は上書き）
  const headers = ['FileName', 'FileID', 'Title', 'LastUpdated', 'Link'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // 既存のデータを取得
  const existing_data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const file_ids = existing_data.map(row => row[1])
  
  const files = folder.getFilesByType(MimeType.PDF);
  let new_data = [];
  
  while (files.hasNext()) {
    const file = files.next();
    const file_id = file.getId();
    const file_name = file.getName();
    const last_updated = file.getLastUpdated();
    const file_link = file.getUrl();
    
    const row_data = [
      file_name,
      file_id,
      file_name.replace('.pdf', ''), // タイトルにするためファイル名を削除
      last_updated,
      '=HYPERLINK("' + file_link + '","開く")' // リンクを追加
    ];
    
    const existing_index = file_ids.indexOf(file_id);
    if (existing_index === -1) {
      // 新しいファイルの場合、新しいデータに追加
      new_data.push(row_data);
    } else {
      // 既存のファイルの場合、既存のデータを更新
      existing_data[existing_index] = row_data;
    }
  }
  
  // 既存のデータと新しいデータを結合
  const all_data = existing_data.concat(new_data);
  
  // データをシートに記入
  if (all_data.length > 0) {
    sheet.getRange(2, 1, all_data.length, headers.length).setValues(all_data);
  }
  
  // リンク列のセルの書式設定
  if (all_data.length > 0) {
    sheet.getRange(2, 5, all_data.length, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }
  
  SpreadsheetApp.getUi().alert('PDFのスキャンが完了しました。');
}
