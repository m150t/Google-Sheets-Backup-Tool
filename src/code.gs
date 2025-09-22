/********** 本体 **********/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📦 Sheet Backup')
    .addItem('① 権限付与を実行', 'authorizeApp_')
    .addItem('② プロパティを設定する', 'openPropertiesHelp_')
    .addItem('③ 日次トリガー作成（毎日23:00）', 'installDailyTrigger_')
    .addSeparator()
    .addItem('④ 今すぐバックアップ（コピー＋PDF）', 'runBackupNow')
    .addToUi();
}

function authorizeApp_() {
  // Spreadsheet スコープを要求（どの環境でもOK）
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const _sheetNames = ss.getSheets().map(s => s.getName()); // 触るだけ

  // Drive スコープ（drive.file）を必要に応じて要求
  // BACKUP_FOLDER_ID が設定済みなら、そのフォルダだけ軽く触っておく
  const id = PropertiesService.getScriptProperties().getProperty('BACKUP_FOLDER_ID');
  if (id) {
    try { DriveApp.getFolderById(id).getName(); } catch(_) { /* 無視：後で詳しく検証する */ }
  }

  SpreadsheetApp.getUi().alert('権限付与が完了しました。バックアップを実行できます。');
}

function installDailyTrigger_() {
  try {
    upsertDailyTrigger('runBackupNow', 23);
    info('日次トリガーを作成しました（毎日23:00）。');
  } catch (err) {
    fail(err, 'トリガー作成でエラー：\n');
  }
}

function runBackupNow() {
  try {
    withScriptLock(() => {
      const folderId = getProp('BACKUP_FOLDER_ID', { required: true });
      const folder = getFolderByIdSafe_(folderId);

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const nameBase = sanitizeName_(ss.getName());
      const stamp = jstStamp('yyyyMMdd_HHmmss');

      // 1) シートのコピー
      const copied = DriveApp.getFileById(ss.getId())
        .makeCopy(`${nameBase}_backup_${stamp}`, folder);

      // 2) PDF
      const pdfBlob = exportSpreadsheetToPdf_(ss, { title: `${nameBase}_${stamp}` });
      const pdfFile = folder.createFile(pdfBlob);

      info(`バックアップ完了！\nコピー: ${copied.getUrl()}\nPDF: ${pdfFile.getUrl()}`);
    });
  } catch (err) {
    fail(err, 'バックアップでエラーが発生しました：\n');
  }
}

function openPropertiesHelp_() {
  const html = HtmlService.createHtmlOutput(renderPropertiesHelpHtml_())
    .setTitle('プロパティ設定の手順');
  SpreadsheetApp.getUi().showSidebar(html);
}

function renderPropertiesHelpHtml_() {
  return `
    <h2>バックアップ先フォルダIDの設定（必須）</h2>
    <ol>
      <li>Googleドライブで保存先フォルダを開き、URL の <code>folders/</code> の後ろが <b>フォルダID</b></li>
      <li>スクリプトエディタ → プロジェクト設定（歯車アイコン） → スクリプトのプロパティ</li>
      <li><b>キー:</b> <code>BACKUP_FOLDER_ID</code>, <b>値:</b> フォルダID</li>
    </ol>
    <p>設定後、「④ 今すぐバックアップ」で動作します。</p>
  `;
}
