// 排他制御
function exclusiveMain(callback) {
  let lock = LockService.getDocumentLock();
  lock.tryLock(0);
  if (lock.hasLock()) {
    // スプレッドシートの読み込み
    // let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    // シート読み込み
    // const sheets = spreadsheet.getSheets();

    //protectSheet(sheets); // シート保護

    // 処理実行
    callback();

    //unProtectSheet(sheets); // シート保護解除
    lock.releaseLock(); //ロックを解除
  } else {
    Browser.msgBox("他の処理が実行中です");
  }
}

// 排他チェック
function exclusiveCheck(callback) {
  let lock = LockService.getDocumentLock();
  lock.tryLock(0);
  if (lock.hasLock()) {
    lock.releaseLock(); //ロックを解除
    // 処理実行
    callback();
  } else {
    Browser.msgBox("他の処理が実行中です");
  }
}
{
  // シート保護
  function protectSheet(sheets) {
    sheets.forEach((sheet) => {
      //読み込んだシートに保護を設定し、Protectionオブジェクトを変数に格納
      let protections = sheet.protect();
      //保護したシートで編集可能なユーザーを取得
      let userList = protections.getEditors();
      //オーナーのみ編集可能にするため、編集ユーザーをすべて削除
      //オーナーの編集権限は削除できないため、オーナーのみ編集可能に
      //protections.removeEditors(userList);
      //保護内容の説明文章を設定
      protections.setDescription("他の処理が実行中です");
      //保護を入力不可ではなく、入力時に警告を表示
      protections.setWarningOnly(true);
    });
  }
  // シート保護解除
  function unProtectSheet(sheets) {
    sheets.forEach((sheet) => {
      const protection = sheet.getProtections(
        SpreadsheetApp.ProtectionType.SHEET
      )[0];
      if (protection && protection.canEdit()) {
        protection.remove();
      }
    });
  }
}
