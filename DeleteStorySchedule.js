function DeleteStoryScheduleMain(){
  exclusiveMain(DeleteStorySchedule)
}
{
  //-----------------
  //定数定義（進行表）
  //-----------------
  //削除後詰めのチェックボックスの場所
  const TSUME_DEL_ON_ROW_INDEX = 1;
  const TSUME_DEL_ON_COLUMN_INDEX = 16;
  const DATA_BASE_SHEET_NAME = "データベース"
  const SCHEDULE_SHEET_NAME = "スケジュール管理仕様"
  
  function DeleteStorySchedule(){
    const label = 'DeleteStorySchedule';
    console.time(label);

    // アクティブなスプレッドシートを取得
    const mainSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = mainSpreadsheet.getActiveSheet(); //進行表
    const databaseSheet = mainSpreadsheet.getSheetByName(DATA_BASE_SHEET_NAME);
    const schedulesheet = mainSpreadsheet.getSheetByName(SCHEDULE_SHEET_NAME);
    
    //グローバル変数のスケジュール表とデータベースシートを作成
    getScheduleSheetInfoC(schedulesheet,databaseSheet)

    //削除後詰めONかどうかを確認する
    let tsumeDelOn = tsumeDelCheck(sourceSheet)
    
    //指定されたsourceSheetのA1セルの値とdatabaseSheetのデータを比較して、一致する範囲を返す
    let foundRanges = findRangesMatchingSourceValue(sourceSheet,databaseSheet)
    
    // 上記のfindCellsWithColorAndLog()関数でfoundRangesを取得した後に、logFoundRanges()関数を呼び出す
    //logFoundRanges(foundRanges);
    
    //スケジュール管理仕様用にoffsetする
    processAndDeleteCellsInRanges(foundRanges,schedulesheet,tsumeDelOn)
  
    // 更新したスケジュール表情報で画面更新
    updateScheduleSheetWithDataValuesC();

    // スケジュール表をActiveにする
    schedulesheet.activate();

    console.timeEnd(label)
  }

  //削除後詰めがONかどうか確認する
  function tsumeDelCheck(sourceSheet){
    //一か所であれば、getValuesするよりもいいかと思い下記に変更。
    let tsumeDelOn = sourceSheet.getRange("Q2").getValue();
    return tsumeDelOn
  }

  // 指定されたsourceSheetのA1セルの値とdatabaseSheetのデータを比較して、一致する範囲を返す
  function findRangesMatchingSourceValue(sourceSheet,databaseSheet) {
    // 対象となる値を進行表の機能検討用 2シートから取得
    const targetValue = sourceSheet.getRange('A1').getValue();
    // データベースシートのデータを2次元配列として取得
    const dabaseData = dataBaseSheetDataValues;
  
    let currentRangeStart = null;  // 現在の範囲の開始位置を追跡
    const foundRanges = [];  // 一致する範囲を格納する配列
  
    // 2次元配列の各要素を走査
    for (let row = 0; row < dabaseData.length; row++) {
      for (let col = 0; col < dabaseData[row].length; col++) {
        let data = null;
        if (typeof dabaseData[row][col] === 'string') {
          data = dabaseData[row][col].split(',', 1)[0];
        } else {
          continue;
        }
  
        // 値が一致する場合、範囲の開始位置をセット
        if (data === targetValue) {
          if (currentRangeStart === null) {
            currentRangeStart = { row: row, col: col };
          }
        } else {
          // 値が一致しない場合、範囲を配列に追加
          if (currentRangeStart) {
            let numberOfRows = 1;  // 行は常に1
            let numberOfCols = col - currentRangeStart.col;
            if (numberOfCols <= 0) {
              numberOfCols = 1;
            }
            foundRanges.push(databaseSheet.getRange(currentRangeStart.row + 1, currentRangeStart.col + 1, numberOfRows, numberOfCols));
            currentRangeStart = null;
          }
        }
        // 列の最後で一致する範囲が残っている場合、配列に追加
        if (currentRangeStart && col === dabaseData[row].length - 1) {
          let numberOfCols = col - currentRangeStart.col + 1;
          foundRanges.push(databaseSheet.getRange(currentRangeStart.row + 1, currentRangeStart.col + 1, 1, numberOfCols));
          currentRangeStart = null;
        }
      }
    }
  
    // データの最後の行で一致する範囲が残っている場合、配列に追加
    if (currentRangeStart) {
      let numberOfCols = dabaseData[0].length - currentRangeStart.col;
      foundRanges.push(databaseSheet.getRange(currentRangeStart.row + 1, currentRangeStart.col + 1, 1, numberOfCols));
    }
  
    return foundRanges;
  }

  function logFoundRanges(foundRanges) {
    // foundRanges内の各Rangeオブジェクトについて処理を行う
    foundRanges.forEach((range, index) => {
      const sheetName = range.getSheet().getName();
      const rowNumber = range.getRow();
      const colNumber = range.getColumn();
      const numRows = range.getNumRows();
      const numCols = range.getNumColumns();
      
      // 座標と範囲のサイズをログに出力
      Logger.log(`Index: ${index}, Sheet: ${sheetName}, Row: ${rowNumber}-${rowNumber+numRows-1}, Column: ${colNumber}-${colNumber+numCols-1}`);
    });
  }
  
  function processAndDeleteCellsInRanges(foundRanges, schedulesheet, tsumeDelOn) {
    // 逆側からrangesを削除して行くことで、削除後詰めを行えるようにする。
    let reversedFoundRanges = [...foundRanges].reverse();
    
    // foundRanges内の各Rangeオブジェクトについて処理を行う
    reversedFoundRanges.forEach((range, index) => {
      // 元のrangeのアドレスを取得
      const rangeAddress = range.getA1Notation();
      // schedulesheetで同じアドレスのRangeを取得
      const targetScheduleRange = schedulesheet.getRange(rangeAddress);
      
      // logRangeInfo(targetScheduleRange)
      
      if(tsumeDelOn){
        // セル削除（削除後詰めあり）
        deleteCellsWithMove(targetScheduleRange)
      }else{
        // セル削除（削除後詰めなし）
        deleteCellsC(targetScheduleRange)
      }
      // 先頭の名前更新
      let rowIndex = targetScheduleRange.getRow() - 1;
      displaySceneNameC(rowIndex);
    });
  }


  function logRangeInfo(range, index) {
    const sheetName = range.getSheet().getName();
    const rowNumber = range.getRow();
    const colNumber = range.getColumn();
    const numRows = range.getNumRows();
    const numCols = range.getNumColumns();
  
    // indexが提供されていればそれもログに出力する
    if (typeof index !== 'undefined') {
      Logger.log(`Index: ${index}, Sheet: ${sheetName}, Row: ${rowNumber}-${rowNumber+numRows-1}, Column: ${colNumber}-${colNumber+numCols-1}`);
    } else {
      Logger.log(`Sheet: ${sheetName}, Row: ${rowNumber}-${rowNumber+numRows-1}, Column: ${colNumber}-${colNumber+numCols-1}`);
    }
  }
}




