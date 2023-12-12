function DeleteStoryScheduleMain(){
  exclusiveMain(DeleteStorySchedule)
}
{
  //-----------------
  //定数定義（進行表）
  //-----------------
  //削除後詰めのチェックボックスの場所
  const TSUME_DEL_ON_ROW = 2;
  const TSUME_DEL_ON_COLUMN = 17;

  const STORY_TITLE_ROW = 1;
  const STORY_TITLE_COLUMN = 1;
  
  function DeleteStorySchedule(){
    const label = 'DeleteStorySchedule';
    console.time(label);

    // アクティブなスプレッドシートを取得
    const mainSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = mainSpreadsheet.getActiveSheet(); //進行表
    const databaseSheet = mainSpreadsheet.getSheetByName(DATA_BASE_SHEET_NAME);
    if (!databaseSheet) {
      throw new Error('「' + DATA_BASE_SHEET_NAME + '」シートが見つかりません');
    }
    const schedulesheet = mainSpreadsheet.getSheetByName(SS_SCHEDULE_SHEET_NAME);
    if (!schedulesheet) {
      throw new Error('「' + SS_SCHEDULE_SHEET_NAME + '」シートが見つかりません');
    }

    
    //グローバル変数のスケジュール表とデータベースシートを作成
    getScheduleSheetInfoC(schedulesheet,databaseSheet)
    
    //指定されたsourceSheetのA1セルの値とdatabaseSheetのデータを比較して、一致する範囲を返す
    let foundRanges = findRangesMatchingSourceValue(sourceSheet,databaseSheet)

    if(isRangesFixedC(dataBaseSheetDataValues, foundRanges)){
      // ダイアログを出して続行するかユーザーに確認
      if(!confirmFixedCellExecutionC()){
        return;//後の処理を実行しない。
      }
    }
    
    // スケジュールが何日まで記載があるか確認する。
    const maxDays = getLastFilledColumnInCalender(schedulesheet,SS_CALENDERDATE_COLUMN_INDEX);  //カレンダーの最大サイズを確認する。
    
    //削除後詰めONかどうかを確認する
    let tsumeDelOn = tsumeDelCheck(sourceSheet)
    //スケジュール管理仕様用にoffsetする
    processAndDeleteCellsInRanges(foundRanges,schedulesheet,tsumeDelOn,maxDays)
  
    // 更新したスケジュール表情報で画面更新
    updateScheduleSheetWithDataValuesC();

    // スケジュール表をActiveにする
    schedulesheet.activate();

    console.timeEnd(label)
  }

  //削除後詰めがONかどうか確認する
  function tsumeDelCheck(sourceSheet){
    // 指定した行と列からセルの値を取得する
    let tsumeDelOn = sourceSheet.getRange(TSUME_DEL_ON_ROW, TSUME_DEL_ON_COLUMN).getValue();
    return tsumeDelOn
  }

  // 指定されたsourceSheetのA1セルの値とdatabaseSheetのデータを比較して、一致する範囲を返す
  function findRangesMatchingSourceValue(sourceSheet,databaseSheet) {
    // 対象となる値を進行表の機能検討用 2シートから取得
    const targetValue = sourceSheet.getRange(STORY_TITLE_ROW, STORY_TITLE_COLUMN).getValue();
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
  
  function processAndDeleteCellsInRanges(foundRanges, schedulesheet, tsumeDelOn, maxDays) {
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
      displaySceneNameC(rowIndex,maxDays);
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




