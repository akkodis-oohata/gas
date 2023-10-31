// 目的の背景色(ここでは黄色)
const TARGET_COLOR = "#ff9900";

// シート名  
const SHEET_NAME = "Sheet1";

function main(){
  const label = 'DeleteStorySchedule'
  console.time(label)
  // アクティブなスプレッドシートを取得
  const mainSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = mainSpreadsheet.getSheetByName('進行表の機能検討用 2');
  const databaseSheet = mainSpreadsheet.getSheetByName('データベース');
  const schedulesheet = mainSpreadsheet.getSheetByName('スケジュール管理仕様');

  getScheduleSheetInfoC(schedulesheet,databaseSheet)



  var foundRanges = findMatchingRanges(sourceSheet,databaseSheet)
  // 上記のfindCellsWithColorAndLog()関数でfoundRangesを取得した後に、logFoundRanges()関数を呼び出す
  logFoundRanges(foundRanges);
  
  //スケジュール管理仕様用にoffsetする
  FoundRangesToRange(foundRanges,schedulesheet)

  // 更新したスケジュール表情報で画面更新
  updateScheduleSheetWithDataValuesC();
  

  

  console.timeEnd(label)
}

function findMatchingRanges(sourceSheet,databaseSheet) {
  console.log("findMatchingRanges()");
  // 対象となる値を進行表の機能検討用 2シートから取得
  const targetValue = sourceSheet.getRange('A1').getValue();
  // データベースシートのデータを2次元配列として取得
  const dabaseData = databaseSheet.getDataRange().getValues();

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


function FoundRangesToRange(foundRanges, schedulesheet) {
  // foundRanges内の各Rangeオブジェクトについて処理を行う
  foundRanges.forEach((range, index) => {
    // 元のrangeのアドレスを取得
    const rangeAddress = range.getA1Notation();
    
    // schedulesheetで同じアドレスのRangeを取得
    const targetScheduleAddress = schedulesheet.getRange(rangeAddress).offset(DATA_BASE_FORMAT_OFFSET, 0).getA1Notation();
    const targetScheduleRange = schedulesheet.getRange(targetScheduleAddress);

    // scheduleRangeの情報を取得
    const sheetName = targetScheduleRange.getSheet().getName();
    const rowNumber = targetScheduleRange.getRow();
    const colNumber = targetScheduleRange.getColumn();
    const numRows = targetScheduleRange.getNumRows();
    const numCols = targetScheduleRange.getNumColumns();
    
    // 座標と範囲のサイズをログに出力
    Logger.log(`Index: ${index}, Sheet: ${sheetName}, Row: ${rowNumber}-${rowNumber + numRows - 1}, Column: ${colNumber}-${colNumber + numCols - 1}`);
    deleteCellsC(targetScheduleRange)
  });
}




