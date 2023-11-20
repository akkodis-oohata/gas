/*
空白セル追加処理
*/
function addBlankCellsMain() {
  exclusiveMain(addBlankCells);
}


{

  function addBlankCells() {
    console.log('addBlankCells in')
    const label = 'addCells'
    console.time(label)
    // スプレッドシートの読み込み
    let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let scheduleSheet = spreadsheet.getActiveSheet();

    // 選択したセルのrange取得
    var selection = scheduleSheet.getSelection();
    var ranges = selection.getActiveRangeList().getRanges();


    // 選択範囲のチェックと取得
    let range = getSelectedRange(ranges, scheduleSheet);
    if (range == undefined) {
      return;
    }

    // データベースシートの読み込み
    let dataBaseSheet = spreadsheet.getSheetByName(DATA_BASE_SHEET_NAME);

    //データスペース（作品話数ベースデータ）へ反映を行う。
    //updateDataSpaceMain(spreadsheet)

    // データベースシートの値を読み込み
    dataValues = dataBaseSheet.getDataRange().getValues();
    // 固定セルが含まれているか確認を行う。
    if(isRangeFixedC(dataValues, range) ){
      // ダイアログを出して続行するかユーザーに確認
      if(!confirmFixedCellExecutionC()){
        return;//後の処理を実行しない。
      }
    }
    
    // スケジュール表情報取得(データスペース反映後)
    getScheduleSheetPersonInfoC(scheduleSheet, dataBaseSheet, range.getRow());

    // 空白セルの挿入
    addBlankCellsC(range, false);

    const maxDays = getLastFilledColumnInCalender(scheduleSheet,SS_CALENDERDATE_COLUMN_INDEX);  //カレンダーの最大サイズを確認する。
    // 先頭の名前更新
    // let rowIndex = range.getRow() - 1;
    displaySceneNameC(0,maxDays);

    // 更新したスケジュール表情報で画面更新
    updateScheduleSheetPresonWithDataValuesC();
    console.timeEnd(label);
    console.log("addBlankCells out");
  }
}
