/*
削除後詰め、削除処理
*/
function deleteCellsMain() {
  exclusiveMain(deleteCells);
}

{
  // 削除後詰めON の場所
  const TSUMEDELON_ROW = 1;
  const TSUMEDELON_COLUMN = 6;

  function deleteCells() {
    console.log("deleteCells in");
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

    const label = "deleteCells";
    console.time(label);

    // スケジュール表情報取得(データスペース反映後)
    getScheduleSheetPersonInfoC(scheduleSheet, dataBaseSheet, range.getRow());

    // 削除後詰めON
    let tsumeDelOn = scheduleSheet
      .getRange(TSUMEDELON_ROW, TSUMEDELON_COLUMN)
      .getValue();
    console.log("tsumeDelOn:" + tsumeDelOn);

    console.log("Row,LastRow: " + range.getRow() + "," + range.getLastRow());
    console.log(
      "Column,LastColumn: " + range.getColumn() + "," + range.getLastColumn()
    );

    if (tsumeDelOn) {
      //セル削除（削除後詰めあり）
      deleteCellsWithMove(range, false);
    } else {
      //セル削除（削除後詰めなし）
      deleteCellsC(range, false);
    }

    // 先頭の名前更新
    // let rowIndex = range.getRow() - 1;
    displaySceneNameC(0);

    // 更新したスケジュール表情報で画面更新
    updateScheduleSheetPresonWithDataValuesC();

    console.timeEnd(label);
    console.log("deleteCells out");
  }
}
