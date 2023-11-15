function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuItems = [
    {name: '固定セルON', functionName: 'addFixedCellMarkerMain'},
    {name: '固定セルOFF', functionName: 'removeFixedCellMarkerMain'}
  ];
  spreadsheet.addMenu('カスタムメニュー', menuItems);
}

function koteitest(){
  addFixedCellMarker()
}
function koteikaizyotest(){
  removeFixedCellMarker()
}



{
  //-----------------
  //定数定義
  //-----------------

  const SCHEDULE_SHEET_NAME = "スケジュール管理仕様";
  const UNSUPPORTED_SHEET_ERROR_MESSAGE = "選択されたシートは処理対象外です。";

  function addFixedCellMarkerMain(){
    exclusiveMain(addFixedCellMarker);
  }
  function removeFixedCellMarkerMain(){
    exclusiveMain(removeFixedCellMarker);
  }
  

  // スプレッドシートに固定セルのマーカーを付ける関数
  function addFixedCellMarker() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = spreadsheet.getActiveSheet();
    const dataSheet = spreadsheet.getSheetByName(DATA_BASE_SHEET_NAME);

    if (!isValidSheet(activeSheet)) {
      Browser.msgBox('エラー',UNSUPPORTED_SHEET_ERROR_MESSAGE, Browser.Buttons.OK);
      return;
    }

    const selectedRange = getSelectedCellRange(activeSheet);
    if (!selectedRange) {
      return;
    }
    // マーカーを付ける
    markBelowSelectedCells(selectedRange, activeSheet ,dataSheet);

  }

  // 選択されたセルの下にマーカーを付ける関数
  function markBelowSelectedCells(range, sheet, dataSheet) {
    const rowBelowSelected = range.getLastRow() + 1;
    const startingCol = range.getColumn();
    const numCols = range.getNumColumns();
    const numRows = range.getNumRows();
    let markerArray = Array(numRows).fill(Array(numCols).fill(FIXED_CELL_KEYWORD));

    // マーカーを設定
    sheet.getRange(rowBelowSelected, startingCol, numRows, numCols)
      .setValues(markerArray)
      .setFontSize(4)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
      // dataSheetに同じマーカーを設定
    dataSheet.getRange(rowBelowSelected, startingCol, numRows, numCols)
      .setValues(markerArray);
  }


// スプレッドシートの固定セルのマーカーを消す関数
function removeFixedCellMarker() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = spreadsheet.getActiveSheet();
  const dataSheet = spreadsheet.getSheetByName(DATA_BASE_SHEET_NAME);

  if (!isValidSheet(activeSheet)) {
    Browser.msgBox('エラー', UNSUPPORTED_SHEET_ERROR_MESSAGE, Browser.Buttons.OK);
    return;
  }

  const selectedRange = getSelectedCellRange(activeSheet);
  if (!selectedRange) {
    return;
  }

  // activeSheetのマーカーをクリア
  clearBelowSelectedCells(selectedRange, activeSheet);
  // dataSheetのマーカーもクリア
  clearBelowSelectedCells(selectedRange, dataSheet);
}

  // 選択されたセルの下のマークをクリアする関数
  function clearBelowSelectedCells(range, sheet) {
    const rowBelowSelected = range.getLastRow() + 1;
    const startingCol = range.getColumn();
    const numCols = range.getNumColumns();
    const numRows = range.getNumRows(); // 選択範囲の行数

    // values配列を取得してマークをクリア
    let values = sheet.getRange(rowBelowSelected, startingCol, numRows, numCols).getValues();
    for (let row = 0; row < values.length; row++) {
      for (let col = 0; col < values[row].length; col++) {
        if (values[row][col] === FIXED_CELL_KEYWORD) {
          values[row][col] = '';
        }
      }
    }

    // クリアしたvalues配列をシートに反映
    sheet.getRange(rowBelowSelected, startingCol, numRows, numCols).setValues(values);
  }
  function isValidSheet(sheet) {
    return sheet.getName() === SCHEDULE_SHEET_NAME;
  }

  function getSelectedCellRange(sheet) {
    const selection = sheet.getSelection();
    const ranges = selection.getActiveRangeList().getRanges();

    return getSelectedRange(ranges,sheet); // この関数内でエラー時はダイアログが出る。
  }
  
}
