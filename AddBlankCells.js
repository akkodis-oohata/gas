/*
空白セル追加処理
*/
function addBlankCellsMain(){
  addBlankCells()
}


{
  //-----------------
  //変数
  //-----------------
  // スプレッドシート上でデータが入力されている最大範囲を選択
  let scheduleSheetAllRange = undefined
  //スケジュール表の全値取得
  let scheduleSheetDataValues = undefined
  //スケジュール表の背景色取得
  let scheduleSheetAllBackGrounds = undefined

  function addBlankCells(){
    console.log('addBlankCells in')
    // スプレッドシートの読み込み
    let spreadsheet =  SpreadsheetApp.getActiveSpreadsheet()
    let sheet = spreadsheet.getActiveSheet();

    //データスペース（作品話数ベースデータ）へ反映を行う。
    updateDataSpaceMain(sheet)

    // スケジュール表情報取得(データスペース反映後)
    getScheduleSheetInfo(sheet);

    // TODO 空白セル追加処理

    // 選択したセルのrange取得
    var selection = sheet.getSelection()
    var ranges =  selection.getActiveRangeList().getRanges();
    for (var i = 0; i < ranges.length; i++) {
      console.log('Active Ranges: ' + ranges[i].getA1Notation());
      
    }

    // scheduleSheetDataValuesでレンジを確認する。
    
    // 選択したセルの行が担当者行が入っているか？
    
    // 固定セルの記入は消す
    // 担当者行のセルとステータスセルを移動する。
    //

    // TODO 千葉さんに聞く。UpdateScheduleSheet.js/movecellは使えるのか？

    console.log('addBlankCells out')

  }
  // スケジュール表情報取得
  function getScheduleSheetInfo(scheduleSheet){
    // スプレッドシート上でデータが入力されている最大範囲を選択
    scheduleSheetAllRange = scheduleSheet.getDataRange();
    // 値を取得
    scheduleSheetDataValues = scheduleSheetAllRange.getValues();
    // 背景色を取得する
    scheduleSheetAllBackGrounds = scheduleSheetAllRange.getBackgrounds();
  }

  
}