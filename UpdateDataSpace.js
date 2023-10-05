function updateDataSpaceMain(){
  updateDataSpace()
}

{
  //-----------------
  //クラス
  //-----------------
  //話数情報
  class Story{
    constructor(){
      //作品話数名
      this.storyName = undefined
      //作品話数色
      this.storyColor = undefined
      //担当者
      this.persons = []
      //最小の開始日(カレンダー生成はマニュアルで事前にボタン押すのでここ不要)
      //this.minStartDate = undefined
    }
  }

  //担当者情報
  class Person{
    constructor(){
      //担当者名
      this.personName = undefined
      //シーン
      this.scenes = []
    }
  }

  //シーン情報
  class Scene{
    constructor(){
      //シーン名
      this.sceneName = undefined
      //開始日
      this.startDate = undefined
      //総工数
      this.totalDays = undefined
      //割込ON
      this.warikomi = undefined
      //置換シーン選択
      this.replaceScene = undefined
    }
  }
  //-----------------
  //変数
  //-----------------
  // スプレッドシート上でデータが入力されている最大範囲を選択
  let scheduleSheetAllRange = undefined
  //スケジュール表の全値取得
  let scheduleSheetDataValues = undefined
  //スケジュール表の背景色取得
  let scheduleSheetAllBackGrounds = undefined
  //枠線のスタイル
  let scheduleSheetAllBorderStyles = undefined

  //-----------------
  //関数
  //-----------------
  //メインの関数
  function updateDataSpace(){
    const label = 'updateDataSpace'
    console.time(label)
    // スプレッドシートの読み込み
    let spreadsheet =  SpreadsheetApp.getActiveSpreadsheet()

    getScheduleSheetInfo(spreadsheet)

    console.timeEnd(label)
  }

  // スケジュール表情報取得
  function getScheduleSheetInfo(scheduleSheet){
    // スプレッドシート上でデータが入力されている最大範囲を選択
    var scheduleSheetAllRange = scheduleSheet.getDataRange();
    // 値を取得
    var scheduleSheetDataValues = scheduleSheetAllRange.getValues();
    // 背景色を取得する
    var scheduleSheetAllBackGrounds = scheduleSheetAllRange.getBackgrounds();
    //console.log(scheduleSheetAllBackGrounds);  // 背景色の情報をコンソールに出力
    //console.log(scheduleSheetDataValues);  // 背景色の情報をコンソールに出力
    // セルの罫線の色を取得する
    getBorderColor(scheduleSheetAllRange);
    // セルの罫線のスタイルを取得する
    getBorderStyle(scheduleSheetAllRange) 
  }


  //罫線のbottomの位置の色の取り方
  function getBorderColor(activeRange) {
    var border = activeRange.getBorder();
    var color = border.getBottom().getColor().asRgbColor().asHexString();
    console.log(color)
    return color;
  }
  
  function getBorderStyle(activeRange) {
    var border = activeRange.getBorder();
    var style = border.getBottom().getBorderStyle()
    console.log(style)
    //logBorderStyles(activeRange)
    return style;
  }

  //一応すべての枠線の下部線のスタイルを表示する。一個ずつセルを確認するので、時間がやばい。1.1326分
  function logBorderStyles(activeRange) {
    var numRows = activeRange.getNumRows();
    var numCols = activeRange.getNumColumns();
    
    for (var i = 1; i <= numRows; i++) {
      for (var j = 1; j <= numCols; j++) {
        var cell = activeRange.getCell(i, j);
        var border = cell.getBorder();
        var style = border ? border.getBottom().getBorderStyle() : null;
        console.log('Cell (' + i + ', ' + j + '): ' + style);
      }
    }
  }
}