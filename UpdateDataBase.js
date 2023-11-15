function UpdateDataBaseMain(){
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  let sheet = spreadsheet.getSheetByName("スケジュール管理仕様");
  let dataBaseSheet = spreadsheet.getSheetByName("データベース");

  exclusiveMain(function() {
    updateDataBaseSheet(sheet, dataBaseSheet);
  });
  
}

{
  //-----------------
  //定数定義
  //-----------------
  const FREE_SPACE_KEYWORD = "以下フリースぺース";
  const DATABASE_START_COLUMN_INDEX = 6;
  const DATABASE_START_ROW_INDEX = 7;


  // スケジュール表の内容をコピーしデータベースを更新するmain関数
  function updateDataBaseSheet(sourceSheet, dataBaseSheet) {
    console.log("---- updateDataBaseSheet2 IN ----");
    const label = "updateDataBaseSheet2";
    console.time(label);
    
    // シートの値を全取得
    getScheduleSheetInfoC(sourceSheet, dataBaseSheet);
    
    // KEY：話数#,シーン名、VALUE：日数,開始日でハッシュ値作成
    //let episodeSceneScheduleMap = createEpisodeSceneScheduleMap(dataBaseSheetDataValues);

    //dataBaseSheetの値を全削除
    dataBaseSheet.clearContents();
    
    //以下フリースペース部分を削除
    scheduleSheetDataValues = removeRowsAfterFreeSpace(scheduleSheetDataValues)
    
    // 特定のstoryColorに基づいて、対応するセルにデータを埋め込む関数。
    scheduleSheetDataValues = fillValuesForStoryColors(scheduleSheetDataValues,scheduleSheetAllBackGrounds)

    // データをデータベースシートにセット
    setDataToSheet(dataBaseSheet, scheduleSheetDataValues);
  
    console.log("---- updateDataBaseSheet2 OUT ----");
    console.timeEnd(label);
  }

  //以下フリースペース部分を削除
  function removeRowsAfterFreeSpace(dataBaseSheetDataValues) {
    let copiedData = [...dataBaseSheetDataValues]; // 配列のコピーを作成
    for (let i = 0; i < copiedData.length; i++) {

      if (copiedData[i][0] === FREE_SPACE_KEYWORD) {
        copiedData.splice(i);
        break;
      }
    }
    return copiedData;
  }

  // 特定のstoryColorに基づいて、対応するセルにデータを埋め込む関数。
  function fillValuesForStoryColors(editedData, scheduleSheetAllBackGrounds) {
    const startColumnIndex = DATABASE_START_COLUMN_INDEX;
    const startRowIndex = DATABASE_START_ROW_INDEX;

    for (let i = startRowIndex; i < editedData.length; i++) {
      let storyColor = undefined;
      let storyValue = undefined;
      for (let j = startColumnIndex; j < editedData[i].length; j++) {
        let scheduleValue = editedData[i][j];
        let currentCellColor = scheduleSheetAllBackGrounds[i][j];
        // 透明だった場合、または土日の色だった場合、次のセルへ
        if (!currentCellColor || currentCellColor === "#ffffff" || currentCellColor === COLOR_HOLIDAY) continue;
        
        // storyColorが異なる場合、これをアップデートします
        if (storyColor !== currentCellColor) {
            storyColor = currentCellColor;
        }
        // 値が入っている時は更新します。
        if (scheduleValue){
          storyValue = scheduleValue;
        }
        //条件が合致する為、データベース用にシーンの一番最初に書いてある情報を格納します。
        if(scheduleSheetAllBackGrounds[i][j] === storyColor){
          editedData[i][j] = storyValue;
        }
      }
    }

    return editedData;
  }

  // KEY：話数#,シーン名、VALUE：日数,開始日でハッシュ値作成
  function createEpisodeSceneScheduleMap(dataBaseSheetDataValues) {
    let episodeSceneScheduleMap = {};
  
    // すべての行に対して
    dataBaseSheetDataValues.forEach(row => {
      //同じ行の作業者の名前を取得する。
      let personName = row[6]
      // すべての列（セル）に対して
      row.forEach(cellValue => {
        // セルが空でなく、文字列である場合のみ処理
        if(cellValue && typeof cellValue === 'string') {
          // セルが空でない場合のみ処理
          if(cellValue) {
            // セルの内容をカンマで分割
            let parts = cellValue.split(",");
    
            // 分割した値をそれぞれの変数に格納
            let episodeNumber = parts[0];
            let sceneName = parts[1];
            let days = parts[2];
            let startDate = parts[3];
    
            // KEY：話数#,シーン名,作業者名、VALUE：開始日
            let key = `${episodeNumber},${sceneName},${personName}`;
            let value = `${startDate}`;
    
            episodeSceneScheduleMap[key] = value;
          }
        }
      });
    });
    return episodeSceneScheduleMap;
  }


  function setDataToSheet(sheet, data) {
    // データの行数と列数を取得
    const numRows = data.length;
    const numCols = data[0].length;
    //let spreadsheetId = sheet.get();
    // シートの指定された範囲にデータをセット
    let range = sheet.getRange(1, 1, numRows, numCols)
    range.setValues(data);
  }
  
}
