function UpdateDataBaseMain(){
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  let sheet = spreadsheet.getSheetByName(SS_SCHEDULE_SHEET_NAME);
  let dataBaseSheet = spreadsheet.getSheetByName(DATA_BASE_SHEET_NAME);

  exclusiveMain(function() {
    updateDataBaseSheet(sheet, dataBaseSheet);
  });
  
}

{
  //-----------------
  //定数定義
  //-----------------
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
    let episodeSceneScheduleMap = createEpisodeSceneScheduleMap(dataBaseSheetDataValues);

    
    //以下フリースペース部分を削除
    scheduleSheetDataValues = removeRowsAfterFreeSpace(scheduleSheetDataValues)
    
    // 特定のstoryColorに基づいて、対応するセルにデータを埋め込む関数。
    scheduleSheetDataValues = fillValuesForStoryColors(scheduleSheetDataValues,scheduleSheetAllBackGrounds,episodeSceneScheduleMap)
    //dataBaseSheetの値を全削除
    dataBaseSheet.clearContents();
    
    // データをデータベースシートにセット
    setDataToSheet(dataBaseSheet, scheduleSheetDataValues);
  
    console.log("---- updateDataBaseSheet2 OUT ----");
    console.timeEnd(label);
  }

  //以下フリースペース部分を削除
  function removeRowsAfterFreeSpace(dataBaseSheetDataValues) {
    let copiedData = [...dataBaseSheetDataValues]; // 配列のコピーを作成
    for (let i = 0; i < copiedData.length; i++) {

      if (copiedData[i][0] === FREE_SPACE_TITLE) {
        copiedData.splice(i+1);
        break;
      }
    }
    return copiedData;
  }

  // 特定のstoryColorに基づいて、対応するセルにデータを埋め込む関数。
  function fillValuesForStoryColors(scheduleSheetDataValues, scheduleSheetAllBackGrounds,episodeSceneScheduleMap) {
    const startColumnIndex = DATABASE_START_COLUMN_INDEX;
    const sceneRows = getSceneRows(scheduleSheetDataValues);

    sceneRows.forEach(rowIdx => {
      let storyColor = undefined;
      let storyValue = undefined;
      for (let j = startColumnIndex; j < scheduleSheetDataValues[rowIdx].length; j++) {
        let scheduleValue = scheduleSheetDataValues[rowIdx][j];
        let currentCellColor = scheduleSheetAllBackGrounds[rowIdx][j];
        if (!currentCellColor || currentCellColor === COLOR_CLEAR || currentCellColor === COLOR_HOLIDAY) continue;
  
        if (storyColor !== currentCellColor) {
          storyColor = currentCellColor;
        }
        if (scheduleValue) {
          storyValue = scheduleValue;
        }
        if (scheduleSheetAllBackGrounds[rowIdx][j] === storyColor) {
          scheduleSheetDataValues[rowIdx][j] = createDatabaseTitle(storyValue, scheduleSheetDataValues, rowIdx, episodeSceneScheduleMap);
        }
      }
    });

    return scheduleSheetDataValues;
    // ストーリーの情報からデータベースタイトルを作成する関数
    function createDatabaseTitle(storyValue, scheduleSheetDataValues, i, episodeSceneScheduleMap) {
      if (storyValue === "" || storyValue === undefined) {
        return ""; // storyValueが空文字の場合は、早期に空文字を返す
      }
      let parts = storyValue.split(",");
      // 分割した値をそれぞれの変数に格納
      let episodeNumber = parts[0];
      let sceneName = parts[1];
      let personName = scheduleSheetDataValues[i][DATABASE_START_COLUMN_INDEX];
      let databaseTitle = undefined;
      // スケジュール表からキーを作成
      let storyKey = `${episodeNumber},${sceneName},${personName}`;
        if (!episodeSceneScheduleMap.hasOwnProperty(storyKey)) {
          // データベースに初めて登録するデータ
          databaseTitle = `${episodeNumber},${sceneName},${undefined},${undefined}`;
        }else{
          // データベースあるデータ
          // データベースに記載があった内容からデータベースのタイトルを算出
          let databaseParts = episodeSceneScheduleMap[storyKey].split(",");
          let datadays = databaseParts[0];
          let datastartdate = databaseParts[1];
          // 総日数は変えたくないので、話数#,シーン名,総日数、開始日とする。
          databaseTitle = `${episodeNumber},${sceneName},${datadays},${datastartdate}`;
        }
      return databaseTitle;
    }
    // シーン行を取得してシーン行を返す関数
    function getSceneRows(scheduleSheetDataValues){
      let startRowIndex = scheduleSheetDataValues.findIndex(
        (row) => row[0] === CLW_ARTWORK_PERSONS_TITLE
      );
      let endRowIndex = scheduleSheetDataValues.findIndex(
        (row) => row[0] === FREE_SPACE_TITLE
      );
      let personRows = [];
      // シーン行取得
      if (startRowIndex !== -1 && endRowIndex !== -1) {
        for (let i = startRowIndex + 1; i < endRowIndex; i++) {
          if (
            scheduleSheetDataValues[i][PERSON_COLUMN_INDEX] != "" &&
            scheduleSheetDataValues[i][PERSON_COLUMN_INDEX] != MEMO_SPACE_TITLE
          ) {
            personRows.push(i);
          }
        }
      }
      return personRows
    }
  }

  // KEY：話数#,シーン名、VALUE：日数,開始日でハッシュ値作成
  function createEpisodeSceneScheduleMap(dataBaseSheetDataValues) {
    let episodeSceneScheduleMap = {};
  
    // すべての行に対して
    dataBaseSheetDataValues.forEach(row => {
      //同じ行の作業者の名前を取得する。
      let personName = row[DATABASE_START_COLUMN_INDEX]
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
    
            // KEY：話数#,シーン名,作業者名、VALUE：日分,開始日
            let key = `${episodeNumber},${sceneName},${personName}`;
            let value = `${days},${startDate}`;
    
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
