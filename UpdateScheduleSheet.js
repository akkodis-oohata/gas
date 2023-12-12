function updateScheduleSheetMain() {
  exclusiveMain(updateScheduleSheet);
}

function updateScheduleSheetDeadline() {
  exclusiveMain(updateDeadline);
}

function addOneMonthCalenderMain() {
  exclusiveMain(addOneMonthCalender);
}

{
  // 進行表→スケジュール表に反映させるスクリプト
  // PS:ProgressSheet(進行表)
  // SS:ScheduleSheet(スケジュール表)
  //-----------------
  //クラス
  //-----------------
  //話数情報
  class Story {
    constructor() {
      //作品話数名
      this.storyName = undefined;
      //作品話数色
      this.storyColor = undefined;
      //担当者
      this.persons = [];
    }
  }

  //担当者情報
  class Person {
    constructor() {
      //担当者名
      this.personName = undefined;
      //シーン
      this.scenes = [];
    }
  }

  //シーン情報
  class Scene {
    constructor() {
      //シーン名
      this.sceneName = undefined;
      //開始日
      this.startDate = undefined;
      //総工数
      this.totalDays = undefined;
      //割込ON
      this.warikomi = undefined;
      //置換シーン選択
      this.replaceScene = undefined;
      //工数背景色
      this.manHoursColor = undefined;
    }
  }

  //締切一覧表シート名
  const DEADLINE_SHEET_NAME = "締切一覧表";
  // 締め切りの色
  const COLOR_DEADLINE = "#ff0000";

  //-----------------
  //定数定義（進行表）
  //-----------------


  //総工数
  const PS_TOTALDAYS_COLUMN_INDEX = 6;
  //担当者名
  const PS_PERSONNAME_COLUMN_INDEX = 15;
  //開始日
  const PS_STARTDATE_COLUMN_INDEX = 16;
  //割込ON
  const PS_WARIKOMI_COLUMN_INDEX = 17;
  // 前回算出増減列
  const PS_SCENE_LAST_CHANGE_VALUE_COLUMN_INDEX = 18;
  //置換シーン選択
  const PS_REPLACESCENE_COLUMN_INDEX = 19;
  // 工数開始位置
  const PS_SCENE_MANHOUR_COLUMN_INDEX = 20;
  //再調整列
  const PS_SCENE_REARANGE_MANHOUR_COLUMN_INDEX = 14;
  
  //前回置換実行日
  const PS_SCENE_LAST_REPLACEMENT_DATE_ROW = 5;
  const PS_SCENE_LAST_REPLACEMENT_DATE_COLUMN = 16;

  // 区切り文字
  const DELIMITER = ",";
  //-----------------
  //定数定義（スケジュール表）
  //-----------------
  //カレンダー生成日付
  const SS_INPUTCALENDERDATE_ROW_INDEX = 5;
  const SS_INPUTCALENDERDATE_COLUMN_INDEX = 4;

  //カレンダー
  const SS_CALENDERDATE_ROW_INDEX = 6;
  //担当者
  const SS_PRSRON_COLUMN_INDEX = 6;
  const SS_PERSON_ROW_START_INDEX = 7;

  //エラーメッセージ
  const ERROR_MESSAGE_FULL = `メモ欄が6行全て埋まっています`;
  const ERROR_MESSAGE_DATE_MISMAYCH = `既に締切が設定されていますが、日付が不一致です`;
  const ERROR_MESSAGE_OUT_OF_DATE = `日付がスケジュールの範囲にありません`;
  // フリースペース開始位置
  const SS_OTHER_CO_START_KEYWORD = "他社ヘルプ"

  //-----------------
  //変数
  //-----------------
  // グローバル変数として行番号を保持
  let personRow = null;
  let otherRow = null;
  let freeSpaceRow = null;
  // 進行表情報
  let progressSheetAllDataRange = undefined; //進行表のすべての範囲
  let progressSheetDataValues = undefined;  //進行表上のすべての値
  let progressSheetAllBackGrounds = undefined;  //進行表上のすべての背景色
  // 進行表シーン名最終行
  let psMaxScenesRow = 0;

  //-----------------
  //関数
  //-----------------
  //メインの関数
  function updateScheduleSheet() {
    console.log("---- updateScheduleSheet IN ----");
    const label = "updateScheduleSheet";
    console.time(label);

    // スプレッドシートの読み込み
    let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // 進行表読み込み
    let progressSheet = spreadsheet.getActiveSheet();

    // 話数・シーン名チェック
    const errorMessages = checkSceneName(progressSheet,false);
    if (errorMessages.length > 0) {
      let ui = SpreadsheetApp.getUi();
      ui.alert(errorMessages.join("\n"));
      return;
    }

    // 話数情報取得
    let storyInfo = getStoryInfo(progressSheet);

    // スケジュール表読み込み
    let scheduleSheet = spreadsheet.getSheetByName(SS_SCHEDULE_SHEET_NAME);
    if (!scheduleSheet) {
      throw new Error('「' + SS_SCHEDULE_SHEET_NAME + '」シートが見つかりません');
    }
    
    // データベースシートの読み込み
    let dataBaseSheet = spreadsheet.getSheetByName(DATA_BASE_SHEET_NAME);
    if (!dataBaseSheet) {
      throw new Error('「' + DATA_BASE_SHEET_NAME + '」シートが見つかりません');
    }

    // スケジュール表情報取得
    getScheduleSheetInfoC(scheduleSheet,dataBaseSheet);

    // スケジュール表情報を進行表の値で更新
    updateDateValues(storyInfo,scheduleSheet);
    // 更新したスケジュール表情報で画面更新
    updateScheduleSheetWithDataValuesC();

    // 進行表のデータ内容を変更する。
    updateProgressSheetWarikomiSceneLastChangeValues(progressSheet,progressSheetDataValues,progressSheetAllBackGrounds,psMaxScenesRow);

    //実行日と時間を前回置換実行日に記入する
    displayCurrentDateTimeC(progressSheet,PS_SCENE_LAST_REPLACEMENT_DATE_ROW,PS_SCENE_LAST_REPLACEMENT_DATE_COLUMN)

    // スケジュール表をActiveにする
    scheduleSheet.activate();

    console.log("---- updateScheduleSheet OUT ----");
    console.timeEnd(label);
  }
  //スコープに公開
  this.updateScheduleSheet = updateScheduleSheet;



  // 開始日から日付の配列を取得する関数
  function getDateArray(startDate, numberOfDays) {
    const minDateObj = new Date(
      startDate ||
      scheduleSheetDataValues[SS_INPUTCALENDERDATE_ROW_INDEX]
      [SS_INPUTCALENDERDATE_COLUMN_INDEX]
      );
    const dateArray = [];
    for (let i = 0; i < numberOfDays; i++) {
      const dateObj = new Date(minDateObj);
      dateObj.setDate(minDateObj.getDate() + i);
      dateArray.push(dateObj);
    }
    return dateArray;
  }

  // 土日の判断
  function isWeekend(date) {
    const day = date.getDay();
    return day === 0 || day === 6; // 0は日曜日、6は土曜日
  }
  // 土日だった場合は色をつける。
  function setWeekendBackgroundColors(dateArray) {
    console.log("---setWeekendBackgroundColors ---");

    const dateColorArray = dateArray.map((date) =>
      isWeekend(date) ? COLOR_HOLIDAY : COLOR_CLEAR
    );
    let personRows = getPersonRows(scheduleSheetDataValues);
    const cellColorArray = Array.from(
      { length: personRows },
      () => dateColorArray
    );
    console.log("---setWeekendBackgroundColors end---");
    return cellColorArray;
  }

  // 'CLW美術作業者一覧'と'他社ヘルプ'の行番号を設定する関数
  function setRowNumbers(values) {
    for (let i = 0; i < values.length; i++) {
      const cellValue = values[i][0];
      
      if (cellValue === CLW_ARTWORK_PERSONS_TITLE) {
        personRow = i + 1;
      } else if (cellValue === SS_OTHER_CO_START_KEYWORD) {
        otherRow = i + 1;
      } else if (cellValue === FREE_SPACE_TITLE) {
        freeSpaceRow = i + 1;
      }
    }
    
    const notFoundItems = [];
    if (personRow === null) {
      notFoundItems.push(`'${CLW_ARTWORK_PERSONS_TITLE}'`);
    }
    if (otherRow === null) {
      notFoundItems.push(`'${SS_OTHER_CO_START_KEYWORD}'`);
    }
    if (freeSpaceRow === null) {
      notFoundItems.push(`'${FREE_SPACE_TITLE}'`);
    }
    
    if (notFoundItems.length > 0) {
      throw new Error(`${notFoundItems.join('、')} が見つかりません`);
    }
  }
  

  // 作業者一覧の行数を取得する関数
  function getPersonRows() {
    if (personRow === null || otherRow === null) {
      throw new Error("行番号が設定されていません");
    }
    
    return otherRow - personRow;
  }


  // 2次元配列の全ての行を最も要素数が多い行の長さに合わせる関数
  function alignArrayRows(array) {
    // 最も要素数が多い行を見つける
    const maxLength = Math.max(...array.map(row => row.length));
  
    // 全ての行を最も要素数が多い行の長さに合わせる
    const alignedArray = array.map(row => {
      while (row.length < maxLength) {
        row.push(''); // 空の要素を追加
      }
      return row;
    });
  
    return alignedArray;
  }

  //枠線の設定を行う
  function setBorderStyles(sheet) {
    // CLW美術作業者一覧と他社ヘルプに中太の上下のみの枠線を設定
    [personRow, otherRow].forEach((row) => {
      sheet
        .getRange(row, 1, 1, sheet.getMaxColumns())
        .setBorder(
          true,
          null,
          true,
          null,
          null,
          null,
          null,
          SpreadsheetApp.BorderStyle.SOLID_MEDIUM
        );
    });
  
    // personRowとotherRowの間で3行おきに細線を下部に設定
    for (let row = personRow + 3; row < otherRow - 1; row += 3) {
      sheet
        .getRange(row, 1, 1, sheet.getMaxColumns())
        .setBorder(
          null,
          null,
          true,
          null,
          null,
          null,
          null,
          SpreadsheetApp.BorderStyle.SOLID
        );
    }
  
    // otherRowとfreeSpaceRowの間で3行おきに細線を下部に設定
    for (let row = otherRow + 3; row < freeSpaceRow - 1; row += 3) {
      sheet
        .getRange(row, 1, 1, sheet.getMaxColumns())
        .setBorder(
          null,
          null,
          true,
          null,
          null,
          null,
          null,
          SpreadsheetApp.BorderStyle.SOLID
        );
    }
  
    // 以下フリースぺースに2重の上下のみの枠線を設定
    sheet
      .getRange(freeSpaceRow, 1, 1, sheet.getMaxColumns())
      .setBorder(
        true,
        null,
        true,
        null,
        null,
        null,
        null,
        SpreadsheetApp.BorderStyle.DOUBLE
      );
  }

  // +1か月カレンダーに追加を行う関数
  function addOneMonthCalender(){
    const label = "addOneMonthCalender";
    console.time(label);
    // スプレッドシートの読み込み
    let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    // スケジュール表読み込み
    let scheduleSheet = spreadsheet.getSheetByName(SS_SCHEDULE_SHEET_NAME);
    //データベースシートの読み込み
    let dataBaseSheet = spreadsheet.getSheetByName(DATA_BASE_SHEET_NAME);
    // スケジュール表情報取得
    getScheduleSheetInfoC(scheduleSheet,dataBaseSheet);
    // 'CLW美術作業者一覧'と'他社ヘルプ'の行番号を設定する関数
    setRowNumbers(scheduleSheetDataValues)
    
    // 現在のスケジュールの最終日を取得
    const { lastDate, lastDateColumnIndex } = getLastDateAndIndex(
      scheduleSheetDataValues
    );

    const startDate = new Date(lastDate);
    // 最終日の翌日をStartDateとする
    startDate.setDate(lastDate.getDate() + 1);
    // +1か月分の日付配列を取得
    const dateArray = getDateArray(startDate, 30);
    const dateArray2D = [dateArray];
    
    // 色を設定する。
    const scheduleBackGrounds = setWeekendBackgroundColors(dateArray); //土日
    
    // カレンダーに日付を設定
    setScheduleFromPosition(
      scheduleSheet,
      dateArray2D,
      SS_CALENDERDATE_ROW_INDEX + 1,
      lastDateColumnIndex + 2
    );
    setBackgroundFromPosition(
      scheduleSheet,
      scheduleBackGrounds,
      SS_CALENDERDATE_ROW_INDEX + 1,
      lastDateColumnIndex + 2
    );

    // フォーマットを枠線を記入する
    setBorderStyles(scheduleSheet)

    console.timeEnd(label);

    function getLastDateAndIndex(sheetDataValues) {
      const rowIndex = SS_CALENDERDATE_ROW_INDEX ;
      let lastDateColumnIndex = sheetDataValues[rowIndex].length - 1;
    
      // 空欄をスキップして最後の日付を探す
      while (
        lastDateColumnIndex >= 0 &&
        !sheetDataValues[rowIndex][lastDateColumnIndex]
      ) {
        lastDateColumnIndex--;
      }
    
      if (lastDateColumnIndex < 0) {
        throw new Error('日付が入力されているセルが見つかりません。');
      }
    
      const lastDateValue = sheetDataValues[rowIndex][lastDateColumnIndex];
      const lastDate = new Date(lastDateValue);
    
      return { lastDate, lastDateColumnIndex };
    }


    // スプレッドシートの指定された位置から2次元配列のデータを書き込む関数
    function setScheduleFromPosition(sheet, values, startRow, startColumn) {
      // 配列のサイズを取得
      const numRows = values.length;
      const numCols = values[0].length;

      // シートの現在の列数と行数が不足していれば、列と行を追加
      if (sheet.getMaxColumns() < startColumn + numCols - 1) {
        sheet.insertColumnsAfter(
          sheet.getMaxColumns(),
          startColumn + numCols - 1 - sheet.getMaxColumns()
        );
      }
      if (sheet.getMaxRows() < startRow + numRows - 1) {
        sheet.insertRowsAfter(
          sheet.getMaxRows(),
          startRow + numRows - 1 - sheet.getMaxRows()
        );
      }

      // スプレッドシートの書き込みたい範囲を再指定
      const range = sheet.getRange(startRow, startColumn, numRows, numCols);

      // データを書き込み
      range.setValues(values);
    }
    // スプレッドシートの指定された位置から2次元配列の背景色を設定する関数
    function setBackgroundFromPosition(sheet, colors, startRow, startColumn) {
      // 配列のサイズを取得
      const numRows = colors.length;
      const numCols = colors[0].length;

      // シートの現在の列数と行数が不足していれば、列と行を追加
      if (sheet.getMaxColumns() < startColumn + numCols - 1) {
        sheet.insertColumnsAfter(
          sheet.getMaxColumns(),
          startColumn + numCols - 1 - sheet.getMaxColumns()
        );
      }
      if (sheet.getMaxRows() < startRow + numRows - 1) {
        sheet.insertRowsAfter(
          sheet.getMaxRows(),
          startRow + numRows - 1 - sheet.getMaxRows()
        );
      }

      // スプレッドシートの設定したい範囲を再指定
      const range = sheet.getRange(startRow, startColumn, numRows, numCols);

      // 背景色を設定
      range.setBackgrounds(colors);
    }

  }
  
  // 進行表の割込みON列と前回算出増減のところのみ、データ内容を変更する
  function updateProgressSheetWarikomiSceneLastChangeValues(progressSheet, progressSheetDataValues, progressSheetAllBackGrounds,psMaxScenesRow) {
    const startRow = SS_PERSON_ROW_START_INDEX; // 開始行番号

    // 各列のデータを更新するための配列を初期化
    const updatedWarikomiValues = [];
    const updatedSceneLastChangeValues = [];
    const updatedBackgrounds = [];

    // データの更新
    updateWarikomiSceneLastChangeValues(progressSheetDataValues, progressSheetAllBackGrounds, updatedWarikomiValues, updatedSceneLastChangeValues, updatedBackgrounds);

    // スプレッドシートに変更を反映
    setColumnValues(progressSheet, startRow, PS_WARIKOMI_COLUMN_INDEX, updatedWarikomiValues);
    setColumnValues(progressSheet, startRow, PS_SCENE_LAST_CHANGE_VALUE_COLUMN_INDEX, updatedSceneLastChangeValues);
    setBackgrounds(progressSheet, startRow, PS_SCENE_LAST_CHANGE_VALUE_COLUMN_INDEX, updatedBackgrounds);

    // データの更新を行う関数
    function updateWarikomiSceneLastChangeValues(dataValues, backgrounds, warikomiValues, sceneLastChangeValues, updatedBackgrounds) {
      for (let i = SS_PERSON_ROW_START_INDEX - 1; i < psMaxScenesRow; i++) {
        if (dataValues[i][PS_REPLACESCENE_COLUMN_INDEX] === true) {
          // チェックボックスがtrueの場合の処理
          dataValues[i][PS_WARIKOMI_COLUMN_INDEX] = false; // 割込みON列をOFFにする
          dataValues[i][PS_SCENE_LAST_CHANGE_VALUE_COLUMN_INDEX] = 0; // 前回算出増減の値を0にする
          backgrounds[i][PS_SCENE_LAST_CHANGE_VALUE_COLUMN_INDEX] = COLOR_CLEAR; // 背景を透明にする
        }
        // 更新された列の値を配列に追加
        warikomiValues.push([dataValues[i][PS_WARIKOMI_COLUMN_INDEX]]);
        sceneLastChangeValues.push([dataValues[i][PS_SCENE_LAST_CHANGE_VALUE_COLUMN_INDEX]]);
        updatedBackgrounds.push([backgrounds[i][PS_SCENE_LAST_CHANGE_VALUE_COLUMN_INDEX]]);
      }
    }
    
    // 特定の列に値を設定する関数
    function setColumnValues(sheet, startRow, columnIndex, values) {
      const range = sheet.getRange(startRow, columnIndex + 1, values.length, 1);
      range.setValues(values);
    }
    
    // 背景色を設定する関数
    function setBackgrounds(sheet, startRow, columnIndex, backgrounds) {
      const range = sheet.getRange(startRow, columnIndex + 1, backgrounds.length, 1);
      range.setBackgrounds(backgrounds);
    }
  }



  // スケジュール表情報を進行表の値で更新
  function updateDateValues(storyInfo,scheduleSheet) {
    const maxDays = getLastFilledColumnInCalender(scheduleSheet,SS_CALENDERDATE_COLUMN_INDEX);  //カレンダーの最大サイズを確認する。

    // 「以下フリースペース」が一列目に書いてある行を取得
    const freeSpaceRowIndex = scheduleSheetDataValues.findIndex(
      (row) => row[0] == FREE_SPACE_TITLE
    );
    // 「CLW美術作業者一覧」が一列目に書いてある行を取得
    const personStartRowIndex = scheduleSheetDataValues.findIndex(
      (row) => row[0] == CLW_ARTWORK_PERSONS_TITLE
    );
    // 「他社ヘルプ」が一列目に書いてある行を取得
    const otherCOStartRowIndex = scheduleSheetDataValues.findIndex(
      (row) => row[0] == SS_OTHER_CO_START_KEYWORD
    );

    for (let storyInfoPerson of storyInfo.persons) {
      if (storyInfoPerson.personName == "") continue;
    
      const rowIndex = scheduleSheetDataValues.findIndex((row, index) => {
        let isCorrectInterval = false;
        if (index > personStartRowIndex && index < otherCOStartRowIndex) {
          // personStartRowIndex と otherCOStartRowIndex の間の行のみをチェック
          isCorrectInterval = (index - personStartRowIndex) % 3 === 1;
        } else if (index > otherCOStartRowIndex && index < freeSpaceRowIndex) {
          // otherCOStartRowIndex と freeSpaceRowIndex の間の行のみをチェック
          isCorrectInterval = (index - otherCOStartRowIndex) % 3 === 1;
        }
    
        return (
          row[SS_PRSRON_COLUMN_INDEX] == storyInfoPerson.personName &&
          index < freeSpaceRowIndex && isCorrectInterval
        );
      });

    

      if (rowIndex > 0) {  
        // 以下メソッドをrowIndexを受けて動作するように変更
        deletePersonTasks(storyInfo, storyInfoPerson, rowIndex,maxDays);
        updatePersonTasks(storyInfo, storyInfoPerson, rowIndex,maxDays);
        //先頭セルにシーン名を表示する
        displaySceneNameC(rowIndex,maxDays);
      }
    }
  }

  // 対象の担当者のシーン情報を削除する
  function deletePersonTasks(storyInfo, person, rowIndex,maxDay) {
    //console.log('---- deletePersonTasks IN ----')
    const sceneTitles = [];
    for (let scene of person.scenes) {
      if (!scene.replaceScene) {
        continue;
      }
      sceneTitles.push(storyInfo.storyName + DELIMITER + scene.sceneName);
    }
    if (sceneTitles.length == 0) {
      return;
    }
    for (
      let i = SS_CALENDERDATE_COLUMN_INDEX;
      i < SS_CALENDERDATE_COLUMN_INDEX + maxDay;
      i++
    ) {
      let tmpSceneTitle = dataBaseSheetDataValues[rowIndex][i];
      if (
        tmpSceneTitle &&
        sceneTitles.findIndex((sceneTitle) =>
          tmpSceneTitle.startsWith(sceneTitle)
        ) > -1
      ) {
        // startsWithメソッドを使用してチェック
        // セルの内容をクリア
        scheduleSheetDataValues[rowIndex][i] = "";
        scheduleSheetDataValues[rowIndex + 1][i] = "";
        dataBaseSheetDataValues[rowIndex][i] = "";
        dataBaseSheetDataValues[rowIndex + 1][i] = "";
        // セルの背景色をクリア
        scheduleSheetAllBackGrounds[rowIndex][i] = COLOR_CLEAR;
        scheduleSheetAllBackGrounds[rowIndex + 1][i] = COLOR_CLEAR;
      }
    }
  }
  
  // 開始日がスケジュール表にない場合はエラーのダイアログを出すようにする。
  // 対象の担当者のシーン情報を更新する
  function updatePersonTasks(storyInfo, person, rowIndex,maxDays) {
    for (let scene of person.scenes) {
      if (!scene.replaceScene) {
        continue;
      }
      // 開始日の列番号を特定
      let startColumnIndex = undefined;
      // dateがDate型でない場合にエラーを投げる
      if (!(scene.startDate instanceof Date)) {
        throw new Error('「' + scene.sceneName + '」の開始日（' + scene.startDate + '）が正しい日付形式ではないため、処理を中断いたしました。');
      }
      let sceneStartDateStr = Utilities.formatDate(
        new Date(scene.startDate),
        "JST",
        "yyyy-MM-dd"
      );
      for (
        let i = SS_CALENDERDATE_COLUMN_INDEX;
        i < maxDays;
        i++
      ) {
        let date = scheduleSheetDataValues[SS_CALENDERDATE_ROW_INDEX][i];

        let dateStr = Utilities.formatDate(date, "JST", "yyyy-MM-dd");
        if (sceneStartDateStr == dateStr) {
          startColumnIndex = i;
          break;
        }
      }
      // 開始日がスケジュール表にない場合はエラーを投げる
      if (startColumnIndex === undefined) {
        throw new Error('「'+ scene.sceneName + '」' + 'の開始日(' + sceneStartDateStr + ')がスケジュール表に見つからないため、処理を中断いたしました。');
      }

      // 開始日から総工数（日）だけセルに色を塗る
      if (scene.totalDays > 0) {
        let sceneTitle =
          storyInfo.storyName +
          DELIMITER +
          scene.sceneName +
          DELIMITER +
          scene.totalDays +
          "日分" +
          DELIMITER +
          formatDate(scene.startDate);
        if (scene.warikomi) {
          warikomiMoveCellC(
            rowIndex,
            startColumnIndex,
            Number(scene.totalDays),
            storyInfo.storyColor,
            sceneTitle,
            scene.manHoursColor
          );
        } else {
          fillClearCellC(
            rowIndex,
            startColumnIndex,
            Number(scene.totalDays),
            storyInfo.storyColor,
            sceneTitle,
            scene.manHoursColor
          );
        }
      }

    }
    // yyyy-mm-ddの形に変換
    function formatDate(date) {
      // dateがDateオブジェクトであることを確認
      if (!(date instanceof Date)) {
        date = new Date(date);
      }
    
      let year = date.getFullYear();
      let month = (date.getMonth() + 1).toString().padStart(2, '0');  // 月は0から始まるため、1を加えます
      let day = date.getDate().toString().padStart(2, '0');
    
      return `${year}-${month}-${day}`;
    }
  }

  //土日祝日判定
  function isHoliday(date,calenderId) {
    //土日の判定
    const day = date.getDay();
    if (day === 0 || day === 6) return true;
    //祝日の取得。ここを使うにはGoogleDrive上で一度「カレンダー」にアクセスする必要あり

    const events = calenderId.getEventsForDay(date);
    //なんらかのイベントがある＝祝日
    if (events.length) return true;

    return false;
  }

  // 進行表情報を読み込みクラスにする
  function getStoryInfo(progressSheet) {
    progressSheetAllDataRange = progressSheet.getDataRange();
    // 値を取得
    progressSheetDataValues = progressSheetAllDataRange.getValues();
    // 背景色を取得する
    progressSheetAllBackGrounds = progressSheetAllDataRange.getBackgrounds();
    // シーン名が何行までか確認する。
    psMaxScenesRow = getPsSceneRowsIndex(progressSheetAllDataRange) + 1;

    //進行表情報
    let storyInfo = new Story();
    //話数
    storyInfo.storyName =
      progressSheetDataValues[PS_STORYNAME_ROW_INDEX][
        PS_STORYNAME_COLUMN_INDEX
      ];
    storyInfo.storyColor =
      progressSheetAllBackGrounds[PS_STORYNAME_ROW_INDEX][
        PS_STORYNAME_COLUMN_INDEX
      ];

    //シーン情報・担当者情報
    for (
      let i = PS_SCENENAME_START_ROW_INDEX;
      i < psMaxScenesRow;
      i++
    ) {
      let personName = progressSheetDataValues[i][PS_PERSONNAME_COLUMN_INDEX];
      let person = undefined;
      for (let tmpperson of storyInfo.persons) {
        if (tmpperson.personName === personName) {
          person = tmpperson;
        }
      }
      if (person === undefined) {
        person = new Person();
        person.personName = personName;
        storyInfo.persons.push(person);
      }

      let scene = new Scene();
      scene.sceneName = progressSheetDataValues[i][PS_SCENENAME_COLUMN_INDEX];
      scene.startDate = progressSheetDataValues[i][PS_STARTDATE_COLUMN_INDEX];
      scene.totalDays = progressSheetDataValues[i][PS_TOTALDAYS_COLUMN_INDEX] + progressSheetDataValues[i][PS_SCENE_REARANGE_MANHOUR_COLUMN_INDEX];  //総工数＋再調整の値にする
      scene.warikomi = progressSheetDataValues[i][PS_WARIKOMI_COLUMN_INDEX];
      scene.replaceScene =
        progressSheetDataValues[i][PS_REPLACESCENE_COLUMN_INDEX];
      // 工数背景色
      const manHoursEndIndex = progressSheetDataValues[i].findIndex(
        (data, index) => data != "" && index >= PS_SCENE_MANHOUR_COLUMN_INDEX
      );  //工数背景色が何列目までにあるか取得する。
      
      scene.manHoursColor = progressSheetAllBackGrounds[i].slice(
        PS_SCENE_MANHOUR_COLUMN_INDEX,
        manHoursEndIndex + 1
      );

      person.scenes.push(scene);
    }

    return storyInfo;
  }

  // 進行表のシーン名の行数(空データに当たるまで)を確認する。//空の行のindexを返す。
  function getPsSceneRowsIndex(progressSheetAllDataRange){
    // 値を取得
    let progressValues = progressSheetAllDataRange.getValues();
    //シーン名 空データにあたるまで
    let rowIndex = PS_SCENENAME_START_ROW_INDEX;
    while (progressValues[rowIndex][PS_SCENENAME_COLUMN_INDEX] != "") {
      rowIndex++;
    }
    return rowIndex - 1;  //最後のシーン名が記入のある行を返す。
  }

  // 締切メモ欄のクリア
  function clearUpperMemos() {
    const SS_CALENDER_COLUMN_INDEX = 7;
    // すでにある締切はクリアしておく
    for (let row = 0; row < SS_CALENDERDATE_ROW_INDEX; row++) {
      for (let column = SS_CALENDER_COLUMN_INDEX; column < scheduleSheetAllBackGrounds[row].length; column++) {
       if(scheduleSheetAllBackGrounds[row][column] == COLOR_DEADLINE){
        scheduleSheetAllBackGrounds[row][column] = COLOR_CLEAR;
        scheduleSheetDataValues[row][column] = "";
       }
      }  
    }
  }

  

  // 締切一覧スケジュール表反映
  function updateDeadline() {
    // スプレッドシートの読み込み
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // 締切一覧表読み込み
    const deadlineSheet = spreadsheet.getSheetByName(DEADLINE_SHEET_NAME);
    // スプレッドシート上でデータが入力されている最大範囲を選択
    const deadlineSheetAllRange = deadlineSheet.getDataRange();
    // 先頭2行分を削除した値を取得し、話数と日付を取得
    const deadlines = deadlineSheetAllRange.getValues().slice(2);

    // スケジュール表読み込み
    const scheduleSheet = spreadsheet.getSheetByName(SS_SCHEDULE_SHEET_NAME);
    // データベース表読み込み
    const dataBaseSheet = spreadsheet.getSheetByName(DATA_BASE_SHEET_NAME);
    // スケジュール表情報取得
    getScheduleSheetInfoC(scheduleSheet,dataBaseSheet);

    // カレンダー部分とメモの切り出し
    const cal = scheduleSheetDataValues[SS_CALENDERDATE_ROW_INDEX];
    const memos = scheduleSheetDataValues.slice(0, SS_CALENDERDATE_ROW_INDEX);

    //締切メモ欄のクリア
    clearUpperMemos();

    // エラー用変数
    let errorMessages = [];

    // 締切毎に処理を行う
    deadlines.map((deadline) => {
      // メモ欄に話数と一致する文字列が存在する場合、位置を取得
      let memoIndex = [];
      memos.map((rowMemo, rowIndex) => {
        let clumnIndex = rowMemo.findIndex((memo) => memo === deadline[0]);
        if (clumnIndex > 0) {
          memoIndex = [rowIndex, clumnIndex];
        }
      });

      // カレンダー部分に締切と一致する日付が存在する場合、位置を取得
      let calIndex = cal.findIndex((date) => isSameDate(date, deadline[1]));

      // カレンダーに一致する日付がない場合、当締切の処理終了
      if (calIndex < 0) {
        errorMessages.push(`${deadline[0]}: ${ERROR_MESSAGE_OUT_OF_DATE}`);
        return;
      }

      if (memoIndex.length == 0) {
        // メモ欄に一致がない場合、締切列の空白部分に話数を設定する
        for (let row = 0; row < SS_CALENDERDATE_ROW_INDEX; row++) {
          if (scheduleSheetDataValues[row][calIndex] == "") {
            scheduleSheetDataValues[row][calIndex] = deadline[0];
            scheduleSheetAllBackGrounds[row][calIndex] = COLOR_DEADLINE;
            return;
          }
        }
        // メモ欄が6行全部埋まっている
        errorMessages.push(`${deadline[0]}: ${ERROR_MESSAGE_FULL}`);
      } else if (memoIndex[1] != calIndex) {
        // 締め切りの日付が不一致
        errorMessages.push(`${deadline[0]}: ${ERROR_MESSAGE_DATE_MISMAYCH}`);
      }
    });

    updateScheduleSheetWithDataValuesC();

    // スケジュール表をActiveにする
    scheduleSheet.activate();

    // エラーメッセージがあればダイアログとして表示
    if (errorMessages.length > 0) {
      let ui = SpreadsheetApp.getUi();
      ui.alert(errorMessages.join("\n"));
    }
  }

  // 同日チェック
  function isSameDate(strDate1, strDate2) {
    if (!strDate1 || !strDate2) return false;
    const date1 = new Date(strDate1);
    const date2 = new Date(strDate2);

    if (Number.isNaN(date1.getTime()) || Number.isNaN(date2.getTime()))
      return false;

    return (
      date1.getFullYear() === date2.getFullYear() &&
      date1.getMonth() === date2.getMonth() &&
      date1.getDate() === date2.getDate()
    );
  }

}