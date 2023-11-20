function updateScheduleSheetMain() {
  exclusiveMain(updateScheduleSheet);
}

function generateCalendarMain() {
  exclusiveCheck(generateCalendar);
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
      //最小の開始日(カレンダー生成はマニュアルで事前にボタン押すのでここ不要)
      //this.minStartDate = undefined
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
  //スケジュール表シート名
  const SS_SCHEDULE_SHEET_NAME = "スケジュール管理仕様";
  //締切一覧表シート名
  const DEADLINE_SHEET_NAME = "締切一覧表";
  // 締め切りの色
  const COLOR_DEADLINE = "#FF0000";

  //-----------------
  //定数定義（進行表）
  //-----------------
  //話数名
  const PS_STORYNAME_ROW_INDEX = 0;
  const PS_STORYNAME_COLUMN_INDEX = 0;
  //シーン名
  const PS_SCENENAME_COLUMN_INDEX = 1;
  const PS_SCENENAME_START_ROW_INDEX = 6;
  //総工数
  const PS_TOTALDAYS_COLUMN_INDEX = 6;
  //担当者名
  const PS_PERSONNAME_COLUMN_INDEX = 15;
  //開始日
  const PS_STARTDATE_COLUMN_INDEX = 16;
  //割込ON
  const PS_WARIKOMI_COLUMN_INDEX = 17;
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
  const SS_CALENDER_COLUMN = 8;
  const SS_PERSON_ROW_START_INDEX = 7;
  const SS_PERSON_ROW_STEP = 3; //作業者間の行間、作業者名、ステータス行、memo行
  //カレンダー生成時前にクリアする領域（適当な値）//TODO
  const SS_CLEAR_ROW_LENGTH = 54;

  //エラーメッセージ
  const ERROR_MESSAGE_FULL = `メモ欄が6行全て埋まっています`;
  const ERROR_MESSAGE_DATE_MISMAYCH = `既に締切が設定されていますが、日付が不一致です`;
  const ERROR_MESSAGE_OUT_OF_DATE = `日付がスケジュールの範囲にありません`;
  const ERROR_MESSAGE_DELIMITER = `値に区切り文字 ${DELIMITER} が含まれています:`;
  // 他社接頭語
  const OTHER_CO = "他社_";
  // フリースペース開始位置
  const SS_FREE_SPACE = "以下フリースぺース";
  const SS_PERSON_START_KEYWORD = "CLW美術作業者一覧"
  const SS_OTHER_CO_START_KEYWORD = "他社ヘルプ"

  //-----------------
  //変数
  //-----------------
  // 「以下フリースペース」が一列目に書いてある行
  let freeSpaceRowIndex = undefined;
    // グローバル変数として行番号を保持
  let personRow = null;
  let otherRow = null;
  let freeSpaceRow = null;

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
    const errorMessages = checkSceneName(progressSheet);
    if (errorMessages.length > 0) {
      let ui = SpreadsheetApp.getUi();
      ui.alert(errorMessages.join("\n"));
      return;
    }

    // 話数情報取得
    let storyInfo = getStoryInfo(progressSheet);

    // スケジュール表読み込み
    let scheduleSheet = spreadsheet.getSheetByName(SS_SCHEDULE_SHEET_NAME);
    // データベースシートの読み込み
    let dataBaseSheet = spreadsheet.getSheetByName(DATA_BASE_SHEET_NAME);
    // カレンダー設定 は事前にやっておく
    //setCallendarDate(scheduleSheet, storyInfo)
    // スケジュール表情報取得
    //getScheduleSheetInfo(scheduleSheet);

    //スケジュール表のデータスペース（作品話数ベースデータ）へ反映を先に行う。
    //updateDataSpaceMain(scheduleSheet)

    // スケジュール表情報取得(データスペース反映後)
    getScheduleSheetInfoC(scheduleSheet,dataBaseSheet);

    // スケジュール表情報を進行表の値で更新
    updateDateValues(storyInfo,scheduleSheet);
    // 更新したスケジュール表情報で画面更新
    updateScheduleSheetWithDataValuesC();

    //実行日と時間を前回置換実行日に記入する
    displayCurrentDateTimeC(progressSheet,PS_SCENE_LAST_REPLACEMENT_DATE_ROW,PS_SCENE_LAST_REPLACEMENT_DATE_COLUMN)

    // スケジュール表をActiveにする
    scheduleSheet.activate();

    console.log("---- updateScheduleSheet OUT ----");
    console.timeEnd(label);
  }
  //スコープに公開
  this.updateScheduleSheet = updateScheduleSheet;

  //カレンダー生成↓
  //カレンダー生成（11/14時点で廃止だが追加にて対応の可能性ある為、関数は残す。）
  function generateCalendar() {
    const label = "generateCalendar";
    console.time(label);
    // スプレッドシートの読み込み
    let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    // スケジュール表読み込み
    //let scheduleSheet = spreadsheet.getActiveSheet();
    let scheduleSheet = spreadsheet.getSheetByName(SS_SCHEDULE_SHEET_NAME);
    //データベースシートの読み込み
    let dataBaseSheet = spreadsheet.getSheetByName(DATA_BASE_SHEET_NAME);

    // スケジュール表情報取得
    getScheduleSheetInfoC(scheduleSheet,dataBaseSheet);

    // 美術ルーム全体(G列)より右側をすべて消す
    clearContentAndFormatting(scheduleSheet)
    
    //カレンダー日付設定
    let scheduleColors = setCalendarDate();

    // 値書き込み
    setSchedule(scheduleSheet,scheduleSheetDataValues)
    setScheduleBackGround(scheduleSheet,scheduleColors)
    
    //updateScheduleSheetWithDataValuesC();
    
    // フォーマットを枠線を記入する
    setBorderStyles(scheduleSheet)

    console.timeEnd(label);
  }
  
  // 美術ルーム全体(G列)より右側をすべて消す
  function clearContentAndFormatting(scheduleSheet) {
    const lastColumn = scheduleSheet.getLastColumn();
    const startColumn = SS_CALENDER_COLUMN;
    
    // 消去する範囲の列数を計算
    const numColumns = lastColumn - startColumn + 1;

    // 列数が1以上の場合のみ範囲をクリア
    if (numColumns > 0) {
      const range = scheduleSheet.getRange(
        1,
        startColumn,
        scheduleSheet.getMaxRows(),
        numColumns
        );
      range.clearContent();
      range.clearFormat();
    } else {
      // すでにクリアされている場合は何もしない
      return;
    }
  }


  // 領域をクリアする関数
  function clearCalendarArea(maxdays) {
    const clearRowIndex = calculateClearRowIndex();
    const clearColumIndex = calculateClearColumnIndex(maxdays);

    for (let i = SS_CALENDERDATE_ROW_INDEX; i < clearRowIndex; i++) {
      for (let j = SS_PRSRON_COLUMN_INDEX; j < clearColumIndex; j++) {
        scheduleSheetDataValues[i][j] = "";
        scheduleSheetAllBackGrounds[i][j] = COLOR_CLEAR;
      }
    }
  }

  // クリアする行のインデックスを計算する関数
  function calculateClearRowIndex() {
    return SS_CALENDERDATE_ROW_INDEX + SS_CLEAR_ROW_LENGTH < scheduleSheetDataValues.length
      ? SS_CALENDERDATE_ROW_INDEX + SS_CLEAR_ROW_LENGTH
      : scheduleSheetDataValues.length;
  }

  // クリアする列のインデックスを計算する関数
  function calculateClearColumnIndex(maxDays) {
    return SS_PRSRON_COLUMN_INDEX + maxDays + 1 < scheduleSheetDataValues[0].length
      ? SS_PRSRON_COLUMN_INDEX + maxDays + 1
      : scheduleSheetDataValues[0].length;
  }

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

  // カレンダーに日付を設定する関数
  function setDatesOnCalendar(dateArray,
     startColumn = SS_CALENDERDATE_COLUMN_INDEX
     ) {
    const scheduleSize = scheduleSheetDataValues[SS_CALENDERDATE_ROW_INDEX].length;
    console.log(`dateArray size: ${dateArray.length}, scheduleSize: ${scheduleSize}`);
  
    // dateArrayのサイズをscheduleSizeに合わせる
    if (dateArray.length < scheduleSize) {
      // 不足分のサイズをnullで埋める
      dateArray = dateArray.concat(new Array(scheduleSize - dateArray.length).fill(null));
    }
  
    // dateArrayをscheduleSheetDataValuesに設定する
    for (let j = 0; j < dateArray.length; j++) {
      if (startColumn + j < scheduleSize) {
        scheduleSheetDataValues[SS_CALENDERDATE_ROW_INDEX][startColumn + j] = dateArray[j];
      } else {
        console.error('dateArray has more dates than the schedule sheet can accommodate.');
        break;
      }
    }
  }
  
  


  // 土日や祝日の背景色を設定する関数
  function setHolidayBackgroundColors(dateArray) {
    console.log("---setHolidayBackgroundColors---");

    // 日本の祝日を含むGoogleカレンダーのID
    const id = "ja.japanese#holiday@group.v.calendar.google.com";
    // 指定されたIDを持つカレンダーを取得し、cal変数に代入
    const cal = CalendarApp.getCalendarById(id);

    // 休日かどうかを判定し、背景色の配列を作成
    const dateColorArray = dateArray.map(date => isHoliday(date, cal) ? COLOR_HOLIDAY : COLOR_CLEAR);
    
    // 作業者一覧の行数を取得する
    let personRows = getPersonRows(scheduleSheetDataValues)
    console.log(personRows)

    // 全ての行に対して背景色の配列を設定
    const cellColorArray = Array.from({ length: personRows }, () => dateColorArray);

    console.log("---setHolidayBackgroundColors end---");
    return cellColorArray

    
  
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
      
      if (cellValue === SS_PERSON_START_KEYWORD) {
        personRow = i + 1;
      } else if (cellValue === SS_OTHER_CO_START_KEYWORD) {
        otherRow = i + 1;
      } else if (cellValue === SS_FREE_SPACE) {
        freeSpaceRow = i + 1;
      }
    }
    
    const notFoundItems = [];
    if (personRow === null) {
      notFoundItems.push(`'${SS_PERSON_START_KEYWORD}'`);
    }
    if (otherRow === null) {
      notFoundItems.push(`'${SS_OTHER_CO_START_KEYWORD}'`);
    }
    if (freeSpaceRow === null) {
      notFoundItems.push(`'${SS_FREE_SPACE}'`);
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

  

  // カレンダー日付設定のメイン関数
  function setCalendarDate() {
    //clearCalendarArea();
    setRowNumbers(scheduleSheetDataValues);
    const dateArray = getDateArray();
    setDatesOnCalendar(dateArray);
    //const scheduleBackGrounds = setHolidayBackgroundColors(dateArray); //祝日有
    const scheduleBackGrounds = setWeekendBackgroundColors(dateArray)  //土日のみ //TODO:千葉さんに相談。
    return scheduleBackGrounds

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

  // スプレッドシートの指定範囲に2次元配列のデータを書き込む関数
  function setSchedule(sheet, values) {
  
    // valuesの各行を最も要素数が多い行に合わせる
    const alignedValues = alignArrayRows(values);
  
    // 配列のサイズを取得
    const numRows = alignedValues.length;
    const numCols = alignedValues[0].length;
  
    // シートの現在の列数が不足していれば、列を追加
    if (sheet.getMaxColumns() < numCols) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), numCols - sheet.getMaxColumns());
    }
  
    // スプレッドシートの書き込みたい範囲を再指定
    const range = sheet.getRange(1, 1, numRows, numCols);
  
    // データを書き込み
    range.setValues(alignedValues);
  }

  // スプレッドシートの特定範囲に背景色の配列を設定する関数
  function setScheduleBackGround(sheet, colors) {
  
    // シートの対象範囲を指定
    const startRow = SS_CALENDERDATE_ROW_INDEX + 1; // getRangeは1インデックス
    const startCol = SS_CALENDERDATE_COLUMN_INDEX + 1; // getRangeは1インデックス
    const numRows = colors.length;
    const numCols = colors[0].length;
    const range = sheet.getRange(startRow, startCol, numRows, numCols);
  
    // 背景色を設定
    range.setBackgrounds(colors);
    
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
    //const scheduleBackGrounds = setHolidayBackgroundColors(dateArray); //祝日有

    //setSchedule(scheduleSheet,scheduleSheetDataValues)
    
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
  
  //カレンダー生成↑
    
  // スケジュール表情報を進行表の値で更新
  function updateDateValues(storyInfo,scheduleSheet) {
  const maxDays = getLastFilledColumnInCalender(scheduleSheet,SS_CALENDERDATE_COLUMN_INDEX);  //カレンダーの最大サイズを確認する。

    // 「以下フリースペース」が一列目に書いてある行を取得
    freeSpaceRowIndex = scheduleSheetDataValues.findIndex(
      (row) => row[0] == SS_FREE_SPACE
    );

    for (let storyInfoPerson of storyInfo.persons) {
      //console.log('storyInfoPerson='+storyInfoPerson.personName)
      if (storyInfoPerson.personName == "") continue;
      // 「以下フリースペース」より上かつ名前が一致、または他社_名前で一致する行の取得
      const rowIndex = scheduleSheetDataValues.findIndex((row, index) => {
        return (
          (row[SS_PRSRON_COLUMN_INDEX] == storyInfoPerson.personName ||
            OTHER_CO + row[SS_PRSRON_COLUMN_INDEX] ==
              storyInfoPerson.personName) &&
          index < freeSpaceRowIndex
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
    //console.log('---- updatePersonTasks IN ----')
    for (let scene of person.scenes) {
      if (!scene.replaceScene) {
        continue;
      }
      // 開始日の列番号を特定
      let startColumnIndex = undefined;
      let sceneStartDateStr = Utilities.formatDate(
        new Date(scene.startDate),
        "JST",
        "yyyy-MM-dd"
      );
      for (
        let i = SS_CALENDERDATE_COLUMN_INDEX;
        i < SS_CALENDERDATE_ROW_INDEX + maxDays;
        i++
      ) {
        let date = scheduleSheetDataValues[SS_CALENDERDATE_ROW_INDEX][i];
        let dateStr = Utilities.formatDate(date, "JST", "yyyy-MM-dd");
        if (sceneStartDateStr == dateStr) {
          startColumnIndex = i;
          break;
        }
      }
      // 開始日がスケジュール表にない場合はエラーのダイアログを出す
      if (startColumnIndex === undefined) {
        SpreadsheetApp.getUi().alert('開始日がスケジュール表に見つかりません: ' + sceneStartDateStr);
        continue; // 次のシーンへスキップ
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

      //先頭セルにシーン名を表示する
      //displaySceneNameC(rowIndex);
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
    let allDataRange = progressSheet.getDataRange();
    // 値を取得
    let progressSheetDataValues = allDataRange.getValues();
    // 背景色を取得する
    let progressSheetAllBackGrounds = allDataRange.getBackgrounds();
    // シーン名が何行までか確認する。
    let maxScenesRow = getPsSceneRowsIndex(progressSheet) + 1;

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
      i < PS_SCENENAME_START_ROW_INDEX + maxScenesRow;
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

  // 話数、シーン名チェック
  function checkSceneName(progressSheet) {
    let allDataRange = progressSheet.getDataRange();
    // 値を取得
    let progressValues = allDataRange.getValues();

    // エラー用変数
    let errorMessages = [];
    // 話数
    if (
      String(
        progressValues[PS_STORYNAME_ROW_INDEX][PS_STORYNAME_COLUMN_INDEX]
      ).indexOf(DELIMITER) > -1
    ) {
      errorMessages.push(`${ERROR_MESSAGE_DELIMITER} 話数`);
    }

    //シーン名 空データにあたるまで
    let i = PS_SCENENAME_START_ROW_INDEX;
    while (progressValues[i][PS_SCENENAME_COLUMN_INDEX] != "") {
      if (
        String(progressValues[i][PS_SCENENAME_COLUMN_INDEX]).indexOf(
          DELIMITER
        ) > -1
      ) {
        errorMessages.push(
          `${ERROR_MESSAGE_DELIMITER} ${i + 1}行目:${
            progressValues[i][PS_SCENENAME_COLUMN_INDEX]
          }`
        );
      }
      i++;
    }
    return errorMessages;
  }

  // 進行表のシーン名の行数(空データに当たるまで)を確認する。//空の行のindexを返す。
  function getPsSceneRowsIndex(sheet){
    let allDataRange = sheet.getDataRange();
    // 値を取得
    let progressValues = allDataRange.getValues();
    //シーン名 空データにあたるまで
    let rowIndex = PS_SCENENAME_START_ROW_INDEX;
    while (progressValues[rowIndex][PS_SCENENAME_COLUMN_INDEX] != "") {
      rowIndex++;
    }
    return rowIndex - 1;  //最後のシーン名が記入のある行を返す。
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