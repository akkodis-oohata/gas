function updateScheduleSheetMain() {
  updateScheduleSheet();
}

function generateCalendarMain() {
  generateCalendar();
}

function updateScheduleSheetDeadline() {
  updateDeadline();
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
      //作品話数文字色
      this.storyFontColor = undefined;
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
  //セルの色塗り初期化時の色
  const COLOR_CLEAR = "#ffffff";
  // 土日祝日の色
  const COLOR_HOLIDAY = "#808080";
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

  //MAX担当者数（便宜的に用意）
  const PS_MAX_SECNES = 31;

  // 区切り文字
  const DELIMITER = ",";
  //-----------------
  //定数定義（スケジュール表）
  //-----------------
  //カレンダー生成日付
  const SS_INPUTCALENDERDATE_ROW_INDEX = 2;
  const SS_INPUTCALENDERDATE_COLUMN_INDEX = 4;

  //カレンダー
  const SS_CALENDERDATE_ROW_INDEX = 6;
  const SS_CALENDERDATE_COLUMN_INDEX = 7;
  //表示最大日数
  const SS_MAXDAYS = 730; //730(24か月)
  //担当者
  const SS_PRSRON_COLUMN_INDEX = 6;
  const SS_PERSON_ROW_START_INDEX = 7;
  const SS_PERSON_ROW_STEP = 3; //作業者間の行間、作業者名、ステータス行、memo行
  //MAX担当者数（便宜的に用意）
  const SS_MAX_PERSONS = 8;
  //HIDDENデータ領域（処理の都合上全セルに話数名#シーン名を持たしておきたいので）
  const SS_HIDDEN_ROW_START_INDEX = 41;
  //カレンダー生成時前にクリアする領域（適当な値）//TODO
  const SS_CLEAR_ROW_LENGTH = 54;
  // データベースとスケジュール表のフォーマットの差を埋める為、美術ルーム全体スケジュールmemoの行数分+CLW美術作業者一覧分-以下データスペース-作業者ベースデータで算出する
  const DATA_BASE_FORMAT_OFFSET = 5;

  //エラーメッセージ
  const ERROR_MESSAGE_FULL = `メモ欄が6行全て埋まっています`;
  const ERROR_MESSAGE_DATE_MISMAYCH = `既に締切が設定されていますが、日付が不一致です`;
  const ERROR_MESSAGE_OUT_OF_DATE = `日付がスケジュールの範囲にありません`;
  const ERROR_MESSAGE_DELIMITER = `値に区切り文字 ${DELIMITER} が含まれています:`;
  // 他社接頭語
  const OTHER_CO = "他社_";
  // フリースペース開始位置
  const SS_FREE_SPACE = "以下フリースぺース";

  const DATA_BASE_SHEET_NAME = "データベース";

  //-----------------
  //変数
  //-----------------
  // スプレッドシート上でデータが入力されている最大範囲を選択
  let scheduleSheetAllRange = undefined;
  //スケジュール表の全値取得
  let scheduleSheetDataValues = undefined;
  //スケジュール表の背景色取得
  let scheduleSheetAllBackGrounds = undefined;
  // データベースシートの最大範囲
  let dataBaseSheetAllRange = undefined;
  // データベースシートに格納する為、データ領域を格納する配列
  let dataBaseSheetDataValues = undefined;
  // 「以下フリースペース」が一列目に書いてある行
  let freeSpaceRowIndex = undefined;

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
    updateDataSpaceMain(scheduleSheet);

    // スケジュール表情報取得(データスペース反映後)
    getScheduleSheetInfo(scheduleSheet, dataBaseSheet);

    // スケジュール表情報を進行表の値で更新
    updateDateValues(storyInfo);
    // 更新したスケジュール表情報で画面更新
    updateScheduleSheetWithDataValues();

    // スケジュール表をActiveにする
    scheduleSheet.activate();

    console.log("---- updateScheduleSheet OUT ----");
    console.timeEnd(label);
  }
  //スコープに公開
  this.updateScheduleSheet = updateScheduleSheet;

  //カレンダー生成↓
  //カレンダー生成（事前にやる）
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
    getScheduleSheetInfo(scheduleSheet, dataBaseSheet);

    //カレンダー日付設定
    setCallendarDate();
    console.timeEnd(label);
  }

  //カレンダー日付設定
  function setCallendarDate() {
    //領域範囲取得
    const clearRowIndex =
      SS_CALENDERDATE_ROW_INDEX + SS_CLEAR_ROW_LENGTH <
      scheduleSheetDataValues.length
        ? SS_CALENDERDATE_ROW_INDEX + SS_CLEAR_ROW_LENGTH
        : scheduleSheetDataValues.length;
    const clearColumIndex =
      SS_PRSRON_COLUMN_INDEX + SS_MAXDAYS + 1 <
      scheduleSheetDataValues[0].length
        ? SS_PRSRON_COLUMN_INDEX + SS_MAXDAYS + 1
        : scheduleSheetDataValues[0].length;

    //領域クリア（カレンダー、担当者名、作業、HIDDEN領域）
    for (let i = SS_CALENDERDATE_ROW_INDEX; i < clearRowIndex; i++) {
      for (let j = SS_PRSRON_COLUMN_INDEX; j < clearColumIndex; j++) {
        scheduleSheetDataValues[i][j] = "";
        scheduleSheetAllBackGrounds[i][j] = COLOR_CLEAR;
      }
    }

    let dateArray = [];
    //開始日取得
    let minDateObj = new Date(
      scheduleSheetDataValues[SS_INPUTCALENDERDATE_ROW_INDEX][
        SS_INPUTCALENDERDATE_COLUMN_INDEX
      ]
    );

    for (let i = 0; i < SS_MAXDAYS; i++) {
      let dateObj = new Date();
      dateObj.setMonth(new Date(minDateObj.getMonth()));
      dateObj.setDate(new Date(minDateObj.getDate() + i));
      dateArray.push(dateObj);
    }
    for (
      let j = SS_CALENDERDATE_COLUMN_INDEX;
      j < SS_CALENDERDATE_COLUMN_INDEX + SS_MAXDAYS;
      j++
    ) {
      let dateArrayJ = j - SS_CALENDERDATE_COLUMN_INDEX;
      scheduleSheetDataValues[SS_CALENDERDATE_ROW_INDEX][j] =
        dateArray[dateArrayJ];
    }
    let dateColorArray = [];
    for (let date of dateArray) {
      let dateColor = isHoliday(date) ? COLOR_HOLIDAY : COLOR_CLEAR;
      dateColorArray.push(dateColor);
    }

    //領域に土日祝日設定
    let cellColorArray = [];
    for (let i = 0; i < SS_MAX_PERSONS * 2 + 1; i++) {
      cellColorArray.push(dateColorArray);
    }
    for (
      let i = SS_CALENDERDATE_ROW_INDEX;
      i < SS_CALENDERDATE_ROW_INDEX + SS_MAX_PERSONS * 2 + 1;
      i++
    ) {
      for (
        let j = SS_CALENDERDATE_COLUMN_INDEX;
        j < SS_CALENDERDATE_COLUMN_INDEX + SS_MAXDAYS;
        j++
      ) {
        let cellColorArrayI = i - SS_CALENDERDATE_ROW_INDEX;
        let cellColorArrayJ = j - SS_CALENDERDATE_COLUMN_INDEX;
        scheduleSheetAllBackGrounds[i][j] =
          cellColorArray[cellColorArrayI][cellColorArrayJ];
      }
    }

    updateScheduleSheetWithDataValues();
  }
  //カレンダー生成↑

  //スケジュール表の画面を更新する
  function updateScheduleSheetWithDataValues() {
    scheduleSheetAllRange.setValues(scheduleSheetDataValues);
    scheduleSheetAllRange.setBackgrounds(scheduleSheetAllBackGrounds);
    dataBaseSheetAllRange.setValues(dataBaseSheetDataValues);
  }
  // スケジュール表情報を進行表の値で更新
  function updateDateValues(storyInfo) {
    // 「以下フリースペース」が一列目に書いてある行を取得
    freeSpaceRowIndex = scheduleSheetDataValues.findIndex(
      (row) => row[0] == SS_FREE_SPACE
    );

    //console.log('---- updateDateValues IN ----')

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
        deletePersonTasks(storyInfo, storyInfoPerson, rowIndex);
        updatePersonTasks(storyInfo, storyInfoPerson, rowIndex);
      }
    }
    //console.log('---- updateDateValues OUT ----')
  }

  // 対象の担当者のシーン情報を削除する
  function deletePersonTasks(storyInfo, person, rowIndex) {
    //console.log('---- deletePersonTasks IN ----')
    for (let scene of person.scenes) {
      if (!scene.replaceScene) {
        continue;
      }
      let sceneTitle = storyInfo.storyName + DELIMITER + scene.sceneName;
      for (
        let i = SS_CALENDERDATE_COLUMN_INDEX;
        i < SS_CALENDERDATE_COLUMN_INDEX + SS_MAXDAYS;
        i++
      ) {
        let tmpSceneTitle =
          dataBaseSheetDataValues[rowIndex - DATA_BASE_FORMAT_OFFSET][i]; //TODO:dataBaseとスケジュール表はデータ行が合わないので調整要
        if (tmpSceneTitle && tmpSceneTitle.startsWith(sceneTitle)) {
          // startsWithメソッドを使用してチェック
          // セルの内容をクリア
          scheduleSheetDataValues[rowIndex][i] = "";
          scheduleSheetDataValues[rowIndex + 1][i] = "";
          dataBaseSheetDataValues[rowIndex - DATA_BASE_FORMAT_OFFSET][i] = "";
          dataBaseSheetDataValues[rowIndex + 1 - DATA_BASE_FORMAT_OFFSET][i] =
            "";
          // セルの背景色をクリア
          scheduleSheetAllBackGrounds[rowIndex][i] = COLOR_CLEAR;
          scheduleSheetAllBackGrounds[rowIndex + 1][i] = COLOR_CLEAR;
        }
      }
    }
  }

  // 対象の担当者のシーン情報を更新する
  function updatePersonTasks(storyInfo, person, rowIndex) {
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
        i < SS_CALENDERDATE_ROW_INDEX + SS_MAXDAYS;
        i++
      ) {
        let date = scheduleSheetDataValues[SS_CALENDERDATE_ROW_INDEX][i];
        let dateStr = Utilities.formatDate(date, "JST", "yyyy-MM-dd");
        if (sceneStartDateStr == dateStr) {
          startColumnIndex = i;
          break;
        }
      }

      // 開始日から総工数（日）だけセルに色を塗る
      if (scene.totalDays > 0) {
        let sceneTitle = storyInfo.storyName + DELIMITER + scene.sceneName;
        for (let i = 0; i < scene.totalDays; i++) {
          let checkColumnIndex = startColumnIndex + i;
          if (scene.warikomi) {
            moveCell(
              rowIndex,
              checkColumnIndex,
              storyInfo.storyColor,
              sceneTitle,
              scene.manHoursColor[i]
            );
          } else {
            while (
              !findAndFillClearCell(
                rowIndex,
                checkColumnIndex,
                storyInfo.storyColor,
                sceneTitle,
                scene.manHoursColor[i]
              )
            ) {
              checkColumnIndex++;
            }
          }
        }
      }

      //先頭セルにシーン名を表示する
      displaySceneName(rowIndex);
    }
  }

  //先頭セルにシーン名を表示する
  function displaySceneName(rowIndex) {
    let prevSceneTitle = "";
    for (
      let columnIndex = SS_CALENDERDATE_COLUMN_INDEX;
      columnIndex < SS_CALENDERDATE_COLUMN_INDEX + SS_MAXDAYS;
      columnIndex++
    ) {
      let sceneTitle =
        dataBaseSheetDataValues[rowIndex - DATA_BASE_FORMAT_OFFSET][
          columnIndex
        ];
      if (sceneTitle != prevSceneTitle) {
        scheduleSheetDataValues[rowIndex][columnIndex] = sceneTitle;
      }
      prevSceneTitle = sceneTitle;
    }
  }

  //セルを隣に移動 //TODO:共通の切り出しとしてつかるか検討
  function moveCell(rowIndex, index, existColor, sceneTitle, manHourColor) {
    let nextExistColor = scheduleSheetAllBackGrounds[rowIndex][index];
    //土日祝日だけ割り込み禁止
    if (nextExistColor == COLOR_HOLIDAY) {
      index++;
      moveCell(rowIndex, index, existColor, sceneTitle, manHourColor);
      return;
    }
    let nextSceneTitle = undefined;
    let nextDataBaseSceneTitle = undefined;
    if (nextExistColor != COLOR_CLEAR) {
      nextSceneTitle = scheduleSheetDataValues[rowIndex][index];

      nextDataBaseSceneTitle = scheduleSheetDataValues[rowIndex][index];
    }
    setSell(rowIndex, index, existColor, sceneTitle, manHourColor);
    if (nextExistColor != COLOR_CLEAR) {
      index++;
      if (nextSceneTitle != undefined && nextSceneTitle != "") {
        moveCell(rowIndex, index, nextExistColor, nextSceneTitle, manHourColor);
      } else {
        moveCell(
          rowIndex,
          index,
          nextExistColor,
          nextDataBaseSceneTitle,
          manHourColor
        );
      }
    }
  }

  // 空白のセルを探して色を塗る
  function findAndFillClearCell(
    rowIndex,
    checkColumnIndex,
    storyInfoColor,
    sceneTitle,
    manHourColor
  ) {
    let existColor = scheduleSheetAllBackGrounds[rowIndex][checkColumnIndex];
    if (existColor != COLOR_CLEAR) {
      return false;
    } else {
      setSell(
        rowIndex,
        checkColumnIndex,
        storyInfoColor,
        sceneTitle,
        manHourColor
      );
      return true;
    }
  }

  //セルに値を設定。表示用シーン名はクリア（後で一気に付ける）
  function setSell(rowIndex, columnIndex, setColor, sceneTitle, manHourColor) {
    scheduleSheetDataValues[rowIndex][columnIndex] = "";
    scheduleSheetAllBackGrounds[rowIndex][columnIndex] = setColor;
    scheduleSheetAllBackGrounds[rowIndex + 1][columnIndex] = manHourColor;
    dataBaseSheetDataValues[rowIndex - DATA_BASE_FORMAT_OFFSET][columnIndex] =
      sceneTitle;
  }

  // スケジュール表情報取得
  function getScheduleSheetInfo(scheduleSheet, dataBaseSheet) {
    // スプレッドシート上でデータが入力されている最大範囲を選択
    scheduleSheetAllRange = scheduleSheet.getDataRange();
    // 値を取得
    scheduleSheetDataValues = scheduleSheetAllRange.getValues();
    // 背景色を取得する
    scheduleSheetAllBackGrounds = scheduleSheetAllRange.getBackgrounds();

    // scheduleSheetAllRange の範囲情報を取得
    var scheduleNumRows = scheduleSheetAllRange.getNumRows();
    var scheduleNumColumns = scheduleSheetAllRange.getNumColumns();

    // dataBaseSheet の範囲を scheduleSheetAllRange の範囲で取得
    dataBaseSheetAllRange = dataBaseSheet.getRange(
      1,
      1,
      scheduleNumRows,
      scheduleNumColumns
    );
    // 値を取得
    dataBaseSheetDataValues = dataBaseSheetAllRange.getValues();
  }

  //土日祝日判定
  function isHoliday(date) {
    //土日の判定
    const day = date.getDay();
    if (day === 0 || day === 6) return true;
    //祝日の取得。ここを使うにはGoogleDrive上で一度「カレンダー」にアクセスする必要あり
    const id = "ja.japanese#holiday@group.v.calendar.google.com";
    const cal = CalendarApp.getCalendarById(id);
    const events = cal.getEventsForDay(date);
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
    storyInfo.storyFontColor = allDataRange.getFontColorObject();

    //シーン情報・担当者情報
    for (
      let i = PS_SCENENAME_START_ROW_INDEX;
      i < PS_SCENENAME_START_ROW_INDEX + PS_MAX_SECNES;
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
      scene.totalDays =
        progressSheetDataValues[i][PS_TOTALDAYS_COLUMN_INDEX] +
        progressSheetDataValues[i][PS_SCENE_REARANGE_MANHOUR_COLUMN_INDEX]; //総工数＋再調整の値にする
      scene.warikomi = progressSheetDataValues[i][PS_WARIKOMI_COLUMN_INDEX];
      scene.replaceScene =
        progressSheetDataValues[i][PS_REPLACESCENE_COLUMN_INDEX];
      // 工数背景色
      const manHoursEndIndex = progressSheetDataValues[i].findIndex(
        (data, index) => data != "" && index >= PS_SCENE_MANHOUR_COLUMN_INDEX
      ); //工数背景色が何列目までにあるか取得する。

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
    getScheduleSheetInfo(scheduleSheet, dataBaseSheet);

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

    updateScheduleSheetWithDataValues();

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
