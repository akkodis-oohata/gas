//-----------------
//定数定義
//-----------------
const GMS_Const = {
  //共通
  //ステータス
  STATUS_MIIRI: 1,
  STATUS_GENZUIRI: 2,
  STATUS_DISPLAY_MIIRI: "未入り",
  STATUS_DISPLAY_GENZUIRI: "原図入り",

  //タスク割り当て状態
  NOT_ASSIGNED: 1, // スケジュール表にまだ割り当てていない
  ASSIGNED: 2, // スケジュール表に割り当て済み
  //Dateクラス用
  TIME_ZONE: "Asia/Tokyo",
  DATE_FORMAT: "yyyy/MM/dd",

  //スケジュール表
  //スケジュール表シート名
  SCHEDULE_SHEET_NAME: "スケジュール表",
  //表示最大日数
  MAXDAYS: 30, //ここ１年にすると時間かかるので30にしておく
  //タスク描画開始地点
  TASK_START_ROW_NUM: 2,
  TASK_START_COLUMN_NUM: 2,
  //カレンダー日付の場所
  CALENDERDATE_ROW_NUM: 1,
  CALENDERDATE_START_COLUMN_NUM: 2,
  //担当者開始地点
  PERSON_START_ROW_NUM: 2,
  PERSON_START_COLUMN_NUM: 1,

  //工数表
  //工数表シート名
  MANHOUR_SHEET_NAME: "工数表",
  //話名場所
  STORY_START_ROW_NUM: 1,
  STORY_COLUMN_NUM: 1,
  //データ表示領域
  //開始行
  DATA_START_ROW_NUM: 3,
  //シーン名
  SCEAN_NAME_COLUMN_NUM: 2,
  //担当者
  PERSON_NAME_COLUMN_NUM: 5,
  //開始日
  START_DATE_COLUMN_NUM: 6,
  //総工数
  TOTAL_COLUMN_NUM: 7,
  //原図入り工数
  GENZUIRI_COLUMN_NUM: 8,
  //話数セパレーター
  STORY_SEPARATOR: "###",

  //進行表
  //進行表シート名
  PROGRESS_SHEET_NAME: "進行表",

  //全工数表
  //全工数表シート名
  MANHOUR_ALL_SHEET_NAME: "全工数表",
  //タスク描画開始地点
  MANHOUR_START_ROW_NUM: 2,
  MANHOUR_START_COLUMN_NUM: 2,

  COLOR_CLEAR: "#ffffff", //セルの色塗り初期化時の色
  COLOR_TOTAL_BACKGROUND: "#d9d2e9", //合計のセルの色
  COLOR_MIIRIMIMAKI_TOTAL_BACKGROUND: "#9fc5e8", //未入・未撒き合計（未作業）の背景色
  COLOR_LO_STATUS_LOZOROI: "#d9ead3", //未入りが0枚の状態に塗りつぶす色
  COLOR_LO_STATUS_MAKIZUMI: "#b7b7b7", //LO揃＆未撒きが0枚の状態のときに塗りつぶす色
  COLOR_LO_STATUS_T1: "#999999", //納品Cut数と総工数が同じだったときに塗りつぶす色
  TEXT_COLOR_NORMAL: "#000000", //通常時の文字色
  TEXT_COLOR_LO_STATUS_LOZOROI: "#000000", //未入りが0枚の状態に塗りつぶす文字色
  TEXT_COLOR_LO_STATUS_MAKIZUMI: "#cc0000", //LO揃＆未撒きが0枚の状態のときに塗りつぶす文字色
  TEXT_COLOR_LO_STATUS_T1: "#0000ff", //納品Cut数と総工数が同じだったときに塗りつぶす文字色
};

//-----------------
//クラス
//-----------------
//担当者情報
class Person {
  constructor(name) {
    //名前
    this.name = name;
    //タスク（各ステータスごとに一要素で格納する）
    this.tasks = [];
  }
}

//タスク情報
//各担当の各シーンごとに１データできる
class Task {
  constructor(storyName, scean, status, color, manHour, startDate) {
    //話名
    this.storyName = storyName;
    //シーン名
    this.scean = scean;
    //ステータス
    this.status = status;
    //ステータスに応じた色
    this.color = color;
    //工数
    this.manHour = manHour;
    //開始日
    this.startDate = startDate;
    //割り当て状態
    this.assign = GMS_Const.NOT_ASSIGNED;
  }
}

//（工数表での）話数情報
class StoryInfo {
  constructor(storyRowNum, startDataRowNum) {
    //話タイトル行番号
    this.storyRowNum = storyRowNum;
    //入力されたデータが始まる行番号
    this.startDataRowNum = startDataRowNum;
    //入力されたデータが終わる行番号（後で更新する）
    this.endDataRowNum = startDataRowNum;
    //話を担当する各担当者の情報
    this.persons = [];
  }
}

//-----------------
//グローバル変数
//-----------------
//各話ごとの情報を格納する配列
let storyInfoArray = [];
//スケジュール表の最小開始日（全話通しての）
let minDate;

//-----------------
//関数
//-----------------
//メインの関数
function generateManHoursAllSheet() {
  // スプレッドシートの読み込み
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // 進行表読み込み
  const progress_sheets = spreadsheet
    .getSheets()
    .filter(
      (sheet) => sheet.getName().indexOf(GMS_Const.PROGRESS_SHEET_NAME) > -1
    );

  // 全工数表読み込み
  let manhour_all_sheet = spreadsheet.getSheetByName(
    GMS_Const.MANHOUR_ALL_SHEET_NAME
  );

  // 工数設定
  setmanhour(progress_sheets, manhour_all_sheet);
}

//-----------------
// 工数表解析関数
//-----------------
// 話数ごとの情報を配列に詰める
function setmanhour(progress_sheets, manhour_all_sheet) {
  const endRow = manhour_all_sheet.getLastRow();
  manhour_all_sheet
    .getRange(
      GMS_Const.SCENE_ALL_SCENE_TOTAL_ROW_NUM,
      GMS_Const.SCENE_LO_STATUS_COLUMN_NUM,
      1,
      GMS_Const.SCENE_MAKIZUMI_NUMBER_COLUMN_NUM
    )
    .clearContent()
    .setBackground(GMS_Const.COLOR_CLEAR);
  // 直前の話数の情報
  let preStoryInfo = undefined;
  //★ここを2固定でなく、何話でもいけるようにする
  for (let z = 0; z < 2; z++) {
    let storyRowNum = preStoryInfo
      ? preStoryInfo.endDataRowNum + 2
      : GMS_Const.STORY_START_ROW_NUM;
    let startDataRowNum = preStoryInfo
      ? preStoryInfo.endDataRowNum + 4
      : GMS_Const.DATA_START_ROW_NUM;

    let storyInfo = new StoryInfo(storyRowNum, startDataRowNum);
    let storyName = getStoryName(sheet, storyInfo);
    let statusInfo = getStatusInfo(sheet, storyInfo);
    storyInfo.endDataRowNum = getEndDataRowNum(sheet, storyInfo);
    //開始行番号から終了行番号までの情報を 作業者 - タスク　の構成に入れていく
    for (let i = storyInfo.startDataRowNum; i <= storyInfo.endDataRowNum; i++) {
      let sceanName = sheet
        .getRange(i, GMS_Const.SCEAN_NAME_COLUMN_NUM)
        .getValue();
      let personName = sheet
        .getRange(i, GMS_Const.PERSON_NAME_COLUMN_NUM)
        .getValue();
      let startDate = sheet
        .getRange(i, GMS_Const.START_DATE_COLUMN_NUM)
        .getValue();
      let total = sheet.getRange(i, GMS_Const.TOTAL_COLUMN_NUM).getValue();
      let genzuiri = sheet
        .getRange(i, GMS_Const.GENZUIRI_COLUMN_NUM)
        .getValue();
      let miiri = total - genzuiri;

      //タスクの箱作る
      let tasks = [];
      //未入りに工数指定あり
      if (miiri > 0) {
        let color = undefined;
        statusInfo.forEach((status) => {
          if (status["status"] == GMS_Const.STATUS_DISPLAY_MIIRI) {
            color = status["color"];
          }
        });
        tasks = addTasks(
          tasks,
          storyName,
          sceanName,
          GMS_Const.STATUS_MIIRI,
          color,
          miiri,
          startDate
        );
      }
      //原図入りに工数指定あり
      if (genzuiri > 0) {
        let color = undefined;
        statusInfo.forEach((status) => {
          if (status["status"] == GMS_Const.STATUS_DISPLAY_GENZUIRI) {
            color = status["color"];
          }
        });
        tasks = addTasks(
          tasks,
          storyName,
          sceanName,
          GMS_Const.GENZUIRI,
          color,
          genzuiri,
          startDate
        );
      }
      //対象となる作業者を選定 or 新しく作成し、タスクを追加
      let personFound = false;
      for (let person of storyInfo.persons) {
        if (person.name == personName) {
          tasks.forEach((task) => {
            person.tasks.push(task);
          });
          personFound = true;
          break;
        }
      }
      if (!personFound) {
        let newPerson = new Person(personName);
        tasks.forEach((task) => {
          newPerson.tasks.push(task);
        });
        storyInfo.persons.push(newPerson);
      }
    }
    //話数 - 担当者 - タスクの設定を追加
    storyInfoArray.push(storyInfo);
    //次の話数解析用に持っておく
    preStoryInfo = storyInfo;
  }
}
