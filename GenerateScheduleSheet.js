//-----------------
//定数定義
//-----------------
const GS_Const = {
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
    this.assign = GS_Const.NOT_ASSIGNED;
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
function generateScheduleSheet() {
  // スプレッドシートの読み込み
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // 工数表読み込み
  let sheet = spreadsheet.getSheetByName(GS_Const.MANHOUR_SHEET_NAME);

  // 話数ごとの情報を配列に詰める
  setStoryInfoArray(sheet);

  // すでにあるスケジュール表削除
  let delsheet = spreadsheet.getSheetByName(GS_Const.SCHEDULE_SHEET_NAME);
  spreadsheet.deleteSheet(delsheet);

  // 進行表作成
  let scheduleSheet = spreadsheet.insertSheet();
  scheduleSheet.setName(GS_Const.SCHEDULE_SHEET_NAME);
  // カレンダーの日付設定
  setScheduleDate(scheduleSheet);
  // 工数設定
  setTasksOnSchedule(scheduleSheet);
}

//-----------------
// 工数表解析関数
//-----------------
// 話数ごとの情報を配列に詰める
function setStoryInfoArray(sheet) {
  // 直前の話数の情報
  let preStoryInfo = undefined;
  //★ここを2固定でなく、何話でもいけるようにする
  for (let z = 0; z < 2; z++) {
    let storyRowNum = preStoryInfo
      ? preStoryInfo.endDataRowNum + 2
      : GS_Const.STORY_START_ROW_NUM;
    let startDataRowNum = preStoryInfo
      ? preStoryInfo.endDataRowNum + 4
      : GS_Const.DATA_START_ROW_NUM;

    let storyInfo = new StoryInfo(storyRowNum, startDataRowNum);
    let storyName = getStoryName(sheet, storyInfo);
    let statusInfo = getStatusInfo(sheet, storyInfo);
    storyInfo.endDataRowNum = getEndDataRowNum(sheet, storyInfo);
    //開始行番号から終了行番号までの情報を 作業者 - タスク　の構成に入れていく
    for (let i = storyInfo.startDataRowNum; i <= storyInfo.endDataRowNum; i++) {
      let sceanName = sheet
        .getRange(i, GS_Const.SCEAN_NAME_COLUMN_NUM)
        .getValue();
      let personName = sheet
        .getRange(i, GS_Const.PERSON_NAME_COLUMN_NUM)
        .getValue();
      let startDate = sheet
        .getRange(i, GS_Const.START_DATE_COLUMN_NUM)
        .getValue();
      let total = sheet.getRange(i, GS_Const.TOTAL_COLUMN_NUM).getValue();
      let genzuiri = sheet.getRange(i, GS_Const.GENZUIRI_COLUMN_NUM).getValue();
      let miiri = total - genzuiri;

      //タスクの箱作る
      let tasks = [];
      //未入りに工数指定あり
      if (miiri > 0) {
        let color = undefined;
        statusInfo.forEach((status) => {
          if (status["status"] == GS_Const.STATUS_DISPLAY_MIIRI) {
            color = status["color"];
          }
        });
        tasks = addTasks(
          tasks,
          storyName,
          sceanName,
          GS_Const.STATUS_MIIRI,
          color,
          miiri,
          startDate
        );
      }
      //原図入りに工数指定あり
      if (genzuiri > 0) {
        let color = undefined;
        statusInfo.forEach((status) => {
          if (status["status"] == GS_Const.STATUS_DISPLAY_GENZUIRI) {
            color = status["color"];
          }
        });
        tasks = addTasks(
          tasks,
          storyName,
          sceanName,
          GS_Const.GENZUIRI,
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

//タスク追加用
function addTasks(
  tasks,
  storyName,
  sceanName,
  status,
  color,
  manHour,
  startDate
) {
  tasks.push(new Task(storyName, sceanName, status, color, manHour, startDate));
  return tasks;
}

//話名取得
function getStoryName(sheet, storyInfo) {
  let range = sheet.getRange(storyInfo.storyRowNum, GS_Const.STORY_COLUMN_NUM);
  return range.getValue();
}

//ステータス - 色の組み合わせ取得
function getStatusInfo(sheet, storyInfo) {
  let statusInfo = [];

  for (
    let i = storyInfo.startDataRowNum;
    i <= storyInfo.startDataRowNum + 1;
    i++
  ) {
    let range = sheet.getRange(i, 1);
    statusInfo.push({ status: range.getValue(), color: range.getBackground() });
  }

  return statusInfo;
}

// 指定した話数の入力最後の行数
function getEndDataRowNum(sheet, storyInfo) {
  let currentNum = storyInfo.startDataRowNum;
  //必ずセパレーターはある想定
  while (true) {
    currentNum++;
    let range = sheet.getRange(currentNum, 1);
    let strVal = range.getValue();
    if (strVal == GS_Const.STORY_SEPARATOR) {
      break;
    }
  }
  return --currentNum;
}

//-----------------
// スケジュール表生成用関数
//-----------------
//スケジュール表に日付設定（とりあえず開始日から30日分を振る）
function setScheduleDate(sheet) {
  //for(var i=0;i<storyInfoArray.length;i++) {
  for (let storyInfo of storyInfoArray) {
    storyInfo.persons.forEach((person) => {
      person.tasks.forEach((task) => {
        if (minDate === undefined) {
          minDate = task.startDate;
        } else {
          if (minDate > task.startDate) {
            minDate = task.taskStartDate;
          }
        }
      });
    });
  }
  let range = sheet.getRange(
    GS_Const.CALENDERDATE_ROW_NUM,
    GS_Const.CALENDERDATE_START_COLUMN_NUM
  );
  range.setValue(minDate);
  let minDateObj = new Date(minDate);

  for (let i = 0; i < GS_Const.MAXDAYS; i++) {
    let nextRange = sheet.getRange(
      GS_Const.CALENDERDATE_ROW_NUM,
      GS_Const.CALENDERDATE_START_COLUMN_NUM + 1 + i
    );
    let dateObj = new Date();
    dateObj.setMonth(new Date(minDateObj.getMonth()));
    dateObj.setDate(new Date(minDateObj.getDate() + (i + 1)));
    nextRange.setValue(dateObj);
  }
}

//スケジュール表に人とタスクを設定
function setTasksOnSchedule(sheet) {
  //人のカウント
  let count = 0;
  //最小開始日
  let minDateObj = new Date(minDate);

  //話数でループしーて全担当者情報のタスクをまとめる
  let allPersons = [];
  for (let storyInfo of storyInfoArray) {
    for (let person of storyInfo.persons) {
      let found = false;
      for (let tmpAllPerson of allPersons) {
        if (tmpAllPerson.name == person.name) {
          for (let task of person.tasks) {
            tmpAllPerson.tasks.push(task);
          }
          found = true;
        }
      }
      if (!found) {
        allPersons.push(person);
      }
    }
  }

  allPersons.forEach((person) => {
    //担当者名
    let range = sheet.getRange(
      GS_Const.PERSON_START_ROW_NUM + count,
      GS_Const.PERSON_START_COLUMN_NUM
    );
    range.setValue(person.name);

    //開始日付から最大日数まで順にチェックしていく
    for (let i = 0; i < GS_Const.MAXDAYS; i++) {
      let today = new Date();
      let targetDate = new Date(
        today.getFullYear(),
        today.getMonth(),
        today.getDate()
      );
      targetDate.setMonth(minDateObj.getMonth());
      targetDate.setDate(minDateObj.getDate() + i);
      let fromatTargetDate = Utilities.formatDate(
        targetDate,
        GS_Const.TIME_ZONE,
        GS_Const.DATE_FORMAT
      );
      let drawTarget = undefined;

      //全タスクをチェック
      person.tasks.forEach((task) => {
        let fromatStartDate = Utilities.formatDate(
          task.startDate,
          GS_Const.TIME_ZONE,
          GS_Const.DATE_FORMAT
        );
        if (
          task.assign == GS_Const.NOT_ASSIGNED &&
          fromatStartDate <= fromatTargetDate
        ) {
          if (drawTarget) {
            let formatDrawTargetStartDate = Utilities.formatDate(
              drawTarget.startDate,
              GS_Const.TIME_ZONE,
              GS_Const.DATE_FORMAT
            );

            if (drawTarget.status < task.status) {
              drawTarget = task;
            } else if (formatDrawTargetStartDate > fromatStartDate) {
              drawTarget = task;
            }
          } else {
            drawTarget = task;
          }
        }
      });

      //描画対象タスクあり
      if (drawTarget) {
        console.log(targetDate);
        console.log(minDateObj);
        let startColumn = (targetDate - minDateObj) / 86400000;
        //var difference = targetDate.getDate().diff(minDateObj.getDate(),"days");
        //const diff = dayjs.dayjs(targetDate).diff(minDateObj, 'day');
        //console.log('diff='+diff)

        //描画
        //getRange(row, column, numRows, numColumns)
        let rowNum = GS_Const.TASK_START_ROW_NUM + count;
        let columnNum = GS_Const.TASK_START_COLUMN_NUM + startColumn;
        let numRows = 1;
        let numColumns = drawTarget.manHour;
        console.log(
          "rowNum=" +
            rowNum +
            ",columnNum=" +
            columnNum +
            ",numRows=" +
            numRows +
            ",numColumns" +
            numColumns
        );
        let taskRange = sheet.getRange(rowNum, columnNum, numRows, numColumns);
        taskRange.setBackground(drawTarget.color);
        let headRange = sheet.getRange(rowNum, columnNum);
        headRange.setValue(drawTarget.storyName + ":" + drawTarget.scean);

        //割り当てを済にする
        drawTarget.assign = GS_Const.ASSIGNED;

        //日数を進める
        i = i + drawTarget.manHour - 1;
      }
    }

    count++;
  });
}
