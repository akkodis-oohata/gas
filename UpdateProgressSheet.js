//-----------------
//定数定義
//-----------------
const UP_Const = {
  //進行表シート名
  PROGRESS_SHEET_NAME: "進行表テストデータ",

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

  EXCLUSION_KEYWORD: "欠番、Bank、全セル",
  MIMAKI_KEYWORD: "原図入稿済",
  BG_AGARI_KEYWORD: "納品済",
  MAKIZUMI_KEYWORD: "撒済",
  COLOR_SETTING_KEYWORDS: ["原図入稿済", "納品済", "撒済"],

  T1_KEYWORD: "T1済",
  LOZOROI_KEYWORD: "LO揃",

  //シーン情報
  SCENE_TOTAL_LO_ROW_NUM: 3,
  SCENE_HEADER_ROW_NUM: 4,
  SCENE_HEADER_ROW_INDEX: 4 - 1,
  SCENE_ALL_SCENE_TOTAL_ROW_NUM: 6,
  SCENE_START_BODY_ROW_NUM: 7,
  SCENE_START_BODY_ROW_INDEX: 7 - 1,

  SCENE_COLUMN_NUM: 2,
  SCENE_COLUMN_INDEX: 2 - 1,
  SCENE_LO_STATUS_COLUMN_NUM: 3,
  SCENE_AVERAGE_CUT_COLUMN_NUM: 4,
  SCENE_AVERAGE_CUT_COLUMN_INDEX: 4 - 1,

  SCENE_CUT_NUMBER_COLUMN_NUM: 5,
  SCENE_CUT_NUMBER_COLUMN_INDEX: 5 - 1,
  SCENE_BG_COLUMN_NUM: 6,
  SCENE_BG_COLUMN_INDEX: 6 - 1,
  SCENE_TOTAL_MANHOUR_COLUMN_NUM: 7,
  SCENE_TOTAL_MANHOUR_COLUMN_INDEX: 7 - 1,
  SCENE_DRAW_COLUMN_NUM: 8,
  SCENE_DRAW_COLUMN_INDEX: 8 - 1,
  SCENE_MIIRI_MANHOUR_COLUMN_NUM: 9,
  SCENE_MIIRI_MANHOUR_COLUMN_INDEX: 9 - 1,
  SCENE_MIIRI_NUMBER_COLUMN_NUM: 10,
  SCENE_MIIRI_NUMBER_COLUMN_INDEX: 10 - 1,

  SCENE_MIMAKI_MANHOUR_COLUMN_NUM: 11,
  SCENE_MIMAKI_MANHOUR_COLUMN_INDEX: 11 - 1,
  SCENE_MIMAKI_NUMBER_COLUMN_NUM: 12,
  SCENE_MIMAKI_NUMBER_COLUMN_INDEX: 12 - 1,
  SCENE_MAKIZUMI_MANHOUR_COLUMN_NUM: 13,
  SCENE_MAKIZUMI_MANHOUR_COLUMN_INDEX: 13 - 1,
  SCENE_MAKIZUMI_NUMBER_COLUMN_NUM: 14,
  SCENE_MAKIZUMI_NUMBER_COLUMN_INDEX: 14 - 1,

  SCENE_REARANGE_MANHOUR_COLUMN_NUM: 15,
  SCENE_REARANGE_MANHOUR_COLUMN_INDEX: 15 - 1,
  SCENE_PERSON_COLUMN_NUM: 16,
  SCENE_PERSON_COLUMN_INDEX: 16 - 1,
  SCENE_STARTDATE_COLUMN_NUM: 17,
  SCENE_STARTDATE_COLUMN_INDEX: 17 - 1,
  SCENE_START_MANHOUR_COLUMN_NUM: 21,
  SCENE_END_MANHOUR_COLUMN_NUM: 170,

  MAX_TOTAL_MANDAY: 150,

  SCENE_TOTAL_MINOU_NUMBER_COLUMN_NUM: 6,
  SCENE_TOTAL_LOMIIRI_NUMBER_COLUMN_NUM: 9,
  SCENE_TOTAL_LOMIIRI_PERCENTAGE_COLUMN_NUM: 10,
  SCENE_TOTAL_LO_NUMBER_COLUMN_NUM: 13,
  SCENE_TOTAL_LO_PERCENTAGE_COLUMN_NUM: 14,
  SCENE_TOTAL_MIIRIMIMAKI_ROW_NUM: 2,
  SCENE_TOTAL_MIIRIMIMAKI_MANDAY_COLUMN_NUM: 10,
  SCENE_TOTAL_MIIRIMIMAKI_NUMBER_COLUMN_NUM: 11,
  SCENE_TOTAL_MIIRIMIMAKI_PERCENTAGE_COLUMN_NUM: 12,

  //シーン算出実行日を記載する
  SCENE_LAST_CALCULATION_DATE_ROW_NUM: 1,
  SCENE_LAST_CALCULATION_DATE_COLUMN_NUM: 9,

  //カット情報
  CUT_START_ROW_NUM: 3,
  CUT_START_ROW_INDEX: 3 - 1,
  //No から 原図フォルダまでの列数
  CUT_BLOCK_COLUMN_NUM: 11,
  CUT_NO_START_COLUMN_NUM: 172,
  CUT_NO_START_COLUMN_INDEX: 172 - 1,
  CUT_MAKIZUMI_COLUMN_FROM_NO_NUM: 7,
  CUT_SCENE_COLUMN_FROM_NO_NUM: 3,
  CUT_MANHOUR_COLUMN_FROM_NO_NUM: 2,
  CUT_DUALUSE_COLUMN_FROM_NO_NUM: 1,
  CUT_BGAGARI_COLUMN_FROM_NO_NUM: 8,
  //一日当たり何時間と設定するか
  HOURS_PER_DAY: 8,

  //エラーメッセージ
  ERROR_MESSAGE_DUPLICATE_NO: `このNo値が複数存在します`, //進行表シート名
  ERROR_MESSAGE_INVALID_MANHOUR: `工数(H)が数字以外になっています`,
  ERROR_MESSAGE_INVALID_DUALUSE_FORMAT: `兼用が/から始まるフォーマットになっていません。`,
  ERROR_DIALOG_MESSAGE_PREFIX:
    "エラーが見つかりました。下記Noを修正してください\n",
  ERROR_MESSAGE_COLOR_MIIRI_NOT_SET: "未入りの色が設定されていません。",
  ERROR_MESSAGE_COLOR_MIMAKI_NOT_SET: "未撒きの色が設定されていません。",
  ERROR_MESSAGE_COLOR_MAKIZUMI_NOT_SET: "撒済の色が設定されていません。",
  ERROR_MESSAGE_EXCLUSION_ROW_NOT_FOUND:
    "「欠番、Bank、全セル」の行が見つかりませんでした。\nシーン列に「欠番、Bank、全セル」を追加してください。",
  ERROR_MESSAGE_EXCLUSION_COLOR_NOT_SET:
    "「欠番、Bank、全セル」の色が設定されていません。",
  ERROR_MESSAGE_WARNING: "警告:",
  ERROR_MESSAGE_UI_ACCESS_ERROR: "UIにアクセスできません:",
  ERROR_DIALOG_MESSAGE_MAN_HOUR_LIMIT:
    "工数は150日以上表示できません\n表示できないシーン",
  ERROR_MESSAGE_COLOR_MISSING:
    "塗りつぶし色設定が存在しないので設定してください。: ",
  ERROR_MESSAGE_COLOR_IS_CLEAR:
    "塗りつぶし色設定が透明です。透明以外に設定してください。: ",
};

//-----------------
//クラス
//-----------------
//シーン情報
class Scene {
  constructor(
    name,
    color,
    rowNumber,
    rearangeManDay,
    person,
    startDate,
    loStatus
  ) {
    //シーン名
    this.name = name;
    //行番号
    this.rowNumber = rowNumber;
    //色
    this.color = color;
    //LO状況
    this.loStatus = loStatus;
    //１日平均Cut数兼用含む(枚)
    this.averageCutsPerDay = 0;
    //Cut数
    this.cutNumber = 0;
    //BGあがり=納品カット
    this.BGNumber = 0;
    //総工数(H)
    this.totalManhour = 0;
    //描枚数
    this.drawNumber = 0;
    //総工数(日)
    this.totalManDay = 0;
    //未入り工数(日)
    this.miiriManDay = 0;
    //未入り(枚)
    this.miiriNumber = 0;
    //未撒き工数(日)
    this.mimakiManDay = 0;
    //未撒き(枚)
    this.mimakiNumber = 0;
    //撒済工数(日)
    this.makizumiManDay = 0;
    //撒済(枚)
    this.makizumiNumber = 0;
    //再調整後の総工数(日)
    this.rearangeTotalManDay = 0;
    //再調整後の撒済工数(日)
    this.rearangeMakizumiManDay = 0;
    //再調整工数(日)
    if (!rearangeManDay) {
      this.rearangeManDay = 0;
    } else {
      this.rearangeManDay = rearangeManDay;
    }
    //作業者
    this.person = person;
    //開始日
    this.startDate = startDate;

    //未入り工数(H)
    this.miiriManHour = 0;
    //未撒き工数(H)
    this.mimakiManHour = 0;
    //撒済工数(H)
    this.makizumiManHour = 0;
  }
}

//-----------------
//グローバル変数
//-----------------
//各シーンごとの情報を格納
let sceneArray = [];
//シーン算出ボタンを押す前のシーン毎の情報を格納
let oldSceneArray = [];
//兼用の文字列を格納
let dualUseArray = [];
//シートの全値取得
let dataValues = undefined;
//シートの全背景色取得
let allBackgrounds = undefined;
//欠番、Bank、全セルの色を格納
let exclusionColor = "";
//欠番、BANK、全セルの配列の行
let exclusionRowIndex = 0;
//欠番、BANK、全セルの枚数
let exclusionCutNumber = 0;
//未撒き色等を取得したものを格納
let colorNoMimaki = "";
let colorMakizumi = "";
let colorBgAgari = "";
//シーン列に背景色が入っている数
let rowsUntilTransparent = 0;
//欠番、BANK、全セルとしてカウントしたNoを格納
let excludedNos = new Set();

//進行表のボタンを押されたときに実行されるメイン関数
function updateProgressSheet() {
  const label = "updateProgressSheet";
  console.time(label);

  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // 進行表読み込み  //動的に読み込む必要あり。
  //let sheet = spreadsheet.getSheetByName(PROGRESS_SHEET_NAME)
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // シートの全データ、背景色取得
  getAllDataAndBackground(sheet);

  //シートの一定部分を初期化する。
  clearSheetSections(sheet);

  // シーンクラスの器作成
  createSceneClasses(sheet);

  //oldシーンクラスの値を設定
  getOldSceneClasses(
    sheet,
    UP_Const.SCENE_START_BODY_ROW_INDEX,
    UP_Const.SCENE_AVERAGE_CUT_COLUMN_INDEX,
    UP_Const.SCENE_MAKIZUMI_NUMBER_COLUMN_NUM
  );

  //カット管理のデータのチェック
  let isDataConsistent = checkDataConsistency(sheet);
  if (!isDataConsistent) {
    return;
  }

  //色設定のデータチェック
  let isColorSettingValid = CheckColorSettings(sheet);
  if (!isColorSettingValid) {
    return;
  }

  // シーンクラスの値設定
  updateSceneClasses(sheet);

  // 描画行く前に値のチェック
  let validateResult = validate();
  if (!validateResult) {
    console.log("Value is invalid.");
    //return
  }

  // 表に値を描画
  drawTable(sheet);

  // 工数セルに描画
  drawProgressCells(sheet);

  //実行日と時間を前回算出実行日を記入する
  displayCurrentDateTime(sheet);

  console.timeEnd(label);
}

//実行日と時間を前回算出実行日を記入する
function displayCurrentDateTime(sheet) {
  // 現在の日付と時間を取得
  var now = new Date();
  // 指定されたフォーマットに変換
  var formattedDate = Utilities.formatDate(
    now,
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(),
    "yyyy/MM/dd HH:mm"
  );
  // 日付と時間を設定
  sheet
    .getRange(
      UP_Const.SCENE_LAST_CALCULATION_DATE_ROW_NUM,
      UP_Const.SCENE_LAST_CALCULATION_DATE_COLUMN_NUM
    )
    .setValue(formattedDate);
}

//表上にあるすべての値と背景色を取得する関数
function getAllDataAndBackground(sheet) {
  // 値を取得
  dataValues = sheet.getDataRange().getValues();
  // 背景色を取得する
  allBackgrounds = sheet
    .getRange(1, 1, dataValues.length, dataValues[0].length)
    .getBackgrounds();
  //シーン列の背景色がついている数を取得する。
  rowsUntilTransparent = getRowsUntilTransparent(
    allBackgrounds,
    UP_Const.SCENE_COLUMN_INDEX,
    UP_Const.SCENE_START_BODY_ROW_INDEX
  );
}

function getRowsUntilTransparent(
  allBackgrounds,
  columnIndex,
  startRowIndex = 0
) {
  // 指定された列の背景色を取得する
  let columnBackgrounds = allBackgrounds
    .slice(startRowIndex)
    .map((row) => row[columnIndex]);

  // 透明な背景色が見つかるまでの行数を取得する
  let rowCount = 0;
  for (let i = 0; i < columnBackgrounds.length; i++) {
    if (columnBackgrounds[i] === UP_Const.COLOR_CLEAR) {
      // 透明な背景色の場合
      break;
    }
    rowCount++;
  }

  return rowCount;
}

//初期化を行う為、値の削除を行う。
function clearSheetSections(sheet) {
  //シーン管理の入力と塗りつぶし初期化(Cut数列～撒済列まで)
  clearContentsAndBackgrounds(
    sheet,
    UP_Const.SCENE_START_BODY_ROW_INDEX,
    UP_Const.SCENE_LO_STATUS_COLUMN_NUM,
    UP_Const.SCENE_MAKIZUMI_NUMBER_COLUMN_NUM
  );
  //合計列だけ消す
  sheet
    .getRange(
      UP_Const.SCENE_ALL_SCENE_TOTAL_ROW_NUM,
      UP_Const.SCENE_LO_STATUS_COLUMN_NUM,
      1,
      UP_Const.SCENE_MAKIZUMI_NUMBER_COLUMN_NUM
    )
    .clearContent()
    .setBackground(UP_Const.COLOR_TOTAL_BACKGROUND);
  //未納Cut,合計LO未入り、合計LO入りを初期化
  sheet
    .getRange(
      UP_Const.SCENE_TOTAL_LO_ROW_NUM,
      UP_Const.SCENE_TOTAL_MINOU_NUMBER_COLUMN_NUM
    )
    .clearContent()
    .setBackground(UP_Const.COLOR_TOTAL_BACKGROUND);
  sheet
    .getRange(
      UP_Const.SCENE_TOTAL_LO_ROW_NUM,
      UP_Const.SCENE_TOTAL_LOMIIRI_NUMBER_COLUMN_NUM,
      1,
      2
    )
    .clearContent()
    .setBackground(UP_Const.COLOR_TOTAL_BACKGROUND);
  sheet
    .getRange(
      UP_Const.SCENE_TOTAL_LO_ROW_NUM,
      UP_Const.SCENE_TOTAL_LO_NUMBER_COLUMN_NUM,
      1,
      2
    )
    .clearContent()
    .setBackground(UP_Const.COLOR_TOTAL_BACKGROUND);
  //未入・未撒き合計％を初期化
  sheet
    .getRange(
      UP_Const.SCENE_TOTAL_MIIRIMIMAKI_ROW_NUM,
      UP_Const.SCENE_TOTAL_MIIRIMIMAKI_MANDAY_COLUMN_NUM,
      1,
      3
    )
    .clearContent()
    .setBackground(UP_Const.COLOR_MIIRIMIMAKI_TOTAL_BACKGROUND);
  //グラフの塗りつぶし範囲L列～BI列まで
  clearContentsAndBackgrounds(
    sheet,
    UP_Const.SCENE_START_BODY_ROW_NUM,
    UP_Const.SCENE_START_MANHOUR_COLUMN_NUM,
    UP_Const.SCENE_END_MANHOUR_COLUMN_NUM
  );
}

//工数列が数字以外になっていないか、兼用列が"/"で始まっているか？カットNoがダブっていないか？
//一度全て確認して、一気にエラーとしてダイアログに出す。
function checkDataConsistency(sheet) {
  dataValues = sheet.getDataRange().getValues();

  let cutNoColumnIndex = UP_Const.CUT_NO_START_COLUMN_INDEX;
  let errorMessages = []; // エラーメッセージを保存するための配列
  let checkedNos = new Set(); // 既にチェックしたno値を保存するためのSetオブジェクト
  let reportedNos = new Set(); // エラーメッセージが既に追加されたno値を保存するためのSetオブジェクト

  while (true) {
    let rowArray = dataValues[UP_Const.CUT_START_ROW_INDEX];
    if (cutNoColumnIndex >= rowArray.length) {
      break;
    }
    let no = rowArray[cutNoColumnIndex];
    if (!no) {
      break;
    }
    //一行ずつNoを見ていき必要な値を取得する。

    let positionIndex = UP_Const.CUT_START_ROW_INDEX;

    while (true) {
      if (positionIndex >= dataValues.length) {
        break;
      }
      let rowArray = dataValues[positionIndex];
      //該当のINDEXのデータがない＝該当の行に何もデータがない場合は処理抜ける
      if (cutNoColumnIndex >= rowArray.length) {
        break;
      }

      let no = rowArray[cutNoColumnIndex].toString();
      //NO列に何も値が設定されていない場合は処理抜ける
      if (!no) {
        break;
      }

      //Noがダブっていないかチェック
      if (checkedNos.has(no)) {
        // まだエラーメッセージが追加されていなければ、エラーメッセージを追加
        if (!reportedNos.has(no)) {
          errorMessages.push(
            `No ${no}: ${UP_Const.ERROR_MESSAGE_DUPLICATE_NO}`
          );
          reportedNos.add(no);
        }
      } else {
        checkedNos.add(no);
      }

      let manHour =
        rowArray[cutNoColumnIndex + UP_Const.CUT_MANHOUR_COLUMN_FROM_NO_NUM];

      //工数列が数字以外になっていないかチェック
      if (isNaN(manHour)) {
        errorMessages.push(
          `No ${no}: ${UP_Const.ERROR_MESSAGE_INVALID_MANHOUR}`
        );
      }
      let dualUse = String(
        rowArray[cutNoColumnIndex + UP_Const.CUT_DUALUSE_COLUMN_FROM_NO_NUM]
      );

      //兼用列が"/○○" や "/○○/○○" の形式であるかチェック（空欄でない場合のみ）
      if (dualUse.length > 0 && !/^\/[^\/]+(\/[^\/]+)*$/.test(dualUse)) {
        errorMessages.push(
          `No ${no}: ${UP_Const.ERROR_MESSAGE_INVALID_DUALUSE_FORMAT}`
        );
      }
      positionIndex++;
    }

    cutNoColumnIndex = cutNoColumnIndex + UP_Const.CUT_BLOCK_COLUMN_NUM;
  }

  // エラーメッセージがあればダイアログとして表示
  if (errorMessages.length > 0) {
    let ui = SpreadsheetApp.getUi();
    ui.alert(UP_Const.ERROR_DIALOG_MESSAGE_PREFIX + errorMessages.join("\n"));
    return false;
  }

  return true;
}

//未入り、未撒き、撒済の色について無色ならば、ダイアログを出してfalseを返す。
//欠番、Bank、全セルの背景色が無色ならば、ダイアログを出してfalseを返す。
//原図入稿済み、納品済、撒済が無色ならばダイアログを出してfalseを返す。
//上記はまとめてダイアログに表示する。
function CheckColorSettings(sheet) {
  let errorMessages = []; // エラーメッセージを保存するための配列

  //未入り、未撒き、撒済の色を取得
  let backgroundArray = allBackgrounds[UP_Const.SCENE_HEADER_ROW_INDEX];
  let miiriColor = backgroundArray[UP_Const.SCENE_MIIRI_MANHOUR_COLUMN_INDEX];
  let mimakiColor = backgroundArray[UP_Const.SCENE_MIMAKI_MANHOUR_COLUMN_INDEX];
  let makizumiColor =
    backgroundArray[UP_Const.SCENE_MAKIZUMI_MANHOUR_COLUMN_INDEX];

  //色が設定されていない場合エラーメッセージを追加
  if (miiriColor == UP_Const.COLOR_CLEAR)
    errorMessages.push(UP_Const.ERROR_MESSAGE_COLOR_MIIRI_NOT_SET);
  if (mimakiColor == UP_Const.COLOR_CLEAR)
    errorMessages.push(UP_Const.ERROR_MESSAGE_COLOR_MIMAKI_NOT_SET);
  if (makizumiColor == UP_Const.COLOR_CLEAR)
    errorMessages.push(UP_Const.ERROR_MESSAGE_COLOR_MAKIZUMI_NOT_SET);

  //欠番、Bank、全セルの背景色を取得
  exclusionRowIndex = findFirstRowWithExclusion(dataValues); //シーン列から欠番等の行を取得
  if (exclusionRowIndex === -1) {
    errorMessages.push(UP_Const.ERROR_MESSAGE_EXCLUSION_ROW_NOT_FOUND);
  } else {
    exclusionColor =
      allBackgrounds[exclusionRowIndex][UP_Const.SCENE_COLUMN_INDEX];
    if (exclusionColor == UP_Const.COLOR_CLEAR)
      errorMessages.push(UP_Const.ERROR_MESSAGE_EXCLUSION_COLOR_NOT_SET);
  }

  //原図入稿済み、納品済、撒済の背景色を取得
  let colorSettings = getColorSettingsFromRows(sheet, dataValues);

  for (let keyword of UP_Const.COLOR_SETTING_KEYWORDS) {
    if (!(keyword in colorSettings) || colorSettings[keyword] === "") {
      errorMessages.push(UP_Const.ERROR_MESSAGE_COLOR_MISSING + keyword);
    } else if (colorSettings[keyword] === UP_Const.COLOR_CLEAR) {
      errorMessages.push(UP_Const.ERROR_MESSAGE_COLOR_IS_CLEAR + keyword);
    }
  }

  // 各色に当てはめる、グローバル変数に色設定を代入
  assignColorSettings(colorSettings);

  // エラーメッセージがあればダイアログとして表示
  if (errorMessages.length > 0) {
    // UIを取得できるか確認
    try {
      let ui = SpreadsheetApp.getUi();
      ui.alert(
        UP_Const.ERROR_MESSAGE_WARNING + "\n" + errorMessages.join("\n")
      );
    } catch (e) {
      Logger.log(errorMessages);
      Logger.log(UP_Const.ERROR_MESSAGE_UI_ACCESS_ERROR + " " + e.toString());
    }
    return false;
  }
  return true;
}

//シーン管理の背景と入力文字を消去する。
function clearContentsAndBackgrounds(
  sheet,
  startRowIndex,
  startColumnIndex,
  endColumnIndex
) {
  //シーン管理Cut数列～撒済列まで
  let endRowIndex = startRowIndex + rowsUntilTransparent - 1;
  // 範囲を取得
  let range = sheet.getRange(
    startRowIndex,
    startColumnIndex,
    endRowIndex - startRowIndex + 1,
    endColumnIndex - startColumnIndex + 1
  );
  // 範囲内のセルの内容と背景色を消す
  range.clearContent();
  range.setBackground(null);
}

//セルの塗りつぶしが可能か判断する
function validate() {
  let validateResult = true;
  let exceededScenes = []; // 50日以上のシーンを保持する配列

  sceneArray.forEach((scene) => {
    //未入り、未撒き、撒済を整数に繰り上げてグラフ化する。
    let rearrangeMakiZumiManDay = Math.ceil(scene.rearangeMakizumiManDay);
    let mimakiManDay = Math.ceil(scene.mimakiManDay);
    let miiriManDay = Math.ceil(scene.miiriManDay);
    let totalManDay = rearrangeMakiZumiManDay + mimakiManDay + miiriManDay;

    if (totalManDay > UP_Const.MAX_TOTAL_MANDAY) {
      exceededScenes.push(scene.name); // エラーが検出されたシーン名を配列に追加
      validateResult = false;
    }
  });

  // 50日以上のシーンが存在する場合、ダイアログを表示
  if (exceededScenes.length > 0) {
    showManHourLimitDialog(exceededScenes);
  }

  return validateResult;
}

//工数の塗りつぶしは50日以上表示できない旨をダイアログにて表示する
function showManHourLimitDialog(exceededScenes) {
  var ui = SpreadsheetApp.getUi();
  const errorMessage = `${
    UP_Const.ERROR_DIALOG_MESSAGE_MAN_HOUR_LIMIT
  }\n${exceededScenes.join("\n")}`;
  ui.alert(errorMessage);
}

// 工数セルに描画
function drawProgressCells(sheet) {
  let backgroundArray = allBackgrounds[UP_Const.SCENE_HEADER_ROW_INDEX];

  let miiriColor = backgroundArray[UP_Const.SCENE_MIIRI_MANHOUR_COLUMN_INDEX];
  let mimakiColor = backgroundArray[UP_Const.SCENE_MIMAKI_MANHOUR_COLUMN_INDEX];
  let makizumiColor =
    backgroundArray[UP_Const.SCENE_MAKIZUMI_MANHOUR_COLUMN_INDEX];

  sceneArray.forEach((scene) => {
    let startPosition = UP_Const.SCENE_START_MANHOUR_COLUMN_NUM;
    let remainingDays = UP_Const.MAX_TOTAL_MANDAY; // 新しい変数を作成し、最初は MAX_TOTAL_MANDAY で初期化

    let makiZumiManDay = Math.ceil(scene.rearangeMakizumiManDay);
    makiZumiManDay = Math.min(makiZumiManDay, remainingDays);
    if (makiZumiManDay > 0) {
      sheet
        .getRange(scene.rowNumber, startPosition, 1, makiZumiManDay)
        .setBackground(makizumiColor);
      startPosition += makiZumiManDay;
      remainingDays -= makiZumiManDay; // 塗りつぶした日数を引く
    }

    let mimakiManDay = Math.ceil(scene.mimakiManDay);
    mimakiManDay = Math.min(mimakiManDay, remainingDays);
    if (mimakiManDay > 0) {
      sheet
        .getRange(scene.rowNumber, startPosition, 1, mimakiManDay)
        .setBackground(mimakiColor);
      startPosition += mimakiManDay;
      remainingDays -= mimakiManDay; // 塗りつぶした日数を引く
    }

    let miiriManDay = Math.ceil(scene.miiriManDay);
    miiriManDay = Math.min(miiriManDay, remainingDays);
    if (miiriManDay > 0) {
      sheet
        .getRange(scene.rowNumber, startPosition, 1, miiriManDay)
        .setBackground(miiriColor);
      startPosition += miiriManDay;
      remainingDays -= miiriManDay; // 塗りつぶした日数を引く
    }

    // 残りの部分を白でクリア
    if (remainingDays > 0) {
      sheet
        .getRange(scene.rowNumber, startPosition, 1, remainingDays)
        .setBackground(UP_Const.COLOR_CLEAR);
    }
    //もしもセルが塗られなかったら、startpostionより左のセルに記入してしまう為、＋1する
    if (startPosition == UP_Const.SCENE_START_MANHOUR_COLUMN_NUM) {
      startPosition++;
    }
    appendTotalManday(sheet, scene, startPosition); //最後のセルに合計値を入れる
  });
}

// 最後のセルに合計値を入れる
function appendTotalManday(sheet, scene, startPosition) {
  let lastPosition = startPosition - 1; // 塗りつぶした最後の列
  var cell = sheet.getRange(scene.rowNumber, lastPosition);
  cell.setValue("'" + scene.rearangeTotalManDay);

  // グラフが1セルだけだった場合は、左揃え、それ以外は右揃え
  if (scene.rearangeTotalManDay > 1) {
    cell.setHorizontalAlignment("right");
  } else {
    cell.setHorizontalAlignment("left");
  }
}

//工数の表を更新
function drawTable(sheet) {
  let SceneTotals = {
    cut: 0, //Cut数合計
    deliveryCut: 0, //納品カット(BGあがり)の合計
    totalManDay: 0, //総工数の合計
    drawNumber: 0, //総描枚数の合計
    miiriManDay: 0, //未入り工数の合計
    miiriNumber: 0, //未入り枚数の合計
    mimakiManDay: 0, //未撒き工数の合計
    mimakiNumber: 0, //未撒き枚数の合計
    makizumiManDay: 0, //撒済工数の合計
    makizumiNumber: 0, //撒済枚数の合計
    rearangeManDay: 0, //再調整の合計
    minouNumber: 0, //未納Cut ＝「Cut数」の合計 -「納品Cut」の合計
    totalLOMiiriNumber: 0, //合計LO未入り = 合計の「未入り(枚)」を転記
    totalLOMiiriPercentage: 0, //合計LO未入り％（端数不要）＝「合計LO未入り」枚 ÷「総描枚数」×100
    totalLONumber: 0, //合計LO入り ＝ 合計の「未撒き(枚)」+ 合計の「撒き済み(枚)」
    totalLOPercentage: 0, //合計LO入り％（端数不要）＝「合計LO入り」枚÷「総描枚数」×100
    totalMiiriMimakiNumber: 0, //合計の「未入り(枚)」＋合計の「未撒き(枚)」
    totalMiiriMimakiManDay: 0, //合計の「未入り(日)」＋合計の「未撒き(日)」
    totalMiiriMimakiPercentage: 0, //未入り・未撒き合計％＝「未入と未撒き合計(枚数)」÷「総描枚数(枚数)」×100
  };

  //シーン毎に値を記入
  sceneArray.forEach((scene) => {
    let range = sheet.getRange(
      scene.rowNumber,
      UP_Const.SCENE_AVERAGE_CUT_COLUMN_NUM
    );
    range.setValue(scene.averageCutsPerDay);
    range = sheet.getRange(
      scene.rowNumber,
      UP_Const.SCENE_CUT_NUMBER_COLUMN_NUM
    );
    range.setValue(scene.cutNumber);
    range = sheet.getRange(scene.rowNumber, UP_Const.SCENE_DRAW_COLUMN_NUM);
    range.setValue(scene.drawNumber);
    range = sheet.getRange(scene.rowNumber, UP_Const.SCENE_BG_COLUMN_NUM);
    range.setValue(scene.BGNumber);
    range = sheet.getRange(
      scene.rowNumber,
      UP_Const.SCENE_TOTAL_MANHOUR_COLUMN_NUM
    );
    range.setValue(scene.totalManDay);
    range = sheet.getRange(
      scene.rowNumber,
      UP_Const.SCENE_MIIRI_MANHOUR_COLUMN_NUM
    );
    range.setValue(scene.miiriManDay);
    range = sheet.getRange(
      scene.rowNumber,
      UP_Const.SCENE_MIMAKI_MANHOUR_COLUMN_NUM
    );
    range.setValue(scene.mimakiManDay);
    range = sheet.getRange(
      scene.rowNumber,
      UP_Const.SCENE_MAKIZUMI_MANHOUR_COLUMN_NUM
    );
    range.setValue(scene.makizumiManDay);
    //未入り、未撒き、撒済の枚数を入力する。
    range = sheet.getRange(
      scene.rowNumber,
      UP_Const.SCENE_MIIRI_NUMBER_COLUMN_NUM
    );
    range.setValue(scene.miiriNumber);
    range = sheet.getRange(
      scene.rowNumber,
      UP_Const.SCENE_MIMAKI_NUMBER_COLUMN_NUM
    );
    range.setValue(scene.mimakiNumber);
    range = sheet.getRange(
      scene.rowNumber,
      UP_Const.SCENE_MAKIZUMI_NUMBER_COLUMN_NUM
    );
    range.setValue(scene.makizumiNumber);

    if (scene.name === UP_Const.EXCLUSION_KEYWORD) {
      //「欠番、Bank、全セル」のカット数記入
      let range = sheet.getRange(
        scene.rowNumber,
        UP_Const.SCENE_CUT_NUMBER_COLUMN_NUM
      );
      range.setValue(exclusionCutNumber);
      SceneTotals.cut += exclusionCutNumber; //Cut数は欠番、Bank、全セルを合わせる
    }

    setLoStatusAndBackgroundColor(sheet, scene); //シーン毎の原図状況確認(LO揃等)

    //合計値を算出するときシーン名が_から始まる場合は、値を含まない。
    if (startsWithUnderscoreRegex(scene.name)) {
      return;
    }
    //各合計値へシーンの各値を足し合わせる
    SceneTotals.cut += scene.cutNumber;
    SceneTotals.deliveryCut += scene.BGNumber;
    SceneTotals.totalManDay += scene.totalManDay;
    SceneTotals.drawNumber += scene.drawNumber;
    SceneTotals.miiriManDay += scene.miiriManDay;
    SceneTotals.miiriNumber += scene.miiriNumber;
    SceneTotals.mimakiManDay += scene.mimakiManDay;
    SceneTotals.mimakiNumber += scene.mimakiNumber;
    SceneTotals.makizumiManDay += scene.makizumiManDay;
    SceneTotals.makizumiNumber += scene.makizumiNumber;
    SceneTotals.rearangeManDay += scene.rearangeManDay;
  });
  //未納Cut、合計LO未入り、合計LO入りを算出
  SceneTotals.minouNumber = SceneTotals.cut - SceneTotals.deliveryCut;
  SceneTotals.totalLOMiiriNumber = SceneTotals.miiriNumber;
  SceneTotals.totalLOMiiriPercentage =
    SceneTotals.totalLOMiiriNumber / SceneTotals.drawNumber;
  SceneTotals.totalLONumber =
    SceneTotals.mimakiNumber + SceneTotals.makizumiNumber;
  SceneTotals.totalLOPercentage =
    SceneTotals.totalLONumber / SceneTotals.drawNumber;
  //未入・未撒き合計％を算出
  SceneTotals.totalMiiriMimakiNumber =
    SceneTotals.miiriNumber + SceneTotals.makizumiNumber;
  SceneTotals.totalMiiriMimakiManDay =
    SceneTotals.miiriManDay + SceneTotals.mimakiManDay;
  SceneTotals.totalMiiriMimakiPercentage =
    SceneTotals.totalMiiriMimakiNumber / SceneTotals.drawNumber;

  setTotalValues(sheet, SceneTotals); //各合計値を入力
}

//各合計値を入力
function setTotalValues(sheet, SceneTotals) {
  sheet
    .getRange(
      UP_Const.SCENE_ALL_SCENE_TOTAL_ROW_NUM,
      UP_Const.SCENE_CUT_NUMBER_COLUMN_NUM
    )
    .setValue(SceneTotals.cut);
  sheet
    .getRange(
      UP_Const.SCENE_ALL_SCENE_TOTAL_ROW_NUM,
      UP_Const.SCENE_BG_COLUMN_NUM
    )
    .setValue(SceneTotals.deliveryCut);
  sheet
    .getRange(
      UP_Const.SCENE_ALL_SCENE_TOTAL_ROW_NUM,
      UP_Const.SCENE_TOTAL_MANHOUR_COLUMN_NUM
    )
    .setValue(SceneTotals.totalManDay)
    .setFontWeight("bold");
  sheet
    .getRange(
      UP_Const.SCENE_ALL_SCENE_TOTAL_ROW_NUM,
      UP_Const.SCENE_DRAW_COLUMN_NUM
    )
    .setValue(SceneTotals.drawNumber)
    .setFontWeight("bold");
  sheet
    .getRange(
      UP_Const.SCENE_ALL_SCENE_TOTAL_ROW_NUM,
      UP_Const.SCENE_MIIRI_MANHOUR_COLUMN_NUM
    )
    .setValue(SceneTotals.miiriManDay);
  sheet
    .getRange(
      UP_Const.SCENE_ALL_SCENE_TOTAL_ROW_NUM,
      UP_Const.SCENE_MIIRI_NUMBER_COLUMN_NUM
    )
    .setValue(SceneTotals.miiriNumber);
  sheet
    .getRange(
      UP_Const.SCENE_ALL_SCENE_TOTAL_ROW_NUM,
      UP_Const.SCENE_MIMAKI_MANHOUR_COLUMN_NUM
    )
    .setValue(SceneTotals.mimakiManDay);
  sheet
    .getRange(
      UP_Const.SCENE_ALL_SCENE_TOTAL_ROW_NUM,
      UP_Const.SCENE_MIMAKI_NUMBER_COLUMN_NUM
    )
    .setValue(SceneTotals.mimakiNumber);
  sheet
    .getRange(
      UP_Const.SCENE_ALL_SCENE_TOTAL_ROW_NUM,
      UP_Const.SCENE_MAKIZUMI_MANHOUR_COLUMN_NUM
    )
    .setValue(SceneTotals.makizumiManDay);
  sheet
    .getRange(
      UP_Const.SCENE_ALL_SCENE_TOTAL_ROW_NUM,
      UP_Const.SCENE_MAKIZUMI_NUMBER_COLUMN_NUM
    )
    .setValue(SceneTotals.makizumiNumber);
  sheet
    .getRange(
      UP_Const.SCENE_ALL_SCENE_TOTAL_ROW_NUM,
      UP_Const.SCENE_REARANGE_MANHOUR_COLUMN_NUM
    )
    .setValue(SceneTotals.rearangeManDay);
  //未納Cut、合計LO未入り、合計LO入りを算出
  sheet
    .getRange(
      UP_Const.SCENE_TOTAL_LO_ROW_NUM,
      UP_Const.SCENE_TOTAL_MINOU_NUMBER_COLUMN_NUM
    )
    .setValue(SceneTotals.minouNumber);
  sheet
    .getRange(
      UP_Const.SCENE_TOTAL_LO_ROW_NUM,
      UP_Const.SCENE_TOTAL_LOMIIRI_NUMBER_COLUMN_NUM
    )
    .setValue(SceneTotals.totalLOMiiriNumber);
  sheet
    .getRange(
      UP_Const.SCENE_TOTAL_LO_ROW_NUM,
      UP_Const.SCENE_TOTAL_LOMIIRI_PERCENTAGE_COLUMN_NUM
    )
    .setValue(SceneTotals.totalLOMiiriPercentage);
  sheet
    .getRange(
      UP_Const.SCENE_TOTAL_LO_ROW_NUM,
      UP_Const.SCENE_TOTAL_LO_NUMBER_COLUMN_NUM
    )
    .setValue(SceneTotals.totalLONumber);
  sheet
    .getRange(
      UP_Const.SCENE_TOTAL_LO_ROW_NUM,
      UP_Const.SCENE_TOTAL_LO_PERCENTAGE_COLUMN_NUM
    )
    .setValue(SceneTotals.totalLOPercentage);
  //未入・未撒き合計％を算出
  sheet
    .getRange(
      UP_Const.SCENE_TOTAL_MIIRIMIMAKI_ROW_NUM,
      UP_Const.SCENE_TOTAL_MIIRIMIMAKI_MANDAY_COLUMN_NUM
    )
    .setValue(SceneTotals.totalMiiriMimakiManDay);
  sheet
    .getRange(
      UP_Const.SCENE_TOTAL_MIIRIMIMAKI_ROW_NUM,
      UP_Const.SCENE_TOTAL_MIIRIMIMAKI_NUMBER_COLUMN_NUM
    )
    .setValue(SceneTotals.totalMiiriMimakiNumber);
  sheet
    .getRange(
      UP_Const.SCENE_TOTAL_MIIRIMIMAKI_ROW_NUM,
      UP_Const.SCENE_TOTAL_MIIRIMIMAKI_PERCENTAGE_COLUMN_NUM
    )
    .setValue(SceneTotals.totalMiiriMimakiPercentage);
}

//LO状況列に入力＋背景色の塗りつぶしを行う。
function setLoStatusAndBackgroundColor(sheet, scene) {
  /*
  LO揃＝シーンの描枚数すべての原図が入庫した。：未入りが0枚の状態
    miiriNumber=0
  撒済＝シーンの描枚数すべてを作業者に撒いた。：LO揃＆未撒きが0枚の状態
    miiriNumber=0&mimakiNumber=0
  T1済＝シーンのCut数すべてがBG上がり状態になり、一度後工程に納品した。：納品Cut数と撒済枚数が同じ状態
    BGNumberとcutNumが一致した場合、指定範囲をグレーで塗りつぶす
  */
  let startColumn = UP_Const.SCENE_LO_STATUS_COLUMN_NUM;
  let endColumn = UP_Const.SCENE_MAKIZUMI_NUMBER_COLUMN_NUM;

  if (scene.BGNumber === scene.cutNumber) {
    //T1済
    sheet
      .getRange(scene.rowNumber, startColumn, 1, endColumn - startColumn + 1)
      .setBackground(UP_Const.COLOR_LO_STATUS_T1);
    sheet
      .getRange(scene.rowNumber, startColumn)
      .setValue(UP_Const.T1_KEYWORD)
      .setHorizontalAlignment("center")
      .setFontWeight("bold")
      .setFontSize(10)
      .setFontColor(UP_Const.TEXT_COLOR_LO_STATUS_T1);
  } else if (scene.miiriNumber == 0 && scene.mimakiNumber == 0) {
    //撒済
    sheet
      .getRange(scene.rowNumber, startColumn, 1, endColumn - startColumn + 1)
      .setBackground(UP_Const.COLOR_LO_STATUS_MAKIZUMI);
    sheet
      .getRange(scene.rowNumber, startColumn)
      .setValue(UP_Const.MAKIZUMI_KEYWORD)
      .setFontWeight("normal")
      .setHorizontalAlignment("center")
      .setFontSize(10)
      .setFontColor(UP_Const.TEXT_COLOR_LO_STATUS_MAKIZUMI);
  } else if (scene.miiriNumber == 0) {
    //LO揃
    sheet
      .getRange(scene.rowNumber, startColumn, 1, endColumn - startColumn + 1)
      .setBackground(UP_Const.COLOR_LO_STATUS_LOZOROI);
    sheet
      .getRange(scene.rowNumber, startColumn)
      .setValue(UP_Const.LOZOROI_KEYWORD)
      .setFontWeight("normal")
      .setHorizontalAlignment("center")
      .setFontSize(10)
      .setFontColor(UP_Const.TEXT_COLOR_LO_STATUS_LOZOROI);
  } else {
    // 一致しない場合は背景色をCOLOR_CLEARに設定
    sheet
      .getRange(scene.rowNumber, startColumn, 1, endColumn - startColumn + 1)
      .setFontWeight("normal")
      .setFontSize(10)
      .setFontColor(UP_Const.TEXT_COLOR_NORMAL)
      .setBackground(UP_Const.COLOR_CLEAR);
  }
}

//シーン情報格納用のクラスを作成
function createSceneClasses(sheet) {
  let positionIndex = UP_Const.SCENE_START_BODY_ROW_INDEX;

  while (true) {
    //シーン
    let rowArray = dataValues[positionIndex];
    let name = rowArray[UP_Const.SCENE_COLUMN_INDEX];
    //シーン名は記載されている想定
    //if(!name){
    //break
    //}
    let backgroundArray = allBackgrounds[positionIndex];
    let background = backgroundArray[UP_Const.SCENE_COLUMN_INDEX];
    //背景が無色となるまで、枠を作成する。
    if (background === UP_Const.COLOR_CLEAR) {
      break;
    }
    //再調整工数
    let rearangeManDay = rowArray[UP_Const.SCENE_REARANGE_MANHOUR_COLUMN_INDEX];

    //作業者
    let person = rowArray[UP_Const.SCENE_PERSON_COLUMN_INDEX];
    //開始日
    let startDate = rowArray[UP_Const.SCENE_STARTDATE_COLUMN_INDEX];

    //シーンクラス作成
    // 配列のインデックスは0スタートだが、スプレッドシートの行は1スタートなので、positionIndexに+1を加える
    let newScene = new Scene(
      name,
      background,
      positionIndex + 1,
      rearangeManDay,
      person,
      startDate
    );
    sceneArray.push(newScene);

    //前回算出増減列用にシーンクラスを作成する
    let oldScene = new Scene(
      name,
      background,
      positionIndex + 1,
      rearangeManDay,
      person,
      startDate
    );
    oldSceneArray.push(oldScene);

    positionIndex++;
  }
}
//一行ずつNoを見ていき必要な値を取得する。
function updateSceneClassesByBlock(sheet, cutNoColumnIndex) {
  let positionIndex = UP_Const.CUT_START_ROW_INDEX;

  while (true) {
    if (positionIndex >= dataValues.length) {
      break;
    }
    let rowArray = dataValues[positionIndex];
    let backgroundArray = allBackgrounds[positionIndex];
    //該当のINDEXのデータがない＝該当の行に何もデータがない場合は処理抜ける
    if (cutNoColumnIndex >= rowArray.length) {
      break;
    }

    let no = rowArray[cutNoColumnIndex].toString();
    //NO列に何も値が設定されていない場合は処理抜ける
    if (!no) {
      break;
    }
    let color =
      backgroundArray[cutNoColumnIndex + UP_Const.CUT_SCENE_COLUMN_FROM_NO_NUM];
    let manHour =
      rowArray[cutNoColumnIndex + UP_Const.CUT_MANHOUR_COLUMN_FROM_NO_NUM];

    let noColor = backgroundArray[cutNoColumnIndex];
    let makizumiColor =
      backgroundArray[
        cutNoColumnIndex + UP_Const.CUT_MAKIZUMI_COLUMN_FROM_NO_NUM
      ];

    let dualUse = String(
      rowArray[cutNoColumnIndex + UP_Const.CUT_DUALUSE_COLUMN_FROM_NO_NUM]
    );
    dualUseArray.push(...dualUse.split("/").filter(Boolean));

    let bgAgariColor =
      backgroundArray[
        cutNoColumnIndex + UP_Const.CUT_BGAGARI_COLUMN_FROM_NO_NUM
      ];

    sceneArray.forEach((scene) => {
      if (scene.color == color) {
        caclulateNumber(
          scene,
          noColor,
          exclusionColor,
          no,
          dualUseArray,
          bgAgariColor,
          makizumiColor,
          manHour
        );
      }
    });

    positionIndex++;
  }
}

//['欠番、Bank、全セル']がある行を取り出す。
function findFirstRowWithExclusion(dataValues) {
  const keywords = [UP_Const.EXCLUSION_KEYWORD];

  let matchingRow = -1; // 初期値を-1（見つからなかった場合）に設定

  for (let i = 0; i < dataValues.length; i++) {
    if (keywords.includes(dataValues[i][UP_Const.SCENE_COLUMN_INDEX])) {
      matchingRow = i;
      break; // 最初の一致を見つけたらループを抜ける
    }
  }

  return matchingRow; // 該当する最初の行のインデックスを返す、見つからなかった場合は-1
}

//欠番、Bank、全セル以外か確認
function caclulateNumber(
  scene,
  noColor,
  exclusionColor,
  no,
  dualUseArray,
  bgAgariColor,
  makizumiColor,
  manHour
) {
  if (noColor === exclusionColor) {
    if (!excludedNos.has(no)) {
      exclusionCutNumber++; //欠番、BANK、全セルの枚数をカウント
      excludedNos.add(no); // カウントしたnoを追跡
    }
    return;
  }

  // noColor != exclusionColor の場合の処理
  //Cut数
  scene.cutNumber++;
  //BGあがり=納品Cut
  if (colorBgAgari == bgAgariColor) {
    scene.BGNumber++;
  }

  //描枚数
  //兼用枚数があった場合はカウントしない。
  if (!dualUseArray.includes(no)) {
    scene.drawNumber++;
  }
  //優先順位に基づいて、工数を入力する。撒済＞未撒き＞未入りという優先順位にする。
  //工数が入力されている場合
  if (manHour) {
    //兼用枚数があった場合はカウントしない。
    if (!dualUseArray.includes(no)) {
      // 総工数
      scene.totalManhour = scene.totalManhour + manHour;
      if (makizumiColor == colorMakizumi) {
        // 撒済工数
        scene.makizumiManHour = scene.makizumiManHour + manHour;

        // 撒済枚数
        scene.makizumiNumber++;
      } else if (noColor == colorNoMimaki) {
        // 未撒き工数
        scene.mimakiManHour = scene.mimakiManHour + manHour;
        // 未撒き枚数
        scene.mimakiNumber++;
      } else {
        // 未入り工数
        scene.miiriManHour = scene.miiriManHour + manHour;
        // 未入り枚数
        scene.miiriNumber++;
      }
    }
  }
}

//シーン情報格納用のクラスを更新
function updateSceneClasses(sheet) {
  //let cutNoColumn = CUT_NO_START_COLUMN_NUM
  let cutNoColumnIndex = UP_Const.CUT_NO_START_COLUMN_INDEX;
  dualUseArray = [];

  while (true) {
    let rowArray = dataValues[UP_Const.CUT_START_ROW_INDEX];
    if (cutNoColumnIndex >= rowArray.length) {
      break;
    }
    let no = rowArray[cutNoColumnIndex];
    if (!no) {
      break;
    }
    updateSceneClassesByBlock(sheet, cutNoColumnIndex);

    cutNoColumnIndex = cutNoColumnIndex + UP_Const.CUT_BLOCK_COLUMN_NUM;
  }
  sceneArray.forEach((scene) => {
    convertSceneManHourToDay(scene); //時間単位を日単位に変換する。
    //再調整列を入れた値を反映する。
    scene.rearangeTotalManDay = scene.totalManDay + scene.rearangeManDay;
    scene.rearangeMakizumiManDay = scene.makizumiManDay + scene.rearangeManDay;
  });

  //値をすべて取得したので、シーン毎の平均値を出す。「1日平均Cut数兼用含む」（0.1以下は四捨五入）＝「Cut数」÷「総工数日」
  sceneArray.forEach((scene) => {
    if (scene.totalManDay === 0) {
      scene.averageCutsPerDay = 0; // 0で割る場合の結果を0とする
    } else {
      scene.averageCutsPerDay = roundToTenth(
        scene.cutNumber / scene.totalManDay
      );
    }
  });
}

// 色設定をグローバル変数に代入する関数
function assignColorSettings(colorSettings) {
  colorNoMimaki = colorSettings[UP_Const.MIMAKI_KEYWORD];
  colorMakizumi = colorSettings[UP_Const.MAKIZUMI_KEYWORD];
  colorBgAgari = colorSettings[UP_Const.BG_AGARI_KEYWORD];
}

// ['原図入稿済', '納品済', '撒済']がある行をすべて取り出す。
function findColorSettingRowIndices(dataValues) {
  let matchingRows = []; // 該当する行のインデックスを格納するための配列

  for (let i = 0; i < dataValues.length; i++) {
    if (UP_Const.COLOR_SETTING_KEYWORDS.includes(dataValues[i][0])) {
      matchingRows.push(i); // 該当する行のインデックスを配列に追加
    }
  }

  return matchingRows; // 該当するすべての行のインデックスの配列を返す、見つからなかった場合は空の配列
}

//原図入稿済、納品済、撒済の背景色を取得する。
function getColorSettingsFromRows(sheet, dataValues) {
  // 色設定を保持するオブジェクトを作成
  let colorSettings = UP_Const.COLOR_SETTING_KEYWORDS.reduce((obj, keyword) => {
    obj[keyword] = "";
    return obj;
  }, {});

  // 色設定行のインデックスを取得
  let colorSettingRowsIndex = findColorSettingRowIndices(dataValues);

  // 各色設定行から色設定を取得
  colorSettingRowsIndex.forEach((rowIndex) => {
    let keyword = dataValues[rowIndex][0]; // キーワードを取得
    let color = sheet.getRange(rowIndex + 1, 2).getBackground(); // 色設定を取得（行インデックスは0から始まるので、シートの行番号に変換するために+1）

    // 色設定をオブジェクトに格納
    if (colorSettings.hasOwnProperty(keyword)) {
      colorSettings[keyword] = color;
    }
  });

  // 色設定を含むオブジェクトを返す
  return colorSettings;
}

//時間単位を日単位に変換する。
function convertSceneManHourToDay(scene) {
  scene.miiriManDay = scene.miiriManHour / UP_Const.HOURS_PER_DAY;
  scene.mimakiManDay = scene.mimakiManHour / UP_Const.HOURS_PER_DAY;
  scene.makizumiManDay = scene.makizumiManHour / UP_Const.HOURS_PER_DAY;
  scene.miiriManDay = roundUpToQuarter(scene.miiriManDay);
  scene.mimakiManDay = roundUpToQuarter(scene.mimakiManDay);
  scene.makizumiManDay = roundUpToQuarter(scene.makizumiManDay);
  scene.totalManDay =
    scene.miiriManDay + scene.mimakiManDay + scene.makizumiManDay; //総工数は0.25で切り上げ後を足し合わせたものとする。
}

//0.25単位で繰り上げる。
function roundUpToQuarter(num) {
  return Math.ceil(num * 4) / 4;
}
//0.1以下を四捨五入する　0.11＝＞0.1
function roundToTenth(value) {
  return Math.round(value * 10) / 10;
}

// 文字列が _ もしくは ＿ で始まる場合、trueを返す。それ以外はfalseを返す。
function startsWithUnderscoreRegex(inputString) {
  return /^[_＿]/.test(inputString);
}

//シーン算出ボタンが押される前の値を取得する
function getOldSceneClasses(
  sheet,
  startRowIndex,
  startColumnIndex,
  endColumnIndex
) {
  //シーン管理Cut数列～撒済列まで
  let endRowIndex = startRowIndex + rowsUntilTransparent;

  //記入があるもののみコピー
  oldSceneArray.forEach((scene) => {
    scene.averageCutsPerDay =
      dataValues[scene.rowNumber - 1][UP_Const.SCENE_AVERAGE_CUT_COLUMN_INDEX];
    scene.cutNumber =
      dataValues[scene.rowNumber - 1][UP_Const.SCENE_CUT_NUMBER_COLUMN_INDEX];
    scene.BGNumber =
      dataValues[scene.rowNumber - 1][UP_Const.SCENE_BG_COLUMN_INDEX];
    scene.drawNumber =
      dataValues[scene.rowNumber - 1][UP_Const.SCENE_DRAW_COLUMN_INDEX];
    scene.totalManDay =
      dataValues[scene.rowNumber - 1][
        UP_Const.SCENE_TOTAL_MANHOUR_COLUMN_INDEX
      ];
    scene.miiriManDay =
      dataValues[scene.rowNumber - 1][
        UP_Const.SCENE_MIIRI_MANHOUR_COLUMN_INDEX
      ];
    scene.mimakiManDay =
      dataValues[scene.rowNumber - 1][
        UP_Const.SCENE_MIMAKI_MANHOUR_COLUMN_INDEX
      ];
    scene.makizumiManDay =
      dataValues[scene.rowNumber - 1][
        UP_Const.SCENE_MAKIZUMI_MANHOUR_COLUMN_NUM
      ];
    scene.miiriNumber =
      dataValues[scene.rowNumber - 1][UP_Const.SCENE_MIIRI_NUMBER_COLUMN_INDEX];
    scene.mimakiNumber =
      dataValues[scene.rowNumber - 1][
        UP_Const.SCENE_MIMAKI_NUMBER_COLUMN_INDEX
      ];
    scene.makizumiNumber =
      dataValues[scene.rowNumber - 1][
        UP_Const.SCENE_MAKIZUMI_NUMBER_COLUMN_INDEX
      ];
    scene.rearangeManDay =
      dataValues[scene.rowNumber - 1][
        UP_Const.SCENE_REARANGE_MANHOUR_COLUMN_INDEX
      ];
    scene.rearangeMakizumiManDay = scene.makizumiManDay + scene.rearangeManDay;
  });
  console.log(oldSceneArray);
}