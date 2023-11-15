//-----------------
//定数定義
//-----------------
const DATA_BASE_SHEET_NAME = "データベース";
//セルの色塗り初期化時の色
const COLOR_CLEAR = "#ffffff";
// 土日祝日の色
const COLOR_HOLIDAY = "#808080";
// データベースとスケジュール表のフォーマットの差を埋める為、美術ルーム全体スケジュールmemoの行数分+CLW美術作業者一覧分-以下データスペース-作業者ベースデータで算出する
//const DATA_BASE_FORMAT_OFFSET = 5;
//カレンダー
const SS_CALENDERDATE_COLUMN_INDEX = 7;
const SS_CALENDERDATE_ROW_INDEX = 6;
//表示最大日数
const SS_MAXDAYS = 730; //730(24か月)
// 区切り文字
const DELIMITER = ",";
// 固定文字
const FIXED_CELL_KEYWORD = "▼▲▼";
const CLW_ARTWORK_PERSONS_TITLE = "CLW美術作業者一覧";
const FREE_SPACE_TITLE = "以下フリースぺース";
const MEMO_SPACE_TITLE = "memo";
//担当者
// const PERSON_COLUMN_INDEX = 6;

//-----------------
//変数
//-----------------
// スプレッドシート上でデータが入力されている最大範囲を選択
let scheduleSheetAllRange = undefined;
//スケジュール表の全値取得
let scheduleSheetDataValues = undefined;
//スケジュール表の背景色取得
let scheduleSheetAllBackGrounds = undefined;
// データベースシートに格納する為、データ領域を格納する配列
let dataBaseSheetDataValues = undefined;
// データベースシートの最大範囲
let dataBaseSheetAllRange = undefined;
// スプレッドシート上でデータが入力されている最大範囲を選択
let scheduleSheetPersonRange = undefined;
//スケジュール表の担当者領域取得
let dataBaseSheetPersonRange = undefined;
// データベースシートの担当者領域

//セル削除（削除後詰めなし）
function deleteCellsC(range, isAllData = true) {
  let rowIndex = isAllData ? range.getRow() - 1 : 0;
  let selectedColumns = range.getLastColumn() + 1 - range.getColumn();
  //選択範囲の数だけループ

  for (let i = 0; i < selectedColumns; i++) {
    let index = range.getColumn() - 1 + i;
    let deleteCellColor = scheduleSheetAllBackGrounds[rowIndex][index];
    //土日祝日だけ割り込み禁止
    if (deleteCellColor == COLOR_HOLIDAY) {
      continue;
    }
    //値、色クリア
    setSellC(rowIndex, index, COLOR_CLEAR, "", COLOR_CLEAR, "");
  }
}

//セルを隣に移動
function moveCellC(rowIndex, index, existColor, sceneTitle, manHourColor) {
  let nextExistColor = scheduleSheetAllBackGrounds[rowIndex][index];
  let isFixed = isCellFixed(scheduleSheetDataValues, rowIndex, index);
  //土日祝日・固定セルだけ割り込み禁止//
  if (nextExistColor == COLOR_HOLIDAY || isFixed) {
    index++;
    moveCellC(rowIndex, index, existColor, sceneTitle, manHourColor);
    return;
  }
  let nextSceneTitle = undefined;
  let nextManHourColor = undefined;
  let nextDataBaseSceneTitle = undefined;
  if (nextExistColor != COLOR_CLEAR) {
    nextSceneTitle = scheduleSheetDataValues[rowIndex][index];
    nextManHourColor = scheduleSheetAllBackGrounds[rowIndex + 1][index];
    //nextDataBaseSceneTitle = scheduleSheetDataValues[rowIndex][index];
    nextDataBaseSceneTitle = dataBaseSheetDataValues[rowIndex][index];
  }
  setSellC(rowIndex, index, existColor, sceneTitle, manHourColor, "");
  if (nextExistColor != COLOR_CLEAR) {
    index++;
    moveCellC(
      rowIndex,
      index,
      nextExistColor,
      nextDataBaseSceneTitle,
      nextManHourColor
    );
  }
}

//割り込みセルの挿入
function warikomiMoveCellC(
  rowIndex,
  startColumnIndex,
  totalDays,
  storyInfoColor,
  sceneTitle,
  manHourColors
) {
  // 何セル移動できるかの判断用と移動のコピー元用に、該当行の配列情報をコピーしておく
  let tmpRowBackGrounds = scheduleSheetAllBackGrounds[rowIndex].slice();
  let tmpRowStatusBackGrounds =
    scheduleSheetAllBackGrounds[rowIndex + 1].slice();
  let tmpRowDataBaseSheetDataValues = dataBaseSheetDataValues[rowIndex].slice();
  let tmpRowStatusDataBaseSheetDataValues =
    dataBaseSheetDataValues[rowIndex + 1].slice();

  // 移動幅・scene幅
  let countMove = 0;
  let countSceneMove = 0;
  for (let i = startColumnIndex; i < startColumnIndex + totalDays; i++) {
    //移動先セルが土日:ずらす値を+１、iをインクリメントせず土日でなくなるまでmovecをずらす
    while (
      tmpRowBackGrounds[i + countMove] == COLOR_HOLIDAY ||
      tmpRowStatusDataBaseSheetDataValues[i + countMove] === FIXED_CELL_KEYWORD
    ) {
      //while (tmpRowBackGrounds[i + countMove] == COLOR_HOLIDAY) {
      countMove++;
      continue;
    }
    setSellC(
      rowIndex,
      i + countMove,
      storyInfoColor,
      sceneTitle,
      manHourColors[i - startColumnIndex],
      ""
    );
    countSceneMove++;
  }
  // 移動幅にscene幅足しこみ
  countMove = countMove + countSceneMove;

  // // 移動幅範囲内の休日数
  // const rangeHoliday = tmpRowBackGrounds
  //   .slice(startColumnIndex, manHourColors.length)
  //   .filter((backGround) => backGround == COLOR_HOLIDAY).length;
  // // 見切れる範囲の休日数
  // const outHoliday = tmpRowBackGrounds
  //   .slice(tmpRowBackGrounds.length - countMove)
  //   .filter((backGround) => backGround == COLOR_HOLIDAY).length;
  // // 見切れる範囲の固定セル数
  // const outFixedDay = tmpRowStatusDataBaseSheetDataValues
  //   .slice(tmpRowStatusDataBaseSheetDataValues.length - countMove)
  //   .filter((value) => value == FIXED_CELL_KEYWORD).length;

  // // 空白セルの挿入時に見切れる範囲に初期値・休日以外がある場合エラー
  // const outRange =
  //   tmpRowBackGrounds.length -
  //   1 -
  //   (countMove - rangeHoliday + outHoliday + outFixedDay);
  // const outRangeInput = tmpRowBackGrounds
  //   .slice(outRange)
  //   .filter(
  //     (backGround) => backGround != COLOR_HOLIDAY && backGround != COLOR_CLEAR
  //   ).length;
  // if (outRangeInput > 0) {
  //   let ui = SpreadsheetApp.getUi();
  //   ui.alert("見切れる範囲に入力があります");
  //   return;
  // }
  //選択開始列　から　一番最終列までループ（１行だけ指定と想定）
  for (let i = startColumnIndex; i < tmpRowBackGrounds.length; i++) {
    //該当セルが土日:ずらす値を-１、iをインクリメント //TODO:
    if (
      tmpRowBackGrounds[i] == COLOR_HOLIDAY ||
      tmpRowStatusDataBaseSheetDataValues[i] === FIXED_CELL_KEYWORD ||
      tmpRowBackGrounds[i] == COLOR_CLEAR
    ) {
      countMove--;
      if (countMove <= 0) {
        break;
      }
      continue;
    }
    //移動先セルが土日:ずらす値を+１、iをインクリメントせず土日でなくなるまでmovecをずらす
    while (
      tmpRowBackGrounds[i + countMove] == COLOR_HOLIDAY ||
      tmpRowStatusDataBaseSheetDataValues[i + countMove] === FIXED_CELL_KEYWORD
    ) {
      countMove++;
    }
    if (i + countMove >= tmpRowBackGrounds.length) {
      throw new Error("見切れる範囲に入力があります");
    }
    setSellC(
      rowIndex,
      i + countMove,
      tmpRowBackGrounds[i],
      tmpRowDataBaseSheetDataValues[i],
      tmpRowStatusBackGrounds[i],
      ""
    );
  }
}

//セルに値を設定。表示用シーン名はクリア（後で一気に付ける）
function setSellC(
  rowIndex,
  columnIndex,
  setColor,
  sceneTitle,
  manHourColor,
  cellStatus
) {
  scheduleSheetDataValues[rowIndex][columnIndex] = "";
  scheduleSheetDataValues[rowIndex + 1][columnIndex] = cellStatus; //固定セルの表記を消す
  scheduleSheetAllBackGrounds[rowIndex][columnIndex] = setColor;
  scheduleSheetAllBackGrounds[rowIndex + 1][columnIndex] = manHourColor;
  dataBaseSheetDataValues[rowIndex][columnIndex] = sceneTitle;
  dataBaseSheetDataValues[rowIndex + 1][columnIndex] = cellStatus; //固定セルの表記を消す
}

//先頭セルにシーン名を表示する
function displaySceneNameC(rowIndex) {
  console.log("---displaySceneNameC---");
  let prevSceneTitle = "";
  let sceneAndDay = {}; // シーン名と日数の対応を保持するオブジェクト
  console.log("sceneAndDay:", sceneAndDay);
  for (
    let columnIndex = SS_CALENDERDATE_COLUMN_INDEX;
    columnIndex < SS_CALENDERDATE_COLUMN_INDEX + SS_MAXDAYS;
    columnIndex++
  ) {
    let sceneTitle = dataBaseSheetDataValues[rowIndex][columnIndex];
    // sceneTitleから開始日を除去
    sceneTitle = truncateTitle(sceneTitle, 3);
    // 日分を除いたシーン名を取得
    let sceneTitleOnly = truncateTitle(sceneTitle, 2);

    if (sceneTitle !== prevSceneTitle) {
      // 空欄後の最初のシーン
      if (!sceneAndDay.hasOwnProperty(sceneTitleOnly)) {
        sceneAndDay[sceneTitleOnly] = 0; // シーン名と日数の初期値をセットシーン名の追加
      }
      // シーンがカバーする日数をカウント
      let columnCount = countSceneDays(
        rowIndex,
        columnIndex,
        dataBaseSheetDataValues
      );
      // 値を追加
      sceneAndDay[sceneTitleOnly] = sceneAndDay[sceneTitleOnly] + columnCount;
      const number = extractDayNumberFromString(sceneTitle);
      if (sceneAndDay[sceneTitleOnly] > number) {
        //小数点分を切り上げてカウントしていて、.xx部分を越した。
        columnCount = columnCount - 1 + extractDecimalPart(number); //TODO:小数点部分のみを足すことで日分の小数点部分を再現
      }
      // シーン名と日数をスケジュールに注釈
      annotateSceneWithDays(
        rowIndex,
        columnIndex,
        sceneTitleOnly,
        columnCount,
        scheduleSheetDataValues
      );
    }
    prevSceneTitle = sceneTitle; // 前回のシーン名を更新
  }
  console.log("---displaySceneNameC end---");
}

// delimiterIndexのDELIMITER以降を削除する
function truncateTitle(sceneTitle, delimiterIndex) {
  // sceneTitle を DELIMITER で分割
  let parts = sceneTitle.split(DELIMITER);
  // 指定されたデリミタのインデックスまでの部分を抽出
  parts = parts.slice(0, delimiterIndex);
  // parts を再度結合して新しい title を作成
  let newTitle = parts.join(DELIMITER);
  // 新しい title を返す
  return newTitle;
}
//日分の前の値を取得する。
function extractDayNumberFromString(str) {
  // 日分の前にある数字（小数点を含む）を抽出する正規表現
  const matches = str.match(/(\d+(\.\d+)?)日分$/);
  return matches ? Number(matches[1]) : null;
}
//小数点以下の部分を取得する関数
function extractDecimalPart(number) {
  // 整数部分を取り出します。
  const integerPart = Math.trunc(number);
  // 元の数値から整数部分を引いて小数部分を得ます。
  const decimalPart = number - integerPart;
  return decimalPart;
}
// シーンがカバーする日数をカウントする関数
function countSceneDays(rowIndex, startColumnIndex, sheetData) {
  let count = 0;
  let targetSheetData = sheetData[rowIndex][startColumnIndex]; //最初にカウントするシーンを取得する。

  for (let i = startColumnIndex; i < sheetData[rowIndex].length; i++) {
    if (sheetData[rowIndex][i] === "") {
      // 空白セルが見つかったら終了
      break;
    }
    if (targetSheetData !== sheetData[rowIndex][i]) {
      //空白以外も違うシーンがあったら終了する。
      targetSheetData = sheetData[rowIndex][i];
      break;
    }
    count++;
  }
  return count; // シーンが続く日数を返す
}
// シーン名と日数をスケジュールに注釈する関数
function annotateSceneWithDays(
  rowIndex,
  columnIndex,
  sceneTitle,
  daysCount,
  scheduleData
) {
  if (sceneTitle !== "" && daysCount > 0) {
    // シーン名に日数を付け加える
    let annotation = sceneTitle + DELIMITER + daysCount + "日分";
    scheduleData[rowIndex][columnIndex] = annotation; // 注釈をシートに設定
  }
}

// 選択範囲のチェックと取得
// ロジックが複雑になるので選択範囲は
// シーン情報の１行のみ複数セル選択可能とする
function getSelectedRange(ranges, scheduleSheet) {
  // 選択領域のチェック
  if (ranges.length == 0) {
    let ui = SpreadsheetApp.getUi();
    ui.alert("選択されていません");
    return undefined;
  } else if (ranges.length > 1) {
    let ui = SpreadsheetApp.getUi();
    ui.alert("複数の選択領域があります");
    return undefined;
  }
  let range = ranges[0];
  // 選択行数は1行限定にする（便宜上）
  let selectedHight = range.getLastRow() + 1 - range.getRow();
  if (selectedHight != 1) {
    let ui = SpreadsheetApp.getUi();
    ui.alert("選択範囲が複数行です");
    return undefined;
  }

  // シーン行チェック
  // 値を取得
  const scheduleSheetDataVal = scheduleSheet.getDataRange().getValues();

  // "CLW美術作業者一覧" と "以下フリースぺース" の行番号を見つける
  let startRowIndex = scheduleSheetDataVal.findIndex(
    (row) => row[0] === CLW_ARTWORK_PERSONS_TITLE
  );
  let endRowIndex = scheduleSheetDataVal.findIndex(
    (row) => row[0] === FREE_SPACE_TITLE
  );
  let personRows = [];
  // シーン行取得
  if (startRowIndex !== -1 && endRowIndex !== -1) {
    for (let i = startRowIndex + 1; i < endRowIndex; i++) {
      if (
        scheduleSheetDataVal[i][PERSON_COLUMN_INDEX] != "" &&
        scheduleSheetDataVal[i][PERSON_COLUMN_INDEX] != MEMO_SPACE_TITLE
      ) {
        personRows.push(i);
      }
    }
  } else {
    // ダイアログを表示してエラーを通知
    let ui = SpreadsheetApp.getUi();
    ui.alert(
      "データ範囲エラー",
      "必要なデータの範囲が見つかりませんでした。スプレッドシートのフォーマットが正しいことを確認し、再度お試しください。",
      ui.ButtonSet.OK
    );
    // 処理を停止
    return undefined;
  }
  if (!personRows.includes(range.getRow() - 1)) {
    let ui = SpreadsheetApp.getUi();
    ui.alert("シーン行が選択されていません");
    return undefined;
  }
  // 選択範囲が日付の表示領域内チェック
  // カレンダー最後の日付の切り出し
  const lastDayIndex = scheduleSheetDataVal[
    SS_CALENDERDATE_ROW_INDEX
  ].findLastIndex((val) => val != "");
  if (
    range.getColumn() - 1 < SS_CALENDERDATE_COLUMN_INDEX ||
    lastDayIndex < range.getLastColumn() - 1
  ) {
    let ui = SpreadsheetApp.getUi();
    ui.alert("選択範囲が日付の表示領域から外れています");
    return undefined;
  }

  return range;
}

// スケジュール表情報取得
function getScheduleSheetInfoC(scheduleSheet, dataBaseSheet) {
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

// スケジュール表担当情報取得
function getScheduleSheetPersonInfoC(scheduleSheet, dataBaseSheet, rowNo) {
  // scheduleSheetPersonRange の範囲情報を取得
  var scheduleNumColumns = scheduleSheet.getLastColumn();
  // スプレッドシート上でデータが入力されている担当情報範囲を選択
  scheduleSheetPersonRange = scheduleSheet.getRange(
    rowNo,
    1,
    3,
    scheduleNumColumns
  );
  // 値を取得
  scheduleSheetDataValues = scheduleSheetPersonRange.getValues();
  // 背景色を取得する
  scheduleSheetAllBackGrounds = scheduleSheetPersonRange.getBackgrounds();

  // dataBaseSheet の範囲を scheduleSheetPersonRange の範囲で取得
  dataBaseSheetPersonRange = dataBaseSheet.getRange(
    rowNo,
    1,
    3,
    scheduleNumColumns
  );
  // 値を取得
  dataBaseSheetDataValues = dataBaseSheetPersonRange.getValues();
}

// 空白のセルを探して色を塗る
function findAndFillClearCellC(
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
    setSellC(
      rowIndex,
      checkColumnIndex,
      storyInfoColor,
      sceneTitle,
      manHourColor,
      ""
    );
    return true;
  }
}

// 空白のセルを探して色を塗る
function fillClearCellC(
  rowIndex,
  startColumnIndex,
  totalDays,
  storyInfoColor,
  sceneTitle,
  manHourColors
) {
  // 何セル移動できるかの判断用に、該当行の配列情報をコピーしておく
  let tmpRowBackGrounds = scheduleSheetAllBackGrounds[rowIndex].slice();

  // 移動幅
  let countMove = 0;
  for (let i = startColumnIndex; i < startColumnIndex + totalDays; i++) {
    //移動先セルが空白のセル以外:ずらす値を+１、iをインクリメントせず空白のセルになるまでmovecをずらす
    while (tmpRowBackGrounds[i + countMove] != COLOR_CLEAR) {
      countMove++;
      continue;
    }
    setSellC(
      rowIndex,
      i + countMove,
      storyInfoColor,
      sceneTitle,
      manHourColors[i - startColumnIndex],
      ""
    );
  }
}
//空白セルの挿入
function addBlankCellsC(addBlankRange, isAllData = true) {
  //選択範囲の幅(土日はカウントから外す)
  let startColumnIndex = addBlankRange.getColumn() - 1;
  let endColumnIndx = addBlankRange.getLastColumn() - 1;
  let rowIndex = isAllData ? addBlankRange.getRow() - 1 : 0;
  console.log("startColumnIndex=" + startColumnIndex);
  console.log("endColumnIndx=" + endColumnIndx);
  console.log("rowIndex=" + rowIndex);

  // 何セル移動できるかの判断用と移動のコピー元用に、該当行の配列情報をコピーしておく
  let tmpRowBackGrounds = scheduleSheetAllBackGrounds[rowIndex].slice();
  let tmpRowStatusBackGrounds =
    scheduleSheetAllBackGrounds[rowIndex + 1].slice();
  let tmpRowDataBaseSheetDataValues = dataBaseSheetDataValues[rowIndex].slice();
  let tmpRowStatusDataBaseSheetDataValues =
    dataBaseSheetDataValues[rowIndex + 1].slice();

  // 選択範囲幅
  const addBlanklength = endColumnIndx - (startColumnIndex - 1);
  // // 選択範囲内の休日数
  // const rangeHoliday = tmpRowBackGrounds
  //   .slice(startColumnIndex, endColumnIndx + 1)
  //   .filter((backGround) => backGround == COLOR_HOLIDAY).length;
  // // 見切れる範囲の休日数
  // const outHoliday = tmpRowBackGrounds
  //   .slice(tmpRowBackGrounds.length - addBlanklength)
  //   .filter((backGround) => backGround == COLOR_HOLIDAY).length;
  // // 見切れる範囲の固定セル数 //TODO:固定セルの数も足し合わせる
  // const outFixedDay = tmpRowStatusDataBaseSheetDataValues
  //   .slice(tmpRowStatusDataBaseSheetDataValues.length - addBlanklength)
  //   .filter((value) => value == FIXED_CELL_KEYWORD).length;

  // // 空白セルの挿入時に見切れる範囲に初期値・休日以外がある場合エラー
  // const outRange =
  //   tmpRowBackGrounds.length -
  //   1 -
  //   (addBlanklength - rangeHoliday + outHoliday + outFixedDay); //TODO:大畑さんに確認
  // //tmpRowBackGrounds.length - 1 - addBlanklength - rangeHoliday + outHoliday - outFixedDay;
  // const outRangeInput = tmpRowBackGrounds
  //   .slice(outRange)
  //   .filter(
  //     (backGround) => backGround != COLOR_HOLIDAY && backGround != COLOR_CLEAR
  //   ).length;
  // if (outRangeInput > 0) {
  //   let ui = SpreadsheetApp.getUi();
  //   ui.alert("見切れる範囲に入力があります");
  //   return;
  // }

  //選択開始列　から　一番最終列までループ（１行だけ指定と想定）
  let countMove = addBlanklength;
  for (let i = startColumnIndex; i < tmpRowBackGrounds.length; i++) {
    //該当セルが土日:ずらす値を-１、iをインクリメント //TODO:
    if (
      tmpRowBackGrounds[i] == COLOR_HOLIDAY ||
      tmpRowStatusDataBaseSheetDataValues[i] === FIXED_CELL_KEYWORD ||
      (tmpRowBackGrounds[i] == COLOR_CLEAR && i > endColumnIndx)
    ) {
      countMove--;
      if (countMove <= 0) {
        break;
      }
      continue;
      //}
    }
    //移動先セルが土日:ずらす値を+１、iをインクリメントせず土日でなくなるまでmovecをずらす//TODO:固定セルは無視する処理
    while (
      tmpRowBackGrounds[i + countMove] == COLOR_HOLIDAY ||
      tmpRowStatusDataBaseSheetDataValues[i + countMove] === FIXED_CELL_KEYWORD
    ) {
      //while (tmpRowBackGrounds[i + countMove] == COLOR_HOLIDAY) {
      countMove++;
    }

    if (i + countMove >= tmpRowBackGrounds.length) {
      throw new Error("見切れる範囲に入力があります");
    }
    setSellC(
      rowIndex,
      i + countMove,
      tmpRowBackGrounds[i],
      tmpRowDataBaseSheetDataValues[i],
      tmpRowStatusBackGrounds[i],
      "" //移動するということは、固定セルではない
    );
  }
  // 選択範囲をClear  //固定セルが選択されている場合は固定が消える。
  for (let i = startColumnIndex; i <= endColumnIndx; i++) {
    //該当セルが土日は何もしない
    if (tmpRowBackGrounds[i] == COLOR_HOLIDAY) {
      continue;
    }
    //値、色クリア
    setSellC(rowIndex, i, COLOR_CLEAR, "", COLOR_CLEAR, "");
  }
}

//セル削除（削除後詰めあり）
function deleteCellsWithMove(deleteRange, isAllData = true) {
  //選択範囲の幅(土日はカウントから外す)
  let startColumnIndex = deleteRange.getColumn() - 1;
  let endColumnIndx = deleteRange.getLastColumn() - 1;
  let rowIndex = isAllData ? deleteRange.getRow() - 1 : 0;
  console.log("startColumnIndex=" + startColumnIndex);
  console.log("endColumnIndx=" + endColumnIndx);
  console.log("rowIndex=" + rowIndex);

  // 移動のコピー元用に、該当行の配列情報をコピーしておく
  let tmpRowBackGrounds = scheduleSheetAllBackGrounds[rowIndex].slice();
  let tmpRowStatusBackGrounds =
    scheduleSheetAllBackGrounds[rowIndex + 1].slice();
  let tmpRowDataBaseSheetDataValues = dataBaseSheetDataValues[rowIndex].slice();
  let tmpRowStatusDataBaseSheetDataValues =
    dataBaseSheetDataValues[rowIndex + 1].slice();
  let countMove = endColumnIndx - (startColumnIndex - 1);
  for (let i = endColumnIndx + 1; i < tmpRowBackGrounds.length; i++) {
    //該当セルが土日:ずらす値を＋１、iをインクリメント//TODO:固定セルは無視する処理
    if (
      tmpRowBackGrounds[i] == COLOR_HOLIDAY ||
      tmpRowStatusDataBaseSheetDataValues[i] === FIXED_CELL_KEYWORD
    ) {
      countMove++;
      continue;
    }
    //移動先セルが土日:ずらす値を-１、iをインクリメントせず土日でなくなるまでmovecをずらす//TODO:固定セルは無視する処理
    while (
      tmpRowBackGrounds[i - countMove] == COLOR_HOLIDAY ||
      tmpRowStatusDataBaseSheetDataValues[i - countMove] === FIXED_CELL_KEYWORD
    ) {
      countMove--;
    }

    setSellC(
      rowIndex,
      i - countMove,
      tmpRowBackGrounds[i],
      tmpRowDataBaseSheetDataValues[i],
      tmpRowStatusBackGrounds[i],
      ""
    );
  }
  // 右端をClear
  for (let i = countMove; i > 0; i--) {
    //該当セルが土日は何もしない
    if (tmpRowBackGrounds[tmpRowBackGrounds.length - i] == COLOR_HOLIDAY) {
      continue;
    }
    //値、色クリア
    setSellC(
      rowIndex,
      tmpRowBackGrounds.length - i,
      COLOR_CLEAR,
      "",
      COLOR_CLEAR,
      ""
    );
  }
}

function updateScheduleSheetWithDataValuesC() {
  // 全行の列数を確認する
  scheduleSheetDataValues.forEach((row, index) => {
    console.log(`行${index + 1}の列数:`, row.length);
    if (row.length !== scheduleSheetAllRange.getNumColumns()) {
      throw new Error(
        `行${index + 1}の列数がscheduleSheetAllRangeの列数と一致しません。`
      );
    }
  });

  // 以降の処理
  scheduleSheetAllRange.setValues(scheduleSheetDataValues);
  scheduleSheetAllRange.setBackgrounds(scheduleSheetAllBackGrounds);
  dataBaseSheetAllRange.setValues(dataBaseSheetDataValues);
}

function updateScheduleSheetPresonWithDataValuesC() {
  // 全行の列数を確認する
  scheduleSheetDataValues.forEach((row, index) => {
    console.log(`行${index + 1}の列数:`, row.length);
    if (row.length !== scheduleSheetPersonRange.getNumColumns()) {
      throw new Error(
        `行${index + 1}の列数がscheduleSheetAllRangeの列数と一致しません。`
      );
    }
  });

  // 以降の処理
  scheduleSheetPersonRange.setValues(scheduleSheetDataValues);
  scheduleSheetPersonRange.setBackgrounds(scheduleSheetAllBackGrounds);
  dataBaseSheetPersonRange.setValues(dataBaseSheetDataValues);
}
//スケジュール表の選択セルが固定セルかどうかを確認する。
function isCellFixed(values, rowIndex, columnIndex) {
  let cellStatus = values[rowIndex + 1][columnIndex]; // 次の行の同じ列に固定セルの状態があるかを確認
  return cellStatus === FIXED_CELL_KEYWORD;
}
