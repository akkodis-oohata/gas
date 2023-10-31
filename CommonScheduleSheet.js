//-----------------
//定数定義
//-----------------
const DATA_BASE_SHEET_NAME = "データベース"
//セルの色塗り初期化時の色
const COLOR_CLEAR = "#ffffff";
// 土日祝日の色
const COLOR_HOLIDAY = "#808080";
// データベースとスケジュール表のフォーマットの差を埋める為、美術ルーム全体スケジュールmemoの行数分+CLW美術作業者一覧分-以下データスペース-作業者ベースデータで算出する
const DATA_BASE_FORMAT_OFFSET = 5;
//カレンダー
const SS_CALENDERDATE_COLUMN_INDEX = 7;
//表示最大日数
const SS_MAXDAYS = 730; //730(24か月)
// 区切り文字
const DELIMITER = ",";


//-----------------
//変数
//-----------------
// スプレッドシート上でデータが入力されている最大範囲を選択
let scheduleSheetAllRange = undefined
//スケジュール表の全値取得
let scheduleSheetDataValues = undefined
//スケジュール表の背景色取得
let scheduleSheetAllBackGrounds = undefined
// データベースシートに格納する為、データ領域を格納する配列
let dataBaseSheetDataValues = undefined;
// データベースシートの最大範囲
let dataBaseSheetAllRange = undefined;

//セル削除（削除後詰めなし）
function deleteCellsC(range) {

  let rowIndex = range.getRow() - 1;
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
    setSellC(rowIndex, index, COLOR_CLEAR, "", COLOR_CLEAR);
    console.log("rowIndex:"+rowIndex+"  "+"index:"+index)
    
    
  }

}

//セルを隣に移動 
function moveCellC(rowIndex, index, existColor, sceneTitle, manHourColor) {
  let nextExistColor = scheduleSheetAllBackGrounds[rowIndex][index];
  //土日祝日だけ割り込み禁止
  if (nextExistColor == COLOR_HOLIDAY) {
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
    nextDataBaseSceneTitle = dataBaseSheetDataValues[rowIndex - DATA_BASE_FORMAT_OFFSET][index];
  }
  setSellC(rowIndex, index, existColor, sceneTitle, manHourColor);
  if (nextExistColor != COLOR_CLEAR) {
    index++;
    if (nextSceneTitle != undefined && nextSceneTitle != "") {
      moveCellC(rowIndex, index, nextExistColor, nextSceneTitle, nextManHourColor);
    } else {
      moveCellC(
        rowIndex,
        index,
        nextExistColor,
        nextDataBaseSceneTitle,
        nextManHourColor
      );
    }
  }
}

//セルに値を設定。表示用シーン名はクリア（後で一気に付ける）
function setSellC(rowIndex, columnIndex, setColor, sceneTitle, manHourColor) {
  scheduleSheetDataValues[rowIndex][columnIndex] = "";
  scheduleSheetAllBackGrounds[rowIndex][columnIndex] = setColor;
  scheduleSheetAllBackGrounds[rowIndex + 1][columnIndex] = manHourColor;
  dataBaseSheetDataValues[rowIndex - DATA_BASE_FORMAT_OFFSET][columnIndex] = sceneTitle;
}

//先頭セルにシーン名を表示する
function displaySceneNameC(rowIndex) {
  let prevSceneTitle = "";
  for (
    let columnIndex = SS_CALENDERDATE_COLUMN_INDEX;
    columnIndex < SS_CALENDERDATE_COLUMN_INDEX + SS_MAXDAYS;
    columnIndex++
  ) {
    let sceneTitle =
      dataBaseSheetDataValues[rowIndex - DATA_BASE_FORMAT_OFFSET][columnIndex];
    // sceneTitleには、開始日が含まれている為削除する。
    sceneTitle = truncateTitle(sceneTitle, 3)
    if (sceneTitle != prevSceneTitle) {
      scheduleSheetDataValues[rowIndex][columnIndex] = sceneTitle;
    }
    prevSceneTitle = sceneTitle;
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
}

// 選択範囲のチェックと取得
// ロジックが複雑になるので選択範囲は
// シーン情報の１行のみ複数セル選択可能とする
// TODO シーン行が選択されている、というチェック入れる
// TODO 選択範囲が日付の表示領域内、というチェック入れる
function getSelectedRange(ranges) {
  // 選択領域のチェック
  if (ranges.length == 0) {
    let ui = SpreadsheetApp.getUi();
    ui.alert('選択されていません');
    return undefined;
  } else if (ranges.length > 1) {
    let ui = SpreadsheetApp.getUi();
    ui.alert('複数の選択領域があります');
    return undefined;
  }
  let range = ranges[0];
  // 選択行数は1行限定にする（便宜上）
  let selectedHight = range.getLastRow() + 1 - range.getRow()
  if (selectedHight != 1) {
    let ui = SpreadsheetApp.getUi();
    ui.alert('選択範囲が複数行です');
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
  dataBaseSheetAllRange = dataBaseSheet.getRange(1, 1, scheduleNumRows, scheduleNumColumns);
  // 値を取得
  dataBaseSheetDataValues = dataBaseSheetAllRange.getValues();

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
      manHourColor
    );
    return true;
  }
}

//空白セルの挿入
function addBlankCellsC(addBlankRange) {

  //選択範囲の幅(土日はカウントから外す)
  let startColumnIndex = addBlankRange.getColumn() - 1;
  let endColumnIndx = addBlankRange.getLastColumn() - 1;
  let rowIndex = addBlankRange.getRow() - 1;
  console.log("startColumnIndex=" + startColumnIndex);
  console.log("endColumnIndx=" + endColumnIndx);
  console.log("rowIndex=" + rowIndex);
  let countWidth = 0;
  for (let i = startColumnIndex; i <= endColumnIndx; i++) {
    if (scheduleSheetAllBackGrounds[rowIndex][i] != COLOR_HOLIDAY) {
      countWidth++;
    }
  }
  //console.log("countWidth=" + countWidth);
  // 何セル移動できるかの判断用と移動のコピー元用に、該当行の配列情報をコピーしておく
  let tmpRowBackGrounds = scheduleSheetAllBackGrounds[rowIndex].slice();
  let tmpRowStatusBackGrounds = scheduleSheetAllBackGrounds[rowIndex + 1].slice();
  let tmpRowDataBaseSheetDataValues = dataBaseSheetDataValues[rowIndex - DATA_BASE_FORMAT_OFFSET].slice();
  //console.log("tmpRowBackGrounds.length=" + tmpRowBackGrounds.length);
  //選択開始列　から　一番最終列までループ（１行だけ指定と想定）
  for (let i = startColumnIndex; i < tmpRowBackGrounds.length; i++) {
    //該当セルが土日は何もしない
    if (tmpRowBackGrounds[i] == COLOR_HOLIDAY) {
      continue;
    }

    //選択範囲の場合
    if (i >= startColumnIndex && i <= endColumnIndx) {
      //Clearする（空白設定）
      setSellC(rowIndex, i, COLOR_CLEAR, "", COLOR_CLEAR);
    }

    //移動距離の計算（土日は飛び越える）
    let countMove = 0;
    let startMoveIndex = i + 1;
    let overMax = false;
    for (let j = startMoveIndex, tmpIndex = startMoveIndex;
      j < startMoveIndex + countWidth; j++, tmpIndex++) {
      countMove++;
      //移動した場所がMAXを超えていたら移動しない
      if (tmpIndex >= tmpRowBackGrounds.length) {
        overMax = true;
        continue;
      }
      while (tmpRowBackGrounds[tmpIndex] == COLOR_HOLIDAY) {
        countMove++;
        tmpIndex++;
      }
    }
    //オリジナルの配列で移動距離足した分の場所に値をコピーする
    if (!overMax) {
      setSellC(rowIndex, i + countMove, tmpRowBackGrounds[i], tmpRowDataBaseSheetDataValues[i], tmpRowStatusBackGrounds[i]);
    }

  }


}

//セル削除（削除後詰めあり）
function deleteCellsWithMove(deleteRange) {

  //選択範囲の幅(土日はカウントから外す)
  let startColumnIndex = deleteRange.getColumn() - 1;
  let endColumnIndx = deleteRange.getLastColumn() - 1;
  let rowIndex = deleteRange.getRow() - 1;
  console.log("startColumnIndex=" + startColumnIndex);
  console.log("endColumnIndx=" + endColumnIndx);
  console.log("rowIndex=" + rowIndex);
  let countWidth = 0;
  for (let i = startColumnIndex; i <= endColumnIndx; i++) {
    if (scheduleSheetAllBackGrounds[rowIndex][i] != COLOR_HOLIDAY) {
      countWidth++;
    }
  }
  // 何セル移動できるかの判断用と移動のコピー元用に、該当行の配列情報をコピーしておく
  let tmpRowBackGrounds = scheduleSheetAllBackGrounds[rowIndex].slice();
  let tmpRowStatusBackGrounds = scheduleSheetAllBackGrounds[rowIndex + 1].slice();
  let tmpRowDataBaseSheetDataValues = dataBaseSheetDataValues[rowIndex - DATA_BASE_FORMAT_OFFSET].slice();
  //選択開始列　から　一番最終列までループ（１行だけ指定と想定）
  for (let i = startColumnIndex; i < tmpRowBackGrounds.length; i++) {
    //該当セルが土日は何もしない
    if (tmpRowBackGrounds[i] == COLOR_HOLIDAY) {
      continue;
    }

    //選択範囲の場合
    //TODO 一番右端をClearする必要ある
    if (i >= startColumnIndex && i <= endColumnIndx) {
       //Clearする（空白設定）
       //setSellC(rowIndex, i, COLOR_CLEAR, "", COLOR_CLEAR);
       continue;
    }

    //移動距離の計算（土日は飛び越える）
    let countMove = 0;
    let startMoveIndex = i - 1;
    
    for (let j = startMoveIndex, tmpIndex = startMoveIndex;
      j > startMoveIndex - countWidth; j--, tmpIndex--) {
      countMove++;
      while (tmpRowBackGrounds[tmpIndex] == COLOR_HOLIDAY) {
        countMove++;
        tmpIndex--;
      }
    }
    setSellC(rowIndex, i - countMove, tmpRowBackGrounds[i], tmpRowDataBaseSheetDataValues[i], tmpRowStatusBackGrounds[i]);

  }


}



//スケジュール表の画面を更新する
function updateScheduleSheetWithDataValuesC() {
  scheduleSheetAllRange.setValues(scheduleSheetDataValues);
  scheduleSheetAllRange.setBackgrounds(scheduleSheetAllBackGrounds);
  dataBaseSheetAllRange.setValues(dataBaseSheetDataValues);

}