//共通変数定義（ファイル間で利用する変数の定義）
//担当者
const PERSON_COLUMN_INDEX = 6;

function initializeDataSpaceMain(scheduleSheetAllRange, scheduleSheetDataValues,
    scheduleSheetAllBackGrounds) {
  console.log("initializeDataSpaceMain in");

  initializeDataSpaceCommon(scheduleSheetAllRange, scheduleSheetDataValues,
    scheduleSheetAllBackGrounds)

  console.log("initializeDataSpaceMain out");
}

{
  //データスペース初期化(以下データ用スペース(最大1000行)より下全削除)
  function initializeDataSpaceCommon(scheduleSheetAllRange, scheduleSheetDataValues,
    scheduleSheetAllBackGrounds){
    console.log("initializeDataSpaceCommon in");

    var startRowIndex = scheduleSheetDataValues.findIndex(row => row[0] === "以下データ用スペース(最大1000行)") + 1;
    
    // 行を削除する
    scheduleSheetDataValues.splice(startRowIndex, Infinity);
    
    // 新しい行を作成する
    var personBaseRow  = new Array(scheduleSheetDataValues[0].length).fill("");  // 新しい行を作成し、全てのセルを空文字列で初期化する
    personBaseRow [PERSON_COLUMN_INDEX] = "作業者ベースデータ";  // "作業者ベースデータ"を設定する
    
    // 削除した行の位置に新しい行を挿入する
    scheduleSheetDataValues.splice(startRowIndex, 0, personBaseRow);

    var newOtherCompanyRow  = new Array(scheduleSheetDataValues[0].length).fill("");  // 新しい行を作成し、全てのセルを空文字列で初期化する
    newOtherCompanyRow [PERSON_COLUMN_INDEX] = "他社ベースデータ";  // "他社ベースデータ"を設定する
    
    // 削除した行の位置に新しい行を挿入する
    scheduleSheetDataValues.splice(startRowIndex + 1, 0, newOtherCompanyRow);
    // タイトルの配列を作成します
    var titles = ["作品話数ベースデータ", "作品名", "作品色", "シーン名", "担当者", "開始日", "未入り", "未撒き", "撒済＋再調整", "未入り色", "未撒色", "撒済＋再調整色", "進行表有無"];
    // 新しい行を作成し、全てのセルを空文字列で初期化します
    var newStoryRow = new Array(scheduleSheetDataValues[0].length).fill("");
    // タイトルの配列をループし、各タイトルを適切なセルに設定します
    for (var i = 0; i < titles.length; i++) {
      newStoryRow[PERSON_COLUMN_INDEX + i] = titles[i];
    }
    // 新しい行をscheduleSheetDataValues配列に挿入します
    scheduleSheetDataValues.splice(startRowIndex + 2, 0, newStoryRow);

    // scheduleSheetAllBackGroundsの作業ベースデータ～作品話数ベースデータの一つ上までの初期化
    // 同様に行を削除し、新しい行を挿入する
    scheduleSheetAllBackGrounds.splice(startRowIndex, Infinity);
    var personBaseBackgroundRow = new Array(scheduleSheetAllBackGrounds[0].length).fill("");  // 新しい背景色行を作成し、全てのセルを空文字列で初期化する
    scheduleSheetAllBackGrounds.splice(startRowIndex, 0, personBaseBackgroundRow);
    scheduleSheetAllBackGrounds.splice(startRowIndex + 1, 0, personBaseBackgroundRow);

    console.log("initializeDataSpaceCommon out");

  }

}