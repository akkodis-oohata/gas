function updateManHoursAllSheetMain() {
  updateManHoursAllSheet();
}

{
  //-----------------
  //定数定義
  //-----------------
  //進行表
  //進行表シート名
  const PROGRESS_SHEET_NAME = "進行表";
  const PROGRESS_SHEET_BLANKNAME = "進行表フォーマット(白紙)  ";
  //シーン情報
  const SCENE_TITLE_ROW_NUM = 1;
  const SCENE_TITLE_COLUMN = 1;
  const SCENE_START_BODY_ROW_NUM = 7;
  const SCENE_START_BODY_COLUMN_NUM = 1;
  const SCENE_END_BODY_ROW_NUM = 60; //適当な値
  const SCENE_END_BODY_COLUMN_NUM = 17;
  const SCENE_START_MANHOUR_COLUMN_NUM = 21;
  const SCENE_END_MANHOUR_COLUMN_NUM = 170;
  const SCENENAME_COLUMN_NUM = 2;

  //全工数表
  //全工数表シート名
  const MANHOUR_ALL_SHEET_NAME = "全工数表";
  //描画開始地点
  const MANHOUR_ALL_START_ROW_NUM = 4;
  const MANHOUR_ALL_START_COLUMN_NUM = 1;
  //タスク描画終了地点
  const MANHOUR_ALL_BODY_COLUMN_NUM = 18;
  // ボーダー点線位置
  const BORDER_DOTTED_START1 = 9;
  const BORDER_DOTTED_END1 = 10;
  const BORDER_DOTTED_START2 = 13;
  const BORDER_DOTTED_END2 = 14;

  //-----------------
  //関数
  //-----------------
  //メインの関数
  function updateManHoursAllSheet() {
    // スプレッドシートの読み込み
    let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // 進行表読み込み
    const progress_sheets = spreadsheet
      .getSheets()
      .filter(
        (sheet) =>
          sheet.getName().indexOf(PROGRESS_SHEET_NAME) > -1 &&
          sheet.getName() != PROGRESS_SHEET_BLANKNAME
      );

    // 全工数表読み込み
    const manhour_all_sheet = spreadsheet.getSheetByName(
      MANHOUR_ALL_SHEET_NAME
    );

    // 工数設定
    setManhour(progress_sheets, manhour_all_sheet);
  }

  //-----------------
  // 工数設定
  //-----------------
  // 進行表のデータを全工数表にコピーする
  function setManhour(progress_sheets, manhour_all_sheet) {
    // 全工数表のデータを全て取得
    const allDataRange = manhour_all_sheet.getDataRange();
    // 値をを取得
    const manhourAllValues = allDataRange.getValues();
    // クリア行範囲：データ最終行 - ヘッダ行 + 1
    const clearRowRange =
      manhourAllValues.length - MANHOUR_ALL_START_ROW_NUM + 1;
    // クリア列範囲：開始日列 - 開始列 + 1
    const clearColumnRange =
      manhourAllValues[0].length - MANHOUR_ALL_START_COLUMN_NUM + 1;
    // シートのクリア
    if (clearRowRange > 0) {
      manhour_all_sheet
        .getRange(
          MANHOUR_ALL_START_ROW_NUM,
          MANHOUR_ALL_START_COLUMN_NUM,
          clearRowRange,
          clearColumnRange
        )
        .clear();
    }
    // コピー先変数
    let pasteVal = [];
    let pasteBg = [];
    progress_sheets.forEach((progress_sheet) => {
      // 進行表のデータを取得
      const progressDataRange = progress_sheet.getRange(
        SCENE_TITLE_ROW_NUM,
        SCENE_TITLE_COLUMN,
        SCENE_END_BODY_ROW_NUM - SCENE_TITLE_ROW_NUM + 1,
        SCENE_END_MANHOUR_COLUMN_NUM - SCENE_TITLE_ROW_NUM + 1
      );
      // 値・背景色をを取得
      const progressAllValues = progressDataRange.getValues();
      const progressAllBackGrounds = progressDataRange.getBackgrounds();

      // コピー行範囲：B列（シーン名）データのヘッダ行以降のデータ最終行
      const progressRowIndex = progressAllValues.findIndex(
        (data, index) =>
          data[SCENENAME_COLUMN_NUM - 1] === "" &&
          index >= SCENE_START_BODY_ROW_NUM - 1
      );
      const progressRowRange = progressRowIndex - SCENE_START_BODY_ROW_NUM + 1;

      if (progressRowRange == 0) return;

      // シーン備考～開始日のコピー
      const copyValSheen = progressAllValues
        .slice(SCENE_START_BODY_ROW_NUM - 1, progressRowIndex)
        .map((dataRow) => {
          return dataRow.slice(
            SCENE_START_BODY_COLUMN_NUM - 1,
            SCENE_END_BODY_COLUMN_NUM
          );
        });
      const copyBgSheen = progressAllBackGrounds
        .slice(SCENE_START_BODY_ROW_NUM - 1, progressRowIndex)
        .map((dataRow) => {
          return dataRow.slice(
            SCENE_START_BODY_COLUMN_NUM - 1,
            SCENE_END_BODY_COLUMN_NUM
          );
        });

      // 工数のコピー
      const copyValManHour = progressAllValues
        .slice(SCENE_START_BODY_ROW_NUM - 1, progressRowIndex)
        .map((dataRow) => {
          return dataRow.slice(
            SCENE_START_MANHOUR_COLUMN_NUM - 1,
            SCENE_END_MANHOUR_COLUMN_NUM
          );
        });
      const copyBgManHour = progressAllBackGrounds
        .slice(SCENE_START_BODY_ROW_NUM - 1, progressRowIndex)
        .map((dataRow) => {
          return dataRow.slice(
            SCENE_START_MANHOUR_COLUMN_NUM - 1,
            SCENE_END_MANHOUR_COLUMN_NUM
          );
        });

      // 作品話数の列範囲分コピー
      const copyValTitle = Array(progressRowRange).fill(
        progressAllValues[SCENE_TITLE_ROW_NUM - 1][SCENE_TITLE_COLUMN - 1]
      );
      const copyBgTitle = Array(progressRowRange).fill(
        progressAllBackGrounds[SCENE_TITLE_ROW_NUM - 1][SCENE_TITLE_COLUMN - 1]
      );

      // 各コピー範囲を結合して格納
      pasteVal.push(
        ...copyValTitle.map((title, index) => {
          const row = [title];
          row.push(...copyValSheen[index], ...copyValManHour[index]);
          return row;
        })
      );
      pasteBg.push(
        ...copyBgTitle.map((title, index) => {
          const row = [title];
          row.push(...copyBgSheen[index], ...copyBgManHour[index]);
          return row;
        })
      );
    });

    // 貼り付け範囲取得
    const pastDataRange = manhour_all_sheet.getRange(
      MANHOUR_ALL_START_ROW_NUM,
      MANHOUR_ALL_START_COLUMN_NUM,
      pasteVal.length,
      pasteVal[0].length
    );

    // 書式・入力規則を貼り付け範囲に適応（一行目をコピー）
    pastDataRange.offset(0, 0, 1, pasteVal[0].length).copyTo(pastDataRange);
    // 値・背景色
    pastDataRange.setValues(pasteVal);
    pastDataRange.setBackgrounds(pasteBg);
    // 罫線
    let boderDataRange = manhour_all_sheet.getRange(
      MANHOUR_ALL_START_ROW_NUM,
      MANHOUR_ALL_START_COLUMN_NUM,
      pasteVal.length,
      MANHOUR_ALL_BODY_COLUMN_NUM - MANHOUR_ALL_START_COLUMN_NUM + 1
    );
    boderDataRange.setBorder(true, true, true, true, true, true);

    // 罫線-点線1
    boderDataRange = manhour_all_sheet.getRange(
      MANHOUR_ALL_START_ROW_NUM,
      BORDER_DOTTED_START1,
      pasteVal.length,
      BORDER_DOTTED_END1 - BORDER_DOTTED_START1 + 1
    );
    boderDataRange.setBorder(
      null,
      true,
      null,
      true,
      null,
      null,
      null,
      SpreadsheetApp.BorderStyle.DASHED
    );

    // 罫線-点線2
    boderDataRange = manhour_all_sheet.getRange(
      MANHOUR_ALL_START_ROW_NUM,
      BORDER_DOTTED_START2,
      pasteVal.length,
      BORDER_DOTTED_END2 - BORDER_DOTTED_START2 + 1
    );
    boderDataRange.setBorder(
      null,
      true,
      null,
      true,
      null,
      null,
      null,
      SpreadsheetApp.BorderStyle.DASHED
    );
  }
}