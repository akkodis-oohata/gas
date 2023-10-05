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
  const SCENE_END_BODY_COLUMN_NUM = 17;
  const SCENE_START_MANHOUR_COLUMN_NUM = 21;
  const SCENE_END_MANHOUR_COLUMN_NUM = 170;
  const SCENENAME_COLUMN_NUM = 2;

  //全工数表
  //全工数表シート名
  const MANHOUR_ALL_SHEET_NAME = "全工数表";
  //タスク描画開始地点
  const MANHOUR_ALL_START_ROW_NUM = 4;
  const MANHOUR_ALL_START_COLUMN_NUM = 1;
  const SMANHOUR_ALL_START_BODY_COLUMN_NUM = 2;

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
    // クリア行範囲：データ最終行 - ヘッダ行 + 1
    const clearRowRange =
      manhour_all_sheet.getLastRow() - MANHOUR_ALL_START_ROW_NUM + 1;
    // クリア列範囲：開始日列 - 開始列 + 1
    const clearColumnRange =
      manhour_all_sheet.getLastColumn() - MANHOUR_ALL_START_COLUMN_NUM + 1;
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
    // コピー先開始行　forEach内で+ 1で計算するため - 1調整
    let paste_endRow = MANHOUR_ALL_START_ROW_NUM - 1;
    progress_sheets.forEach((progress_sheet) => {
      let progressRowRange = 0;
      try {
        // コピー行範囲：B列（シーン名）データのヘッダ行以降のデータ最終行 - ヘッダ行 + 1
        progressRowRange =
          progress_sheet
            .getRange(SCENE_START_BODY_ROW_NUM, SCENENAME_COLUMN_NUM)
            .getNextDataCell(SpreadsheetApp.Direction.DOWN)
            .getRow() -
          SCENE_START_BODY_ROW_NUM +
          1;
      } catch (err) {
        return;
      }
      if (progressRowRange == 0) return;

      // コピー列範囲：開始日列 - 開始列 + 1
      const progressBodyColumnRange =
        SCENE_END_BODY_COLUMN_NUM - SCENE_START_BODY_COLUMN_NUM + 1;

      // シーン備考～開始日のコピー
      let copyRange = progress_sheet.getRange(
        SCENE_START_BODY_ROW_NUM,
        SCENE_START_BODY_COLUMN_NUM,
        progressRowRange,
        progressBodyColumnRange
      );
      let pasteRange = manhour_all_sheet.getRange(
        paste_endRow + 1,
        SMANHOUR_ALL_START_BODY_COLUMN_NUM,
        progressRowRange,
        progressBodyColumnRange
      );
      copyRange.copyTo(pasteRange);

      // 工数のコピー
      copyRange = progress_sheet.getRange(
        SCENE_START_BODY_ROW_NUM,
        SCENE_START_MANHOUR_COLUMN_NUM,
        progressRowRange,
        SCENE_END_MANHOUR_COLUMN_NUM - SCENE_START_MANHOUR_COLUMN_NUM + 1
      );
      pasteRange = manhour_all_sheet.getRange(
        paste_endRow + 1,
        SMANHOUR_ALL_START_BODY_COLUMN_NUM + progressBodyColumnRange,
        progressRowRange,
        SCENE_END_MANHOUR_COLUMN_NUM - SCENE_START_MANHOUR_COLUMN_NUM + 1
      );
      copyRange.copyTo(pasteRange);

      // 作品話数のコピー
      copyRange = progress_sheet.getRange(
        SCENE_TITLE_ROW_NUM,
        SCENE_TITLE_COLUMN
      );
      pasteRange = manhour_all_sheet.getRange(
        paste_endRow + 1,
        MANHOUR_ALL_START_COLUMN_NUM,
        progressRowRange,
        1
      );
      copyRange.copyTo(pasteRange);
      pasteRange.setFontWeight("normal").setFontSize(10);

      paste_endRow += progressRowRange;
    });
  }
}
