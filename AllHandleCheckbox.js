// 編集時にAllHandleCheckbox関数をトリガーします。
function onEdit(e) {
  try {
    allHandleCheckbox(e);
  } catch (error) {
    let ui = SpreadsheetApp.getUi();
    ui.alert(error.message);
  }
}

{
  const ALL_WARIKOMI_CHECKBOX_ROW = 3;
  const ALL_WARIKOMI_CHECKBOX_COLUMN = 18;
  const ALL_REPLACE_CHECKBOX_ROW = 3;
  const ALL_REPLACE_CHECKBOX_COLUMN = 20;
  const SCENE_START_BODY_ROW_NUM = 7;

  // 特定のチェックボックスの状態に応じてhandleCheckbox関数を呼び出します。
  function allHandleCheckbox(e){
    var range = e.range;
    var row = range.getRow();
    var col = range.getColumn();
  
    if (row == ALL_WARIKOMI_CHECKBOX_ROW && col == ALL_WARIKOMI_CHECKBOX_COLUMN) {
      handleCheckbox(e, ALL_WARIKOMI_CHECKBOX_COLUMN);
    } else if (row == ALL_REPLACE_CHECKBOX_ROW && col == ALL_REPLACE_CHECKBOX_COLUMN) {
      handleCheckbox(e, ALL_REPLACE_CHECKBOX_COLUMN);
    }
  }

  // 指定された列のチェックボックスの状態を同期します。
  function handleCheckbox(e, targetColumn) {
    var sheet = e.source.getActiveSheet();
    var value = e.range.getValue();
    var firstEmptyRow = countRowsWithCheckboxes(sheet, targetColumn);
    var numRows = firstEmptyRow - SCENE_START_BODY_ROW_NUM;
  
    // numRowsが負でないことを確認
    if (numRows > 0) {
      // 値の2次元配列を作成
      var valuesArray = Array(numRows).fill([value]);
  
      // 一度に複数のセルの値を設定
      sheet.getRange(SCENE_START_BODY_ROW_NUM, targetColumn, numRows, 1).setValues(valuesArray);
    }
  }
  
  // 指定された列にあるチェックボックスの数をカウントします。
  function countRowsWithCheckboxes(sheet, col) {
    var lastRow = sheet.getLastRow();
    var dataValidations = sheet.getRange(SCENE_START_BODY_ROW_NUM, col, lastRow - SCENE_START_BODY_ROW_NUM + 1).getDataValidations();
    var count = 0;
    for (var i = 0; i < dataValidations.length; i++) {
      var validation = dataValidations[i][0];
      if (validation != null && validation.getCriteriaType() == SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
        count++;
      } else {
        // 空のセルを見つけたらループを終了
        break;
      }
    }
    return count + SCENE_START_BODY_ROW_NUM;
  }
}
