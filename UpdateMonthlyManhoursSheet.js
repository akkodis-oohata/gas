//月工数表の処理を記入

function updateMonthlyManhoursSheetMain() {
  exclusiveMain(updateMonthlyManhoursSheet);
}

{
  let monthlyManhoursSheet = undefined;

  function updateMonthlyManhoursSheet() {
    let dataBaseSheetDataValues = getAllDataFromDatabaseSheet();
    //スケジュール表より月を取得する
    let months = extractAndSortUniqueMonthsFromDates(dataBaseSheetDataValues);
    console.log(months);
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    monthlyManhoursSheet = spreadsheet.getSheetByName("月工数表_検討中");
    //月工数表に記入する。
    //月列を記入する。
    writeToMonthlyManhoursSheet(months);

    // 作業可能工数
    writeManhours(months);
    // 全作品工数一覧
    writeTitlehours(months);
    // 話数工数一覧
    writeScenehours(months);
  }

  // 'データベース' シートの全データを取得する
  function getAllDataFromDatabaseSheet() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName("データベース");
    const data = sheet.getDataRange().getValues();
    return data;
  }

  // 7行目のH列から空文字列が見つかるまでのデータを取得し、重複のない月を古い順に並べ替えて抽出する
  function extractAndSortUniqueMonthsFromDates(dataBaseSheetDataValues) {
    let data = dataBaseSheetDataValues;
    let row = data[6]; // 7行目（配列は0から始まるため6を指定）
    let monthsSet = new Set();

    for (let i = 7; i < row.length; i++) {
      // H列は8番目の列（0から始まるので7を指定）
      if (row[i] === "") break; // 空文字列が見つかったら終了
      let date = new Date(row[i]);
      let month = Utilities.formatDate(
        date,
        Session.getScriptTimeZone(),
        "yy/MM"
      );
      monthsSet.add(month);
    }

    let sortedMonths = Array.from(monthsSet).sort((a, b) => {
      // 'yy/MM' 形式の文字列を日付オブジェクトに変換して比較
      let dateA = new Date(a.substring(0, 2), a.substring(3) - 1);
      let dateB = new Date(b.substring(0, 2), b.substring(3) - 1);
      return dateA - dateB; // ここを変更して古い順に並び替え
    });

    return sortedMonths;
  }

  // 月工数表の5行目E列から始まり、monthsの数だけ列にデータを記入する
  function writeToMonthlyManhoursSheet(monthData) {
    const months = monthData;
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName("月工数表_検討中");
    const startRow = 5;
    const startColumn = 5; // E列

    months.forEach((month, index) => {
      sheet.getRange(startRow, startColumn + index).setValue(month);
    });
  }

  // 作業可能工数
  function writeManhours(monthData) {
    const startRow = 6;
    const startColumn = 5; // E列

    // 値を取得
    const values = monthlyManhoursSheet.getDataRange().getValues();
    //スタッフ作業可能日：合計行の取得
    const rowIndex = values.findIndex(
      (row) => row[3] == "スタッフ作業可能日：合計"
    );
    // 作業可能工数スタッフ部分の取得
    const staffRange = monthlyManhoursSheet.getRange(
      startRow,
      1,
      rowIndex + 1 - startRow,
      startColumn + monthData.length - 1
    );
    const staffValue = staffRange.getValues();
    // 一般スタッフの値更新
    const updateStaffValue = staffValue.map((row) => {
      let updateRow = row;
      if (row[2] == "一般") {
        for (let index = 4; index < row.length; index++) {
          updateRow[index] = 20; //TODO 進行表から算出
        }
      }
      return updateRow;
    });
    staffRange.setValues(updateStaffValue);
  }

  // 全作品工数一覧
  function writeTitlehours(monthData) {
    const startColumn = 4; // D列

    // 値を取得
    const values = monthlyManhoursSheet.getDataRange().getValues();
    //開始行の取得
    const startRowIndex =
      values.findIndex((row) => row[0] == "全作品工数一覧") + 3;
    //終了行の取得
    const endRowIndex = values.findIndex(
      (row) => row[3] == "トータル必要作業日"
    );
    //TODO 進行表から算出
    const title = [["BTR"], ["FRR"], ["SPY"], ["その他"]];
    const updateValue = title.map((row) => {
      row.push(...Array(monthData.length).fill(20));
      return row;
    });

    // 進行表から算出した行数と開始行・終了行から算出した行数の差分
    const diff = endRowIndex - startRowIndex - updateValue.length;
    // 差分がマイナスなら行追加、プラスなら行削除
    if (diff < 0) {
      monthlyManhoursSheet.insertRows(endRowIndex, diff * -1);
    } else if (diff > 0) {
      monthlyManhoursSheet.deleteRows(startRowIndex + 1, diff);
    }

    // 貼り付け範囲の取得
    const pasteRange = monthlyManhoursSheet.getRange(
      startRowIndex + 1,
      startColumn,
      updateValue.length,
      monthData.length + 1
    );
    // 値更新
    pasteRange.setValues(updateValue);
  }

  // 話数工数一覧
  function writeScenehours(monthData) {
    const startColumn = 2; // B列

    // 値を取得
    const values = monthlyManhoursSheet.getDataRange().getValues();
    //開始行の取得
    const startRowIndex =
      values.findIndex((row) => row[0] == "話数工数一覧") + 3;
    //終了行の取得
    const endRowIndex = values.findIndex(
      (row) => row[3] == "月間必要作業日合計"
    );
    //TODO 進行表から算出
    const title = [
      ["BTR", "#4", "2023/12/25"],
      ["FRR", "#A", ""],
      ["SPY", "#5", "2023/12/25"],
      ["その他", "BTR", ""],
    ];
    const updateValue = title.map((row) => {
      row.push(...Array(monthData.length).fill(20));
      return row;
    });

    // 進行表から算出した行数と開始行・終了行から算出した行数の差分
    const diff = endRowIndex - startRowIndex - updateValue.length;
    // 差分がマイナスなら行追加、プラスなら行削除
    if (diff < 0) {
      monthlyManhoursSheet.insertRows(endRowIndex, diff * -1);
    } else if (diff > 0) {
      monthlyManhoursSheet.deleteRows(startRowIndex + 1, diff);
    }

    // 貼り付け範囲の取得
    const pasteRange = monthlyManhoursSheet.getRange(
      startRowIndex + 1,
      startColumn,
      updateValue.length,
      monthData.length + 3
    );
    // 値更新
    pasteRange.setValues(updateValue);
  }
}
