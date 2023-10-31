/*データスペース(作業者ベース部分)へ手入力にて入力した値を入れる。
カスタムメニュー固定、空白、進行表からの反映の際に使用され登録する。

セルの入力値(日付に紐づく)
固定(日付に紐づく)
空白(日付に紐づく)

データスペース（作品話数ベースデータ）へ反映を行う。
進行表から反映した際に登録する。また、手作業でのスケジュール表反映の為に反映を行う。
作品名
  作品色
  シーン名
    担当者
    開始日
    未入り(日)
    未撒き(日)
    撒済(日)＋再調整(日)
    未入り(色)
    未撒き(色)
    撒済(色)
*/
//TODO:フォーマットがおかしな場合のエラー対策を必要。
function demoUpdateDataSpace(){
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  let sheet = spreadsheet.getActiveSheet();
  let databaseSheet = spreadsheet.getSheetByName("データベース");

  updateDataBaseSheet(sheet,databaseSheet)
}

function updateDataBaseSheetMain(){
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // ソースシート（Aシート）を取得
  let sourceSheet = spreadsheet.getSheetByName("スケジュール管理仕様");
  // データベースシート（Bシート）を取得
  let databaseSheet = spreadsheet.getSheetByName("データベース");
  updateDataBaseSheet(sourceSheet,databaseSheet)
}


function updateDataSpaceMain(sheet){
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  let databaseSheet = spreadsheet.getSheetByName("データベース");
  
  updateDataBaseSheet(sheet,databaseSheet)

}




{
  //-----------------
  //定数定義（スケジュール表）
  //-----------------
  // 土日祝日の色
  const COLOR_HOLIDAY = "#808080";
  //担当者
  const PERSON_COLUMN_INDEX = 6;
  const PERSON_ROW_START_INDEX = 7;

  const BD_DATA_SPACE_START_KEYWORD = "以下データ用スペース(最大1100行)"
  const BD_DATA_SPACE_ROWS = 180  //データスペースの行数
  const BD_WORKER_BASE_DATA_KEYWORD = "作業者ベースデータ";
  const BD_OTHER_COMPANY_BASE_DATA_KEYWORD = "他社ベースデータ";
  const CLW_ARTWORK_PERSONS_TITLE = "CLW美術作業者一覧";
  const OTHER_COMPANY_HELP_TITLE = "他社ヘルプ";
  const FREE_SPACE_TITLE = "以下フリースぺース";

  const DATA_BASE_SHEET_NAME = "データベース"  //データベースシートのシート名
  const DATA_SOURCE_SHEET_NAME = "スケジュール管理仕様"


  //-----------------
  //変数
  //-----------------
  // スプレッドシート上でデータが入力されている最大範囲を選択
  let scheduleSheetAllRange = undefined
  //スケジュール表の全値取得
  let scheduleSheetDataValues = undefined
  //スケジュール表の背景色取得
  let scheduleSheetAllBackGrounds = undefined
  // データベースシートの範囲
  let dataBaseSheetAllRange = undefined
  // データベースシートの全値
  let dataBaseSheetValues = undefined

  //-----------------
  //関数
  //-----------------
  //メインの関数
  //sheetだけ引数。
  //getValuesの値やgetbackgroundsの値を引数として渡すと、この関数内で配列のサイズが変わるが、
  //呼び出し元にはそれが反映されずエラーになるので、sheetだけ渡している。
  //データベースシートにデータを書き込む
  function updateDataBaseSheet(sourceSheet,dataBaseSheet){
    const label = 'updateDataBaseSheet'
    console.time(label)

    if(dataBaseSheet === null){
      return
    }

    getScheduleSheetInfo(sourceSheet,dataBaseSheet)
    

    //データスペース初期化(以下データ用スペース(最大1000行)より下全削除)
    initializeDataSpace(scheduleSheetDataValues,dataBaseSheetValues)
    
    //スケジュール表の値を読み取り、作業者用ベースデータのデータスペースを更新
    let personsArray = updateDataSpaceParson()
    if(personsArray === null ){
      return;
    }
    
    //他社ベースデータのデータスペースを更新
    let otherCompanieArray = updateDataSpaceOtherCompanie()
    if(otherCompanieArray === null ){
      return;
    }

    //let dataBaseValues = extractDataBelowKeyword()
    // BD_DATA_SPACE_ROWS分のみに配列を加工
    dataBaseSheetValues.splice(BD_DATA_SPACE_ROWS);


    // 更新したスケジュール表情報でデータベースシートを更新
    updateScheduleSheetWithDataValues(dataBaseSheet,dataBaseSheetValues);

    console.timeEnd(label)
  }

  // スケジュール表情報取得
  function getScheduleSheetInfo(scheduleSheet,dataBaseSheet){
    // スプレッドシート上でデータが入力されている最大範囲を選択
    scheduleSheetAllRange = scheduleSheet.getDataRange();
    // 値を取得
    scheduleSheetDataValues = scheduleSheetAllRange.getValues();
    // 背景色を取得する
    scheduleSheetAllBackGrounds = scheduleSheetAllRange.getBackgrounds();
    // データベースの範囲を取得
    dataBaseSheetAllRange = dataBaseSheet.getDataRange();
    // 値を取得
    dataBaseSheetValues = dataBaseSheetAllRange.getValues()
  }

  // データスペース初期化(以下データ用スペース(最大1100行)より下を空行で塗りつぶす)
  // scduleSheetValuesの列数でdataBaseValuesを作成
  function initializeDataSpace(sourceValue,dataValues){
    const startRowIndex = dataValues.findIndex(row => row[0] === BD_DATA_SPACE_START_KEYWORD) + 1;
    //console.log(startRowIndex)
    const maxRows = BD_DATA_SPACE_ROWS + startRowIndex;

    // 既存の行数を確認
    const existingRows = dataValues.length - startRowIndex;

    // sourceValueの列数に基づいて新しい行を作成
    let newFirstRow = new Array(sourceValue[0].length).fill("");

    // dataValuesの1行目が存在する場合は、その行を更新
    if (dataValues[0]) {
      // 既存の1行目のデータを保持
      newFirstRow.splice(0, dataValues[0].length, ...dataValues[0]);
      // 1行目を更新
      dataValues[0] = newFirstRow;
    } else {
      // 1行目が存在しない場合は、新しい行を追加
      dataValues.unshift(newFirstRow);
    }

    // 既存の行がBD_DATA_SPACE_ROWS行未満の場合、追加
    if (existingRows < maxRows) {
      const rowsToAdd = maxRows - existingRows;
      //let newRowBackground = new Array(dataValues[0].length).fill("");

      for (let i = 0; i < rowsToAdd; i++) {
        let newRow = new Array(sourceValue[0].length).fill("");
        dataValues.push(newRow);
        //scheduleSheetAllBackGrounds.push(newRowBackground);
      }
    }
    // 既存の行がBD_DATA_SPACE_ROWS行以上の場合、上書き
    //let newRowBackground = new Array(scheduleSheetAllBackGrounds[0].length).fill("");
    //初期化の為空白にて塗りつぶし
    for (let i = 0; i < maxRows; i++) {
      let newRow = new Array(sourceValue[0].length).fill("");
      dataValues[startRowIndex + i] = newRow;
      //scheduleSheetAllBackGrounds[startRowIndex + i] = newRowBackground;
    }

    // "作業者ベースデータ"と"他社ベースデータ"を追加または上書き
    let workerBaseRow = new Array(sourceValue[0].length).fill("");
    workerBaseRow[PERSON_COLUMN_INDEX] = BD_WORKER_BASE_DATA_KEYWORD;

    let otherCompanyBaseRow = new Array(sourceValue[0].length).fill("");
    otherCompanyBaseRow[PERSON_COLUMN_INDEX] = BD_OTHER_COMPANY_BASE_DATA_KEYWORD;
    
    // 既存の行を上書きまたは新しい行を追加
    dataValues[startRowIndex] = workerBaseRow;
    dataValues[startRowIndex + 1] = otherCompanyBaseRow;
    
  }
  

  //スケジュール表の値を読み取り、作業者用ベースデータのデータスペースを更新
  function updateDataSpaceParson(){
    //担当者単位でループ
    //スケジュール表部分読み込み
    let personsArray = createPersonsArray(CLW_ARTWORK_PERSONS_TITLE,OTHER_COMPANY_HELP_TITLE)
    if(personsArray !== null){
      //データスペース用にpersonsArrayを編集
      editDataForDataSpace(personsArray, BD_DATA_SPACE_START_KEYWORD, BD_OTHER_COMPANY_BASE_DATA_KEYWORD, BD_WORKER_BASE_DATA_KEYWORD)
    }

    return personsArray

  }

  function updateDataSpaceOtherCompanie(){
    let otherCompanieArray = createPersonsArray(OTHER_COMPANY_HELP_TITLE, FREE_SPACE_TITLE)
    if(otherCompanieArray !== null){
      //データスペース用にpersonsArrayを編集
      editDataForDataSpace(otherCompanieArray, BD_DATA_SPACE_START_KEYWORD, BD_OTHER_COMPANY_BASE_DATA_KEYWORD, BD_OTHER_COMPANY_BASE_DATA_KEYWORD)
    }
    return otherCompanieArray
  }

  //A列にあるCLW美術作業者一覧行~"他社ヘルプ"行まで取得
  //その下にある作業者名を取得。作業者名行に記載がある背景色と、値を取り出す。
  //一個下のステータス行の背景色、値を取得
  function createPersonsArray(startKeyword, endKeyword) {
    class Person {
      constructor() {
        //担当者名
        this.personName = undefined;
        //スケジュール
        this.schedule = undefined;
        //スケジュール背景色
        this.scheduleColor = undefined;
        //スケジュールステータス
        this.states = undefined;
        //スケジュールステータス背景色
        this.statesColor = undefined;
      }
    }
  
    // "CLW美術作業者一覧" と "他社ヘルプ" の行番号を見つける
    let startRowIndex = scheduleSheetDataValues.findIndex(row => row[0] === startKeyword);
    let endRowIndex = scheduleSheetDataValues.findIndex(row => row[0] === endKeyword);
  
    // Personオブジェクトの配列を準備
    let persons = [];
    let person;  // 現在のPersonオブジェクトを保持するための変数
  
    // 指定された範囲の行を処理する
    if (startRowIndex !== -1 && endRowIndex !== -1) {
      for (let i = startRowIndex + 1; i < endRowIndex; i += 3) {
        // 新しいPersonオブジェクトを作成
        person = new Person();
        person.personName = scheduleSheetDataValues[i][PERSON_COLUMN_INDEX];  // personNameプロパティを設定
        
        // 右側の全てのセルの値を配列に格納し、scheduleプロパティを設定
        person.schedule = scheduleSheetDataValues[i].slice(PERSON_COLUMN_INDEX + 1);
        person.scheduleColor = scheduleSheetAllBackGrounds[i].slice(PERSON_COLUMN_INDEX + 1);
        
        // 翌行からstatesとstatesColorを取得
        person.states = scheduleSheetDataValues[i+1].slice(PERSON_COLUMN_INDEX + 1);
        person.statesColor = scheduleSheetAllBackGrounds[i+1].slice(PERSON_COLUMN_INDEX + 1);
        
        // Personオブジェクトをpersons配列に追加
        persons.push(person);
      }
    } else {
      // ダイアログを表示してエラーを通知
      let ui = SpreadsheetApp.getUi();
      ui.alert(
        'データ範囲エラー', 
        '必要なデータの範囲が見つかりませんでした。スプレッドシートのフォーマットが正しいことを確認し、再度お試しください。',
        ui.ButtonSet.OK
      );
      // 処理を停止
      return null;
    }
    // Personオブジェクトの配列を返す
    return persons;
  }

  // データスペース用に編集する関数
  function editDataForDataSpace(personsArray, startKeyword, endKeyword, targetKeyword) {
    // 行インデックスを取得する関数
    let {startIndex, endIndex, targetIndex} = findRowIndices({
      dataValues: dataBaseSheetValues,
      startKeyword: startKeyword,
      endKeyword: endKeyword,
      targetKeyword: targetKeyword,
      targetColumnIndex: PERSON_COLUMN_INDEX
    });
    let startPersonRowIndex = targetIndex;
    
    // エラーチェック: "作業者ベースデータ"行が見つかったかどうかを確認
    if (startPersonRowIndex === -1) {
      console.error(`Error: "${targetKeyword}" row not found.`);
      return;
    }

    for (let i = 0; i < personsArray.length; i++) {
      let person = personsArray[i];
      let targetRowIndexForPerson = startPersonRowIndex + i * 3 + 1;  //１作業者あたり、3行
      let targetRowIndexForStates = targetRowIndexForPerson + 1;
      let targetRowIndexForMemo = targetRowIndexForStates + 1;
      
      insertPersonNameAndStatesRows(targetRowIndexForPerson, targetRowIndexForStates, targetRowIndexForMemo, person); // personNameとstatesの行を設定する関数
      for (let j = 0; j < person.schedule.length; j++) {
        setBackgroundColorAndDataValues(targetRowIndexForPerson, targetRowIndexForStates, person, j); //背景色とデータ値を設定する関数
        
        if (personsArray[i].schedule[j] !== "") {  // スケジュールのシーンがある場合に背景色を取得
          let storyColor = personsArray[i].scheduleColor[j];
          let sceneName = personsArray[i].schedule[j];
          // 背景色が同じ限り、sceneNameの内容を次の配列personsArray[i].schedule[j＋1....n]にコピーする。
          for (let k = j + 1; k < personsArray[i].schedule.length; k++) {
            if (personsArray[i].schedule[k] !== "" && personsArray[i].schedule[k] !== sceneName) { // すでにテキストが存在し、かつsceneNameと異なる場合、sceneNameを更新し、ループの残りの部分をスキップする
              sceneName = personsArray[i].schedule[k];
              continue;
            }
            if (personsArray[i].scheduleColor[k] === storyColor) {
              personsArray[i].schedule[k] = sceneName;
            } else if(personsArray[i].scheduleColor[k] === COLOR_HOLIDAY){
              personsArray[i].schedule[k] = "";  //土日の背景色COLOR_HOLIDAYだった場合は、空欄にして、別の背景色が来るまでコピーし続ける。
            } else {
              break;  // 背景色が異なる場合、ループを終了する
            }
          }
        }
      }
    }
    /* editDataForDataSpace()のインナー関数 */
    // personNameとstatesを新規で行を挿入する関数
    function insertPersonNameAndStatesRows(targetRowIndexForPerson, targetRowIndexForStates, targetRowIndexForMemo, person) {
      let personNameRow  = new Array(scheduleSheetDataValues[0].length).fill("");  // 新しい行を作成し、全てのセルを空文字列で初期化する
      personNameRow [PERSON_COLUMN_INDEX] = person.personName;  // 作業者名を設定する
      dataBaseSheetValues.splice(targetRowIndexForPerson,0,personNameRow)
      
      let statesRow  = new Array(scheduleSheetDataValues[0].length).fill("");  // 新しい行を作成し、全てのセルを空文字列で初期化する
      statesRow [PERSON_COLUMN_INDEX] = "states";  // statesを設定する
      dataBaseSheetValues.splice(targetRowIndexForStates,0,statesRow)

      let memoRow = new Array(scheduleSheetDataValues[0].length).fill("");  // 新しい行を作成し、全てのセルを空文字列で初期化する
      memoRow [PERSON_COLUMN_INDEX] = "memo";  // memoを設定する
      dataBaseSheetValues.splice(targetRowIndexForMemo,0,memoRow)
      
      let personBaseBackgroundRow = new Array(scheduleSheetAllBackGrounds[0].length).fill("");  // 新しい背景色行を作成し、全てのセルを空文字列で初期化する
      let statesBaseBackgroundRow = new Array(scheduleSheetAllBackGrounds[0].length).fill("");  // 新しい背景色行を作成し、全てのセルを空文字列で初期化する
      let memoBaseBackgroundRow = new Array(scheduleSheetAllBackGrounds[0].length).fill("");  // 新しい背景色行を作成し、全てのセルを空文字列で初期化する
      
      scheduleSheetAllBackGrounds.splice(targetRowIndexForPerson, 0, personBaseBackgroundRow);
      scheduleSheetAllBackGrounds.splice(targetRowIndexForStates, 0, statesBaseBackgroundRow);
      scheduleSheetAllBackGrounds.splice(targetRowIndexForMemo, 0, memoBaseBackgroundRow);
    }

    // 背景色とデータ値を設定する関数
    function setBackgroundColorAndDataValues(targetRowIndexForPerson, targetRowIndexForStates, person, j) {
      let targetColumnIndex = PERSON_COLUMN_INDEX + 1 + j;
      dataBaseSheetValues[targetRowIndexForPerson][targetColumnIndex] = person.schedule[j];
      dataBaseSheetValues[targetRowIndexForStates][targetColumnIndex] = person.states[j];
    }
    /* editDataForDataSpace()のインナー関数 */
  }
  
  // A列の一定の範囲行からキーワードの行を取得する方法
  function findRowIndices({dataValues, startKeyword, endKeyword, targetKeyword, targetColumnIndex}) {
    let startIndex = dataValues.findIndex(row => row[0] === startKeyword);
    let endIndex = dataValues.findIndex(row => row[0] === endKeyword);
    let targetIndex = dataValues.findIndex(
      row => row[targetColumnIndex] === targetKeyword,
      startIndex + 1
    );
    return {startIndex, endIndex, targetIndex};
  }
    
  //スケジュール表の画面を更新する
  //TODO：getRangeでとった値は、値が入っていないととってこないので、この方式の方がよさそうだと思う。
  function updateScheduleSheetWithDataValues(sheet,values) {
    let startRow = 1;  // 1行目から開始
    let startColumn = 1;  // 1列目から開始
    let numRows = values.length;  // データの行数
    let numColumns = values[0].length;  // データの列数
    let range = sheet.getRange(startRow, startColumn, numRows, numColumns); 
    range.setValues(values);
      //スケジュール表の画面を更新する
    //scheduleSheetAllRange.setValues(scheduleSheetDataValues);
  }

  //データスペースのみに配列を加工
  function extractDataBelowKeyword() {
    const keywordIndex = scheduleSheetDataValues.findIndex(row => row[0] === BD_DATA_SPACE_START_KEYWORD);
    
    // キーワードが見つかったかどうかをチェック
    if (keywordIndex !== -1) {
      // キーワードの行から配列の末尾までスライスを取得
      const dataBelowKeyword = scheduleSheetDataValues.slice(keywordIndex);
      
      // 新しい配列 dataBelowKeyword が BD_DATA_SPACE_START_KEYWORD 以下の行を含む
      return dataBelowKeyword;
    } else {
      // キーワードが見つからなかった場合は、エラーをログに記録またはエラーをスロー
      console.error('Keyword not found');
      return null;
    }
  }
  
}