/*
時間測定用ファイルの為、最終的には削除予定。
*/
function timeTestMain(){
  timeTest()
}
{
  function timeTest(){
    //テスト開始
    var sheetName = "性能テスト";
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    var range = getRangeData(sheet)
    var backgroundcolor = getRandomColor()
    var bordercolor1 = getRandomColor()
    var bordercolor2 = getRandomColor()
    var borderStyle1 = setRandomBorderStyle()
    var borderStyle2 = setRandomBorderStyle()
    var randomString1 = generateRandomSymbols()
    var randomString2 = generateRandomString()
    //setのテスト
    const labelset = 'setTimeTest'
    console.time(labelset)
    setRangeValues(range, randomString1)
    setAllBackgrounds(range, backgroundcolor)
    setBottomBorder1(range, bordercolor1, borderStyle1)
    setBottomBorders(sheet, range, bordercolor2, borderStyle2)
    setRangeNotes(range,randomString2)
    console.timeEnd(labelset)
    //getのテスト
    const label = 'getTimeTest'
    console.time(label)
    getValuesTime(range)
    getBackgroundsTime(range)
    getNoteTime(range)
    /*
    getBorderBottom(range)
    //getBorderBottomStyle(range)
    //テスト終了
    console.timeEnd(label)
    */
  }
  function getRangeData(sheet) {
    const label = 'getRangeData'
    console.time(label)
    var range = sheet.getRange(2, 2, 60, 730);  //2,2,30,730
    console.timeEnd(label)
    return range
  }

  //setのテスト
  function setRangeValues(range, inputString) {
    var numRows = range.getNumRows();
    var numCols = range.getNumColumns();
  
    var values = [];
    for (var i = 0; i < numRows; i++) {
      values[i] = [];
      for (var j = 0; j < numCols; j++) {
        // 行番号が奇数である場合にのみ値を設定
        if ((i + 1) % 2 == 1) {
          values[i][j] = inputString;
        } else {
          values[i][j] = "";  // それ以外の場合は空の文字列を設定
        }
      }
    }
  
    const label = 'setRangeValues';
    console.time(label);
    range.setValues(values);
    console.timeEnd(label);
  }
  
  function setAllBackgrounds(range, color) {
  
    var numRows = range.getNumRows();
    var numCols = range.getNumColumns();
  
    var colors = [];
    for (var i = 0; i < numRows; i++) {
      colors[i] = [];
      for (var j = 0; j < numCols; j++) {
        // 行番号が奇数である場合にのみ色を設定
        if ((i + 1) % 2 == 1) {
          colors[i][j] = color;
        } else {
          colors[i][j] = null;  // それ以外の場合は背景色を変更しない
        }
      }
    }
    const label = 'setAllBackgrounds';
    console.time(label);
    range.setBackgrounds(colors);
    console.timeEnd(label);
  }
  
  function setBottomBorder1(range,color,borderStyle) {
    const label = 'setBottomBorder1'
    console.time(label)
    range.setBorder(
      false,  // top
      false,  // left
      true,   // bottom
      false,  // right
      false,  // vertical
      false,  // horizontal
      color,  // color
      borderStyle // style
    );
    console.timeEnd(label)
  }

  function setBottomBorders(sheet, range, color, borderStyle) {
    const label = 'setBottomBorders';
    console.time(label);
    
    var startRow = range.getRow();
    var numCols = range.getNumColumns();
    var endRow = startRow + range.getNumRows();
    
    for (var i = startRow; i < endRow; i++) {
      // 行番号が奇数である場合にのみ境界線を設定(start行がずれている為偶数になる)
      if (i % 2 == 0) {
        var rowRange = sheet.getRange(i, range.getColumn(), 1, numCols);
        rowRange.setBorder(
          false,  // top
          false,  // left
          true,   // bottom
          false,  // right
          false,  // vertical
          false,  // horizontal
          color,  // color
          borderStyle  // style
        );
      }
    }
    
    console.timeEnd(label);
  }

  function setRangeNotes(range, inputNote) {
    var numRows = range.getNumRows();
    var numCols = range.getNumColumns();
  
    var notes = [];
    for (var i = 0; i < numRows; i++) {
      notes[i] = [];
      for (var j = 0; j < numCols; j++) {
        // 行番号が奇数である場合にのみノートを設定
        if ((i + 1) % 2 == 1) {
          notes[i][j] = inputNote;
        } else {
          notes[i][j] = "";  // それ以外の場合は空のノートを設定
        }
      }
    }
  
    const label = 'setRangeNotes';
    console.time(label);
    range.setNotes(notes);
    console.timeEnd(label);
  }
  
  
  
  //getのテスト

  function getValuesTime(range){
    const label = 'getValuesTime'
    console.time(label)
    var values = range.getValues()
    console.log(values)
    console.timeEnd(label)
  }
  function getBackgroundsTime(range){
    const label = 'getBackgroundsTime'
    console.time(label)
    var BackgroundColors = range.getBackgrounds()
    //console.log(BackgroundColors)
    console.timeEnd(label)
  }
  function getBorderBottom(range){
    const label = 'getBorderBottom'
    console.time(label)
    var numRows = range.getNumRows();
    var numCols = range.getNumColumns();
    
    var borderBottomColors = [];
    for (var i = 1; i <= numRows; i++) {
      borderBottomColors[i] = [];
      for (var j = 1; j <= numCols; j++) {
        var cell = range.getCell(i, j);
        var border = cell.getBorder();
        var borderColor = border ? border.getBottom().getColor().asRgbColor().asHexString() : null;
        //console(borderColor)
        borderBottomColors[i][j] = borderColor
      }
    }
    console.timeEnd(label)
  }

  function getBorderBottomStyle(range) {
    const label = 'getBorderBottomStyle'
    console.time(label)

    var numRows = range.getNumRows();
    var numCols = range.getNumColumns();

    var borderBottomStyles = [];
    for (var i = 1; i <= numRows; i++) {
      borderBottomStyles[i] = [];

      for (var j = 1; j <= numCols; j++) {
        var cell = range.getCell(i, j);
        var border = cell.getBorder();
        var style = border ? border.getBottom().getBorderStyle() : null;
        borderBottomStyles[i][j] = style
      }
    }
    console.timeEnd(label)
  }
  function getRandomColor() {
    var letters = '0123456789ABCDEF';
    var color = '#';
    for (var i = 0; i < 6; i++) {
      color += letters[Math.floor(Math.random() * 16)];
    }
    return color;
  }

  function getNoteTime(range){
    const label = 'getNoteTime'
    console.time(label)
    var results = range.getNotes();
    console.timeEnd(label)
    for (var i in results) {
      for (var j in results[i]) {
        //Logger.log(results[i][j]);
      }
    }
  }

  function setRandomBorderStyle() {
    var borderStyles = [
      SpreadsheetApp.BorderStyle.SOLID,
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM,
      SpreadsheetApp.BorderStyle.SOLID_THICK,
      SpreadsheetApp.BorderStyle.DOTTED,
      SpreadsheetApp.BorderStyle.DASHED,
      SpreadsheetApp.BorderStyle.DOUBLE
    ];
    return randomStyle = borderStyles[Math.floor(Math.random() * borderStyles.length)];
  }  
  function generateRandomString() {
    var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    var result = '';
    for (var i = 0; i < 2; i++) {
      var randomIndex = Math.floor(Math.random() * chars.length);
      result += chars.charAt(randomIndex);
    }
    return result;
  }
  function generateRandomSymbols() {
    var symbols = '▼▲●■□○★☆⇧⇨⇩⇦';
    var result = '';
    for (var i = 0; i < 3; i++) {
      var randomIndex = Math.floor(Math.random() * symbols.length);
      result += symbols.charAt(randomIndex);
    }
    return result;
  }
}

