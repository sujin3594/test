function moveAndFormatData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName('오수진');
  var targetSheet = ss.getSheetByName('form');

  var lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) return;

  var result = [];
  for (var i = 0; i < lastRow - 1; i++) {
    // A열: 수식 =TEXT(오수진!A2, "yyyymmdd")
    var formula = '=TEXT(오수진!A' + (i + 2) + ', "yyyymmdd")';
    // F열: 6자리 숫자만 추출
    var code = '';
    var fValue = sourceSheet.getRange(i + 2, 6).getValue();
    if (typeof fValue === 'string') {
      var match = fValue.match(/^(\d{6})/);
      if (match) code = match[1];
    } else if (typeof fValue === 'number') {
      code = ('000000' + fValue).slice(-6);
    }
    var newRow = [formula, '', '', '', '', code];
    result.push(newRow);
  }

  if (result.length > 0) {
    targetSheet.getRange(2, 1, result.length, 6).setFormulas(result);
  }
}