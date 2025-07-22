// 불량유형추가 3트

function moveAndFormatData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("오수진");
  var targetSheet = ss.getSheetByName("form");

  var lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) return;

  var result = [];
  for (var i = 0; i < lastRow - 1; i++) {
    // A열: 수식 =TEXT(오수진!A2, "yyyymmdd")
    var formula = "=TEXT(오수진!A" + (i + 2) + ', "yyyymmdd")';
    // F열: "품목코드 / 포장지명"에서 분리 (split 방식)
    var fValue = sourceSheet.getRange(i + 2, 6).getValue();
    var code = "";
    var name = "";
    if (typeof fValue === "string") {
      var parts = fValue.split("/");
      if (parts.length >= 2) {
        code = parts[0].trim();
        name = parts[1].trim();
      } else {
        code = fValue.trim();
        name = "";
      }
    } else if (typeof fValue === "number") {
      code = ("000000" + fValue).slice(-6);
    }
    // F열: 품목코드, G열: 품목명(포장지명)
    var newRow = [formula, "", "", "", "", code, name];
    result.push(newRow);
  }

  if (result.length > 0) {
    // 7열까지 포함해서 넣기
    targetSheet.getRange(2, 1, result.length, 7).setValues(result);
  }
}

// 하나 시트의 불량유형별로 둘 시트에 행 추가
function addDefectRowsFromHana() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hanaSheet = ss.getSheetByName("오수진"); // 시트명 수정
  var dulSheet = ss.getSheetByName("form");

  var hanaLastRow = hanaSheet.getLastRow();
  if (hanaLastRow < 2) return;

  var hanaData = hanaSheet.getRange(2, 1, hanaLastRow - 1, 11).getValues(); // A~K열
  var defectRows = [];

  for (var i = 0; i < hanaData.length; i++) {
    var row = hanaData[i];
    var date = row[0]; // A열
    var name = row[5]; // F열(포장지명)
    var sealing = row[8]; // I열(실링불량)
    var weight = row[9]; // J열(중량불량)
    var print = row[10]; // K열(날인불량)

    Logger.log(
      "row " +
        (i + 2) +
        ": date=" +
        date +
        ", name=" +
        name +
        ", sealing=" +
        sealing +
        ", weight=" +
        weight +
        ", print=" +
        print
    );

    // 실링불량 (I열: 9번째, K열: 11번째)
    if (sealing && sealing != 0) {
      defectRows.push([
        date,
        "",
        "",
        "",
        "",
        name,
        "",
        "",
        sealing,
        "",
        "00002",
      ]);
    }
    // 중량불량
    if (weight && weight != 0) {
      defectRows.push([
        date,
        "",
        "",
        "",
        "",
        name,
        "",
        "",
        weight,
        "",
        "00004",
      ]);
    }
    // 날인불량
    if (print && print != 0) {
      defectRows.push([date, "", "", "", "", name, "", "", print, "", "00003"]);
    }
  }

  if (defectRows.length > 0) {
    // form(둘) 시트의 마지막 행 다음에 추가
    var dulStartRow = dulSheet.getLastRow() + 1;
    dulSheet
      .getRange(dulStartRow, 1, defectRows.length, 11)
      .setValues(defectRows);
  }
}
