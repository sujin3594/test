// Google Apps Script - 원본시트(오수진)에서 이동할시트(form)로 조건별 데이터 이동
// result 배열의 각 행을 13개로 맞추고, 데이터가 오른쪽으로 밀리지 않도록 정확히 매핑

function moveRowsByCondition() {
  const SOURCE_SHEET = "오수진"; // 원본시트
  const TARGET_SHEET = "form"; // 이동할시트
  const START_ROW = 2; // 데이터 시작 행
  const SOURCE_COLS = 16; // A~P열(16열)
  const TARGET_COLS = 13; // form 시트의 열 개수 (A~M)

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = ss.getSheetByName(SOURCE_SHEET);
    var targetSheet = ss.getSheetByName(TARGET_SHEET);
    if (!sourceSheet || !targetSheet)
      throw new Error("시트를 찾을 수 없습니다.");

    var lastRow = sourceSheet.getLastRow();
    if (lastRow < START_ROW) return;

    // 원본시트의 2행~마지막행, A~P열 데이터 읽기
    var sourceData = sourceSheet
      .getRange(START_ROW, 1, lastRow - 1, SOURCE_COLS)
      .getValues();
    var result = [];

    // 조건 1: I열(9) 값이 있는 행만
    for (var i = 0; i < sourceData.length; i++) {
      var row = sourceData[i];
      var iVal = row[8]; // I열
      if (iVal !== "" && iVal != null && iVal != 0) {
        result.push([
          `=TEXT(${SOURCE_SHEET}!A${i + START_ROW},"yyyymmdd")`, // A열(수식)
          "",
          "",
          "",
          "", // B~E열 비움
          extractCode(row[5]), // F열: F열 텍스트 중 / 왼쪽 6자리 숫자
          `=VLOOKUP(F${result.length + 2},'백데이터'!A:I,2,0)`, // G열: 수식
          "", // H열 비움
          iVal, // I열: 값 그대로
          "", // J열 비움
          String("00002").padStart(5, "0"), // K열: 조건1 코드 (항상 5자리 문자열)
          "", // L열 비움
          `=TEXT(${SOURCE_SHEET}!A${i + START_ROW},"yy.mm.dd")`, // M열: P열 대신 A열 기준 수식
        ]);
      }
    }

    // 조건 2: J열(10) 값이 있는 행만 (조건1 아래부터)
    for (var i = 0; i < sourceData.length; i++) {
      var row = sourceData[i];
      var jVal = row[9]; // J열
      if (jVal !== "" && jVal != null && jVal != 0) {
        result.push([
          `=TEXT(${SOURCE_SHEET}!A${i + START_ROW},"yyyymmdd")`,
          "",
          "",
          "",
          "",
          extractCode(row[5]),
          `=VLOOKUP(F${result.length + 2},'백데이터'!A:I,2,0)`,
          "",
          jVal, // I열에 J열 값
          "",
          String("00004").padStart(5, "0"), // K열: 조건2 코드 (항상 5자리 문자열)
          "",
          `=TEXT(${SOURCE_SHEET}!A${i + START_ROW},"yy.mm.dd")`,
        ]);
      }
    }

    // 조건 3: K열(11) 값이 있는 행만 (조건2 아래부터)
    for (var i = 0; i < sourceData.length; i++) {
      var row = sourceData[i];
      var kVal = row[10]; // K열
      if (kVal !== "" && kVal != null && kVal != 0) {
        result.push([
          `=TEXT(${SOURCE_SHEET}!A${i + START_ROW},"yyyymmdd")`,
          "",
          "",
          "",
          "",
          extractCode(row[5]),
          `=VLOOKUP(F${result.length + 2},'백데이터'!A:I,2,0)`,
          "",
          kVal, // I열에 K열 값
          "",
          String("00003").padStart(5, "0"), // K열: 조건3 코드 (항상 5자리 문자열)
          "",
          `=TEXT(${SOURCE_SHEET}!A${i + START_ROW},"yy.mm.dd")`,
        ]);
      }
    }

    if (result.length > 0) {
      // form 시트의 2행부터 결과를 한 번에 복사 (13열)
      targetSheet
        .getRange(START_ROW, 1, result.length, TARGET_COLS)
        .setValues(result);

      // K열(11번째, 즉 'K')의 서식을 텍스트로 지정 (2행~마지막행)
      targetSheet
        .getRange(START_ROW, 11, result.length, 1)
        .setNumberFormat("@STRING@");

      // 시트 전체(2행~마지막행)를 A열 기준 오름차순 정렬
      var totalRows = targetSheet.getLastRow();
      if (totalRows >= START_ROW) {
        targetSheet
          .getRange(START_ROW, 1, totalRows - START_ROW + 1, TARGET_COLS)
          .sort({ column: 1, ascending: true });
      }
    }
  } catch (e) {
    Logger.log("moveRowsByCondition 오류: " + e.message);
  }
}

/**
 * F열 텍스트에서 / 왼쪽 6자리 숫자만 추출하는 함수
 * @param {string} fText
 * @returns {string}
 */
function extractCode(fText) {
  if (typeof fText === "string") {
    var parts = fText.split("/");
    var code = parts[0].replace(/[^0-9]/g, "").slice(0, 6);
    return code;
  } else if (typeof fText === "number") {
    return ("000000" + fText).slice(-6);
  }
  return "";
}

/**
 * 스프레드시트 메뉴에 '이카운트 >> 불량 엑셀 변환' 항목을 추가하는 함수
 * 메뉴를 클릭하면 moveRowsByCondition 함수가 실행됨
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("이카운트")
    .addItem("불량 엑셀 변환", "moveRowsByCondition")
    .addToUi();
}
