// eslint-disable-next-line @typescript-eslint/no-unused-vars
function onFormSubmit(e: GoogleAppsScript.Events.SheetsOnFormSubmit): void {
  // 送信された内容を取得
  const timestamp =
    e.namedValues['出席日（当日の場合は空欄）'].toString() == ''
      ? e.namedValues['タイムスタンプ'].toString()
      : e.namedValues['出席日（当日の場合は空欄）'].toString();
  const partName = e.namedValues['パート名'].toString();
  const name = e.namedValues['氏名'].toString();
  console.log('タイムスタンプ：' + timestamp + ', パート名：' + partName + ', 出席者名：' + name);

  // フォーム送信された日から出席日を作成
  const date = new Date(Date.parse(timestamp));
  const sheetName = `${date.getFullYear()}-${('0' + (date.getMonth() + 1)).slice(-2)}-${(
    '0' + date.getDate()
  ).slice(-2)}`;
  console.log(`出席日：${sheetName}`);

  // 書き込む対象のスプレッドシートのオブジェクト作成
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  let newSheet = spreadsheet.getSheetByName(sheetName);

  // 出席日名のシートを作成
  // すでに同名のシートがあればスキップ、なければ作成
  if (newSheet == null) {
    spreadsheet.insertSheet(sheetName, 0);
    newSheet = spreadsheet.getSheetByName(sheetName);
    newSheet.getRange(1, 1).setValue('パート名');
    newSheet.getRange(1, 2).setValue('氏名');
  }

  const nameDataIndex = newSheet
    .getRange(1, 2, newSheet.getDataRange().getLastRow(), 1)
    .getValues()
    .flat()
    .indexOf(name);
  if (
    nameDataIndex != -1 &&
    newSheet
      .getRange(nameDataIndex + 1, 1)
      .getValues()
      .flat()
      .indexOf(partName) != -1
  ) {
    console.log('同日に同パート・同名の出席登録があるため、スキップ');
  } else {
    console.log('新規に出席登録');
    const AValues = newSheet.getRange('A:A').getValues();
    const LastRow = AValues.filter(String).length + 1;

    newSheet.getRange(LastRow, 1).setValue(partName);
    newSheet.getRange(LastRow, 2).setValue(name);
  }
}
