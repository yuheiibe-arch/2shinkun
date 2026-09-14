/**
 * ====================================================================
 * メイン処理（集計実行・シート作成）
 * ====================================================================
 */

function runProcess(selectedSheetName) {
  const urlBookId = '1Ky5fXKvEWFodUwcu-HnHKiOBn6zdb090j79OjI6KNtk';
  const urlSheetName = 'URL';
  const sourceRangeA1 = 'A150:CQ195';
  const destinationBook = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const trimmedSheetName = selectedSheetName.trim();

  try {
    const urlBook = SpreadsheetApp.openById(urlBookId);
    const urlSheet = urlBook.getSheetByName(urlSheetName);
    const urlData = urlSheet.getDataRange().getValues();
    let sourceUrl = '';
    for (let i = 1; i < urlData.length; i++) {
      if (urlData[i][0].trim() === trimmedSheetName) {
        sourceUrl = urlData[i][1];
        break;
      }
    }
    if (!sourceUrl) throw new Error(`URLが見つかりません: ${trimmedSheetName}`);

    const match = trimmedSheetName.match(/(\d{4})[年\/](\d{1,2})/);
    let newSheetName = match ? `${match[1]}/${parseInt(match[2], 10)}` : trimmedSheetName;

    const destinationSheet = setupDestinationSheet(destinationBook, newSheetName);
    
    destinationBook.toast(`転記開始: ${trimmedSheetName} → シート: ${newSheetName}`);
    const sourceBook = SpreadsheetApp.openByUrl(sourceUrl);
    const sourceSheet = sourceBook.getSheetByName(trimmedSheetName);
    if (!sourceSheet) throw new Error(`シートが見つかりません: ${trimmedSheetName}`);
    
    const sourceData = sourceSheet.getRange(sourceRangeA1).getValues();
    processAndPasteData(destinationSheet, sourceData);

    destinationBook.toast(`「${newSheetName}」の作成が完了しました。`, '完了', 3);
  } catch (e) {
    ui.alert('エラー: ' + e.message);
  }
}