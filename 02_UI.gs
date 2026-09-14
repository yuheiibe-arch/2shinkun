/**
 * ====================================================================
 * UI・メニュー関連
 * ====================================================================
 */

/**
 * メニュー作成
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('カスタム集計')
    .addItem('集計を実行', 'showSheetSelectionDialog')
    .addSeparator()
    .addItem('1. 作業リストを作成', 'createWorkList')
    .addItem('2. ★募集リストを作成', 'createRecruitmentListFromActiveSheet')
    .addToUi();
}

/**
 * 集計実行ダイアログ
 */
function showSheetSelectionDialog() {
  const urlBookId = '1Ky5fXKvEWFodUwcu-HnHKiOBn6zdb090j79OjI6KNtk';
  const urlSheetName = 'URL';
  const urlBook = SpreadsheetApp.openById(urlBookId);
  const urlSheet = urlBook.getSheetByName(urlSheetName);

  if (urlSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('URLシートにデータがありません。');
    return;
  }

  const sheetNames = urlSheet.getRange(2, 1, urlSheet.getLastRow() - 1, 1).getValues().flat().filter(name => name);
  if (sheetNames.length === 0) {
    SpreadsheetApp.getUi().alert('URLシートにシート名がありません。');
    return;
  }

  // デフォルト選択肢（リスト内の選択用）
  const nextMonth = new Date();
  nextMonth.setMonth(nextMonth.getMonth() + 1);
  const defaultSelection = `${nextMonth.getFullYear()}年${nextMonth.getMonth() + 1}月`;

  let html = HtmlService.createTemplateFromFile('Dialog');
  html.sheetNames = sheetNames;
  html.defaultSelection = defaultSelection;

  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setWidth(350).setHeight(150), '集計対象シートの選択');
}