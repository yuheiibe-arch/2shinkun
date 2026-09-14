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

  // ★修正: 今月より前の過去の月をリストから除外する
  const today = new Date();
  const currentYear = today.getFullYear();
  const currentMonth = today.getMonth() + 1;

  const sheetNames = urlSheet.getRange(2, 1, urlSheet.getLastRow() - 1, 1).getValues().flat().filter(name => {
    if (!name) return false;
    const match = String(name).match(/(\d{4})[年\/](\d{1,2})/);
    if (match) {
      const y = parseInt(match[1], 10);
      const m = parseInt(match[2], 10);
      if (y < currentYear) return false;
      if (y === currentYear && m < currentMonth) return false;
    }
    return true;
  });

  if (sheetNames.length === 0) {
    SpreadsheetApp.getUi().alert('処理可能な対象月（今月以降）がURLシートにありません。');
    return;
  }

  const nextMonth = new Date();
  nextMonth.setMonth(nextMonth.getMonth() + 1);
  const defaultSelection = `${nextMonth.getFullYear()}年${nextMonth.getMonth() + 1}月`;

  let html = HtmlService.createTemplateFromFile('Dialog');
  html.sheetNames = sheetNames;
  html.defaultSelection = defaultSelection;

  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setWidth(350).setHeight(150), '集計対象シートの選択');
}