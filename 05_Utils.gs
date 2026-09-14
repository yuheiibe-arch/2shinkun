/**
 * ====================================================================
 * 補助関数・ユーティリティ
 * ====================================================================
 */

/**
 * 祝日シートから日付を取得
 */
function getHolidaysFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const holidaySheet = ss.getSheetByName('祝日');
  if (!holidaySheet) {
    SpreadsheetApp.getUi().alert('注意', '「祝日」シートが見つかりませんでした。', SpreadsheetApp.getUi().ButtonSet.OK);
    return [];
  }
  const lastRow = holidaySheet.getLastRow();
  if (lastRow < 1) return [];
  
  const values = holidaySheet.getRange(1, 1, lastRow, 1).getValues();
  return values.flat()
    .filter(date => date)
    .map(date => {
      const d = new Date(date);
      return !isNaN(d) ? Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") : null;
    })
    .filter(d => d);
}

/**
 * シート初期化（作成または上書き）
 */
function setupDestinationSheet(book, name) {
  const existing = book.getSheetByName(name);
  if (existing) book.deleteSheet(existing);
  return book.insertSheet(name);
}

/**
 * データ転記と並び替え
 */
function processAndPasteData(sheet, data) {
  const outputData = [];
  const headerRows = 3;
  for (let i = 0; i < headerRows; i++) outputData.push(data[i]);
  
  for (let i = headerRows; i < data.length; i++) {
    const row = data[i];
    if (row[0] === '千葉NT') row[0] = '千葉ニュータウン中央';
    
    let hasValue = false;
    const newRow = [row[0], row[1]];
    for (let j = 2; j < row.length; j++) {
      const val = row[j];
      newRow.push(val ? val : '');
      if (parseInt(val, 10) >= 1) hasValue = true;
    }
    if (hasValue) outputData.push(newRow);
  }

  if (outputData.length > headerRows) {
    const sortOrder = [ "西葛西", "北葛西", "代官山", "東品川", "武蔵小山", "東雲", "亀有", "錦糸町", "光が丘", "板橋", "西新井", "国立", "北綾瀬", "武蔵小杉", "天王町", "海老名", "茅ヶ崎", "小田栄", "東戸塚", "相模原", "高田", "新百合ヶ丘", "柏の葉", "流山おおたかの森", "八千代緑が丘", "千葉ニュータウン中央", "新鎌ケ谷", "稲毛海岸", "村上", "志木", "越谷レイクタウン", "川口", "南浦和", "草加松原", "川越", "所沢", "与野", "東岸和田", "阿波座", "セブンパーク天美", "堺鉄砲町", "長吉長原", "鶴見緑地", "豊中" ];
    const header = outputData.slice(0, headerRows);
    const body = outputData.slice(headerRows);
    
    body.sort((a, b) => {
      let iA = sortOrder.indexOf(a[0]);
      let iB = sortOrder.indexOf(b[0]);
      if (iA === -1) iA = 999;
      if (iB === -1) iB = 999;
      return iA === iB ? a[1].localeCompare(b[1]) : iA - iB;
    });
    
    const finalData = header.concat(body);
    sheet.getRange(1, 1, finalData.length, finalData[0].length).setValues(finalData);
  }
}