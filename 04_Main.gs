/**
 * ====================================================================
 * メイン処理（集計実行・シート作成）
 * ====================================================================
 */

function runProcess(selectedSheetName) {
  const urlBookId = '1Ky5fXKvEWFodUwcu-HnHKiOBn6zdb090j79OjI6KNtk';
  const urlSheetName = 'URL';
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

    const sourceBook = SpreadsheetApp.openByUrl(sourceUrl);
    const sourceSheet = sourceBook.getSheetByName(trimmedSheetName);
    if (!sourceSheet) throw new Error(`シートが見つかりません: ${trimmedSheetName}`);

    // ★動的検出：「2診体制」ブロックの開始行を特定する
    const startRow = find2ndConsultationStartRow_(sourceSheet);
    const lastRow = sourceSheet.getLastRow();
    if (lastRow < startRow + 3) {
      throw new Error(`「2診体制」のデータ行が見つかりません（検出開始行: ${startRow}行目）`);
    }

    const destinationSheet = setupDestinationSheet(destinationBook, newSheetName);
    destinationBook.toast(`転記開始: ${trimmedSheetName} (${startRow}行目〜) → シート: ${newSheetName}`);

    // 開始行から最終行までの A〜CQ列(95列分) を取得
    const sourceRangeA1 = `A${startRow}:CQ${lastRow}`;
    const sourceData = sourceSheet.getRange(sourceRangeA1).getValues();
    processAndPasteData(destinationSheet, sourceData);

    destinationBook.toast(`「${newSheetName}」の作成が完了しました。（元シート ${startRow}行目〜取得）`, '完了', 3);
  } catch (e) {
    ui.alert('エラー: ' + e.message);
  }
}

/**
 * 元シートから「2診体制」テーブルの開始行番号（1始まり）を動的に検索する
 */
function find2ndConsultationStartRow_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 4) {
    throw new Error('シートの行数が不足しています。');
  }

  // A〜E列のみ一括取得して高速スキャン
  const scanValues = sheet.getRange(1, 1, lastRow, 5).getValues();

  const normalizeTitle = (val) =>
    String(val || '')
      .normalize('NFKC')
      .replace(/\s+/g, '')
      .replace(/^▼/, '');

  const isTimeHeaderRow = (rIdx) => {
    if (rIdx >= scanValues.length) return false;
    return (
      String(scanValues[rIdx][2]).trim() === '午前' &&
      String(scanValues[rIdx][3]).trim() === '午後' &&
      String(scanValues[rIdx][4]).trim() === '夜間'
    );
  };

  // 第1優先：A列が「2診体制」と完全一致 ＆ 2行下が「午前 / 午後 / 夜間」
  for (let r = 0; r < scanValues.length - 2; r++) {
    if (normalizeTitle(scanValues[r][0]) === '2診体制' && isTimeHeaderRow(r + 2)) {
      return r + 1; // 1始まりの行番号
    }
  }

  // 第2優先（表記揺れ対応）：A列に「2診体制」を含む ＆ 2行下が「午前 / 午後 / 夜間」
  for (let r = 0; r < scanValues.length - 2; r++) {
    if (normalizeTitle(scanValues[r][0]).includes('2診体制') && isTimeHeaderRow(r + 2)) {
      return r + 1;
    }
  }

  throw new Error('シート内に「2診体制」ブロック（ヘッダー: 午前/午後/夜間）が見つかりませんでした。');
}