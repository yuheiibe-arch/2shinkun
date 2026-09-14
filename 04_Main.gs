/**
 * ====================================================================
 * メイン処理（集計・リスト作成）
 * ====================================================================
 */

/**
 * 集計実行（シート作成）
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

    // 今日を基準に「翌月1日」を計算し、シート名(YYYY/MM)を決定
    const today = new Date();
    const nextMonthDate = new Date(today.getFullYear(), today.getMonth() + 1, 1);
    const newSheetName = Utilities.formatDate(nextMonthDate, Session.getScriptTimeZone(), "yyyy/MM");

    // シート作成・削除処理
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

/**
 * 1. 作業リストを作成 (基幹エンジン・年度判定対応版)
 */
function createWorkList() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const book = sheet.getParent();
  const HOLIDAYS_LIST = getHolidaysFromSheet();

  try {
    let targetYear, targetMonth;
    const sheetNameMatch = sheet.getName().match(/^(\d{4})\/(\d{1,2})$/);
    if (sheetNameMatch) {
      targetYear = parseInt(sheetNameMatch[1], 10);
      targetMonth = parseInt(sheetNameMatch[2], 10);
    } else {
      const nextMonth = new Date();
      nextMonth.setMonth(nextMonth.getMonth() + 1);
      targetYear = nextMonth.getFullYear();
      targetMonth = nextMonth.getMonth() + 1;
    }

    // ★★★ 修正: 月が1〜3月の場合は、前年を「年度(Fiscal Year)」として扱う ★★★
    const fiscalYear = (targetMonth <= 3) ? targetYear - 1 : targetYear;

    // 基幹エンジン初期化
    initializeWageData();

    const data = sheet.getDataRange().getValues();
    const dateHeaders = data[0];
    const timeHeaders = data[2];
    const output = [['拠点名', '該当日', '曜日', '曜日週数', '時間帯', '時給', '対応済', '', '日給', '勤務時間']];
    let totalReqH = 0, totalReqW = 0, totalExH = 0, totalExW = 0;

    for (let i = 3; i < data.length; i++) {
      let loc = data[i][0];
      if (loc === '千葉NT') loc = '千葉ニュータウン中央';
      let dept = data[i][1] || '小児科'; 

      for (let j = 2; j < data[i].length; j++) {
        const val = parseInt(data[i][j], 10);
        if (!val || val < 1) continue;

        // カレンダー上の年はそのまま適用（日付がズレないようにするため）
        const date = new Date(dateHeaders[j]);
        date.setFullYear(targetYear); 

        const dateStr = Utilities.formatDate(date, "JST", "MM/dd");
        const holidayStr = Utilities.formatDate(date, "JST", "yyyy-MM-dd");
        const dayOfWeek = HOLIDAYS_LIST.includes(holidayStr) ? '祝' : DAYS_OF_WEEK_JP[date.getDay()];
        const weekNum = `第${Math.floor((date.getDate() - 1) / 7) + 1}週`;

        let timeSlot = timeHeaders[j], hours = 0;
        if (timeHeaders[j] === '午前') {
          if (val === 2) { timeSlot = '10:00~12:00'; hours = 2; }
          if (val === 3) { timeSlot = '10:00~13:00'; hours = 3; }
          if (val === 4) { timeSlot = '09:00~13:00'; hours = 4; }
        } else if (timeHeaders[j] === '午後') {
          if (val === 2) { timeSlot = '15:00~17:00'; hours = 2; }
          if (val === 3) { timeSlot = '15:00~18:00'; hours = 3; }
        } else if (timeHeaders[j] === '夜間') {
          hours = 3;
        }
        
        if (hours === 0) continue;

        // 時給取得ロジック
        const dayType = (dayOfWeek === '祝' || dayOfWeek === '日' || dayOfWeek === '土') ? 'hol' : 'wd';
        let timeType = 'am';
        if (timeHeaders[j] === '午後') timeType = 'pm';
        if (timeHeaders[j] === '夜間') timeType = 'nt';

        const rateKey = `${dayType}_${timeType}`; 

        const clinicWages = getClinicWages(loc);
        let wage = 0;

        if (clinicWages && clinicWages.length > 0) {
          let targetDeptData = clinicWages.find(w => w.department.includes(dept)) || clinicWages[0];
          
          // ★★★ 修正: カレンダー年ではなく「年度(fiscalYear)」を基準に判定する ★★★
          const targetRates = fiscalYear >= 2026 ? targetDeptData.rates.y2026 : targetDeptData.rates.y2025;
          
          if (targetRates && targetRates[rateKey]) {
            wage = Number(targetRates[rateKey]) || 0;
          }
        }

        const dailyPay = wage * hours;

        totalReqH += hours;
        totalReqW += dailyPay;
        
        if (typeof data[i][j] === 'string' && /[a-zA-Z]/.test(data[i][j])) {
          totalExH += hours;
          totalExW += dailyPay;
        } else if (wage > 0) {
          output.push([loc, dateStr, dayOfWeek, weekNum, timeSlot, wage, '', '', dailyPay, hours]);
        }
      }
    }

    if (output.length > 1) {
      const startRow = sheet.getLastRow() + 4;
      const range = sheet.getRange(startRow, 1, output.length, 10);
      range.setValues(output);
      sheet.getRange(startRow + 1, 7, output.length - 1, 1).insertCheckboxes();
      
      const summaryStartRow = sheet.getLastRow() + 4;
      sheet.getRange('K' + summaryStartRow).setValue('追加拠点数').setFontWeight('bold');
      sheet.getRange('L' + summaryStartRow).setValue([...new Set(output.slice(1).map(r=>r[0]))].length);
      sheet.getRange('K' + (summaryStartRow + 1)).setValue('追加枠数').setFontWeight('bold');
      sheet.getRange('L' + (summaryStartRow + 1)).setValue(output.length - 1);
      sheet.getRange('K' + (summaryStartRow + 2)).setValue('総時間').setFontWeight('bold');
      sheet.getRange('L' + (summaryStartRow + 2)).setValue(totalReqH);
      sheet.getRange('K' + (summaryStartRow + 3)).setValue('総額').setFontWeight('bold');
      sheet.getRange('L' + (summaryStartRow + 3)).setValue(totalReqW).setNumberFormat('"¥"#,##0');

      // ★追加: 作業リスト作成と同時に２診要望一覧へ自動転記する
      append2ndConsultationRequests(sheet);

      book.toast('作業リストの作成が完了しました。', '完了', 3);
    } else {
      ui.alert('データがありません。あるいは時給が取得できませんでした。');
    }
  } catch (e) {
    ui.alert('エラー: ' + e.message);
  }
}

/**
 * 2. 募集リストを作成
 */
function createRecruitmentListFromActiveSheet() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const book = sheet.getParent();
  const HOLIDAYS_LIST = getHolidaysFromSheet();
  const START_ROW = 39;

  try {
    let targetYear, targetMonth;
    const sheetNameMatch = sheet.getName().match(/^(\d{4})\/(\d{1,2})$/);
    if (sheetNameMatch) {
      targetYear = parseInt(sheetNameMatch[1], 10);
      targetMonth = parseInt(sheetNameMatch[2], 10);
    } else {
      const nextMonth = new Date();
      nextMonth.setMonth(nextMonth.getMonth() + 1);
      targetYear = nextMonth.getFullYear();
      targetMonth = nextMonth.getMonth() + 1;
    }

    const data = sheet.getDataRange().getValues();
    let startIdx = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === '拠点名' && data[i][1] === '該当日') {
        startIdx = i;
        break;
      }
    }
    if (startIdx === -1) throw new Error('作業リストが見つかりません');

    const workData = data.slice(startIdx + 1);
    const groups = {};

    workData.forEach(row => {
      const [loc, dateVal, , , timeSlot, wage] = row;
      if (!loc || !dateVal || !wage) return;

      let date;
      if (dateVal instanceof Date) {
        date = new Date(dateVal);
        date.setFullYear(targetYear);
      } else {
        date = new Date(`${targetYear}/${dateVal}`);
      }
      
      if (isNaN(date.getTime())) return;

      const holidayStr = Utilities.formatDate(date, "JST", "yyyy-MM-dd");
      const isHoliday = HOLIDAYS_LIST.includes(holidayStr);
      const dayOfWeek = DAYS_OF_WEEK_JP[date.getDay()];
      const groupDay = isHoliday ? '祝' : dayOfWeek;
      
      const timeStr = String(timeSlot);
      const simpleTime = timeStr.replace('~', '-');
      const key = `${loc}-${timeStr}-${groupDay}`;

      if (!groups[key]) {
        let hours = 0, start = '', end = '';
        const m = timeStr.match(/(\d{2}:\d{2})~(\d{2}:\d{2})/);
        if (m) {
          start = m[1]; end = m[2];
          hours = (new Date(`2000/1/1 ${end}`) - new Date(`2000/1/1 ${start}`)) / 36e5;
        }
        groups[key] = {
          loc, title: `${loc}／${simpleTime}／${groupDay}`,
          start, end, hours, wage, dates: []
        };
      }
      
      const dateDisp = `${Utilities.formatDate(date, "JST", "MM/dd")}(${dayOfWeek})`;
      groups[key].dates.push({ date, display: dateDisp });
    });

    const output = [];
    const sortOrder = [ "西葛西", "北葛西", "代官山", "東品川", "武蔵小山", "東雲", "亀有", "錦糸町", "光が丘", "板橋", "西新井", "国立", "北綾瀬", "武蔵小杉", "天王町", "海老名", "茅ヶ崎", "小田栄", "東戸塚", "相模原", "高田", "新百合ヶ丘", "柏の葉", "流山おおたかの森", "八千代緑が丘", "千葉ニュータウン中央", "新鎌ケ谷", "稲毛海岸", "村上", "志木", "越谷レイクタウン", "川口", "南浦和", "草加松原", "川越", "所沢", "与野", "東岸和田", "阿波座", "セブンパーク天美", "堺鉄砲町", "長吉長原", "鶴見緑地", "豊中" ];
    
    const locKeys = Object.keys(groups).reduce((acc, k) => {
      const loc = groups[k].loc;
      if (!acc[loc]) acc[loc] = [];
      acc[loc].push(groups[k]);
      return acc;
    }, {});

    const sortedLocs = Object.keys(locKeys).sort((a, b) => {
      let iA = sortOrder.indexOf(a), iB = sortOrder.indexOf(b);
      if (iA === -1) iA = 999;
      if (iB === -1) iB = 999;
      return iA - iB;
    });

    const monthStart = Utilities.formatDate(new Date(targetYear, targetMonth - 1, 1), "JST", "yyyy/MM/dd");
    const monthEnd = Utilities.formatDate(new Date(targetYear, targetMonth, 0), "JST", "yyyy/MM/dd");

    let totalH = 0, totalCost = 0, totalCount = 0;

    sortedLocs.forEach(loc => {
      output.push([`▼▼▼ ${loc} ▼▼▼`, '', '', '', '', '', '', '', '', '']);
      const dayRank = { '月':1, '火':2, '水':3, '木':4, '金':5, '土':6, '日':7, '祝':8 };
      locKeys[loc].sort((a, b) => {
        const da = a.title.split('／')[2];
        const db = b.title.split('／')[2];
        return (dayRank[da] || 9) - (dayRank[db] || 9);
      });

      locKeys[loc].forEach(g => {
        const count = g.dates.length;
        const subH = g.hours * count;
        const subC = subH * g.wage;
        totalH += subH; totalCost += subC; totalCount += count;

        if (count === 1) {
          output.push([
            `${g.loc} 単独日`, 
            Utilities.formatDate(g.dates[0].date, "JST", "MM/dd"), 
            g.start, g.end, 
            'なし', g.dates[0].display, 
            g.wage, g.hours, g.hours * g.wage, ''
          ]);
        } else {
          const pattern = g.dates.map(d => {
            const w = Math.floor((d.date.getDate() - 1) / 7) + 1;
            const dw = DAYS_OF_WEEK_JP[d.date.getDay()];
            return `第${w}週${dw}`;
          }).join('\n');
          const dateList = g.dates.map(d => d.display).join('\n');
          
          output.push([
            g.title, `${monthStart}～${monthEnd}`, 
            g.start, g.end, 
            pattern, dateList, 
            g.wage, subH, subC, ''
          ]);
        }
      });
    });

    const header = [['シフトタイトル', '期間', '開始時間', '終了時間', '繰り返し曜日', '該当日', '時給', '募集時間', 'コスト', '対応済']];
    const footer = [['', '', '', '', '', '', '最終合計', totalH, totalCost, '']];
    const finalData = header.concat(output).concat(footer);

    const outRange = sheet.getRange(START_ROW, 13, finalData.length, 10);
    sheet.getRange(START_ROW, 13, sheet.getMaxRows(), 10).clearContent(); 
    outRange.setValues(finalData);
    outRange.setVerticalAlignment('top');

    const summaryBox = [
      ['追加拠点数', sortedLocs.length],
      ['追加枠数', totalCount],
      ['総時間', totalH],
      ['総額', totalCost]
    ];
    const sRow = START_ROW + finalData.length + 2;
    sheet.getRange(sRow, 13, 4, 2).setValues(summaryBox).setNumberFormat('#,##0');
    sheet.getRange(sRow + 3, 14).setNumberFormat('"¥"#,##0');

    finalData.forEach((r, i) => {
      if (i > 0 && i < finalData.length - 1 && !r[0].startsWith('▼▼▼')) {
        sheet.getRange(START_ROW + i, 22).insertCheckboxes();
      }
    });

    book.toast('募集リストの作成が完了しました。', '完了', 3);

  } catch (e) {
    ui.alert('エラー: ' + e.message);
  }
}

/**
 * ====================================================================
 * ２診要望一覧へのデータ転記ロジック
 * ====================================================================
 */
function append2ndConsultationRequests(sourceSheet) {
  const ss = sourceSheet.getParent();
  const masterIdMap = getClinicIdMap();
  
  let targetYear;
  const sheetNameMatch = sourceSheet.getName().match(/^(\d{4})\/(\d{1,2})$/);
  if (sheetNameMatch) {
    targetYear = parseInt(sheetNameMatch[1], 10);
  } else {
    const nextMonth = new Date();
    nextMonth.setMonth(nextMonth.getMonth() + 1);
    targetYear = nextMonth.getFullYear();
  }

  const data = sourceSheet.getDataRange().getValues();
  const dateHeaders = data[0]; 
  const timeHeaders = data[2]; 
  
  const records = [];
  let grandTotalCounts = 0;
  
  for (let i = 3; i < data.length; i++) {
    let baseLoc = data[i][0];
    if (!baseLoc) break; 
    
    if (baseLoc === '千葉NT') baseLoc = '千葉ニュータウン中央';
    
    let rawDept = String(data[i][1] || '小児科');
    let dept = rawDept.includes('内科') ? '内科' : '小児科';
    
    let locName = baseLoc;
    if (baseLoc.includes('亀有') || baseLoc.includes('北葛西')) {
      locName = `${baseLoc}（${dept}）`; 
    }
    
    const searchKey = _normalizeForIdSearch(locName);
    const fallbackKey = _normalizeForIdSearch(baseLoc);
    const clinicId = masterIdMap[searchKey] || masterIdMap[fallbackKey] || '';

    for (let j = 2; j < data[i].length; j++) {
      const val = parseInt(data[i][j], 10);
      if (!val || val < 1) continue;
      
      if (typeof data[i][j] === 'string' && /[a-zA-Z]/.test(data[i][j])) continue;
      
      const dateVal = dateHeaders[j];
      if (!dateVal) continue;
      
      let date = new Date(dateVal);
      date.setFullYear(targetYear);
      if (isNaN(date.getTime())) continue;
      
      let timeSlot = timeHeaders[j];
      let hours = 0;
      let start = '', end = '';
      
      if (timeSlot === '午前') {
        if (val === 2) { start = '10:00'; end = '12:00'; hours = 2; }
        else if (val === 3) { start = '10:00'; end = '13:00'; hours = 3; }
        else if (val === 4) { start = '09:00'; end = '13:00'; hours = 4; }
      } else if (timeSlot === '午後') {
        if (val === 2) { start = '15:00'; end = '17:00'; hours = 2; }
        else if (val === 3) { start = '15:00'; end = '18:00'; hours = 3; }
      } else if (timeSlot === '夜間') {
        start = '18:00'; end = '21:00'; hours = 3;
      }
      
      if (hours === 0) continue;
      
      records.push({
        loc: locName,
        dateObj: date,
        dateStr: Utilities.formatDate(date, "JST", "yyyy/MM/dd"),
        start: start,
        end: end,
        hours: hours,
        clinicId: clinicId
      });
      
      grandTotalCounts += 1;
    }
  }
  
  if (records.length === 0) return;
  
  records.sort((a, b) => {
    if (a.dateObj.getTime() !== b.dateObj.getTime()) {
      return a.dateObj.getTime() - b.dateObj.getTime();
    }
    if (a.start !== b.start) {
      return a.start.localeCompare(b.start);
    }
    return a.loc.localeCompare(b.loc);
  });
  
  const outputValues = [];
  const outputColors = [];
  
  for (let i = 0; i < records.length; i++) {
    const r = records[i];
    const isLastRowOfAll = (i === records.length - 1);
    
    outputValues.push([
      r.loc,
      r.dateStr,
      r.start,
      r.end,
      r.hours,
      isLastRowOfAll ? grandTotalCounts : '',
      r.clinicId
    ]);
    
    outputColors.push([
      null, null, null, null, null,
      isLastRowOfAll ? null : '#cccccc',
      null
    ]);
  }
  
  let targetSheet = ss.getSheetByName('２診要望一覧');
  if (!targetSheet) {
    targetSheet = ss.insertSheet('２診要望一覧');
  }
  
  if (targetSheet.getLastRow() === 0) {
    const headerRange = targetSheet.getRange(1, 1, 1, 7);
    headerRange.setValues([['拠点', '日付', '開始時間', '終了時間', '募集時間', '月間要望数', 'clinicID']]);
    headerRange.setHorizontalAlignment('left');
  }
  
  const lastRow = targetSheet.getLastRow();
  const dataRange = targetSheet.getRange(lastRow + 1, 1, outputValues.length, 7);
  
  dataRange.setValues(outputValues);
  dataRange.setBackgrounds(outputColors); 
  dataRange.setHorizontalAlignment('left'); 
  targetSheet.getRange(lastRow + 1, 5, outputValues.length, 2).setNumberFormat('0');
}

/**
 * 外部マスタからIDマップを取得する関数
 */
function getClinicIdMap() {
  const map = {};
  try {
    const masterUrl = 'https://docs.google.com/spreadsheets/d/14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs/edit';
    const masterSs = SpreadsheetApp.openByUrl(masterUrl);
    const masterSheet = masterSs.getSheetByName('拠点名');
    
    if (!masterSheet) return map;
    
    const data = masterSheet.getDataRange().getValues();
    if (data.length < 2) return map;
    
    let officialCol = -1, idCol = -1;
    let yureCols = [];
    
    for (let c = 0; c < data[0].length; c++) {
      const header = String(data[0][c]).trim();
      if (header === '正規記載') officialCol = c;
      if (header === 'クリニックNo') idCol = c;
      if (header.includes('表記揺れ')) yureCols.push(c);
    }
    
    if (officialCol === -1 || idCol === -1) return map;
    
    for (let i = 1; i < data.length; i++) {
      let officialName = data[i][officialCol];
      let clinicId = data[i][idCol];
      if (!officialName || !clinicId) continue;
      
      map[_normalizeForIdSearch(officialName)] = clinicId;
      
      yureCols.forEach(c => {
        let yureName = data[i][c];
        if (yureName) map[_normalizeForIdSearch(yureName)] = clinicId;
      });
    }
  } catch (e) {
    console.warn('IDマスタ取得エラー: ' + e.message);
  }
  return map;
}

/**
 * ID検索用の表記揺れ補正関数
 */
function _normalizeForIdSearch(rawName) {
  if (!rawName) return "";
  let name = rawName.toString();
  if (name.normalize) name = name.normalize('NFKC'); 
  
  return name.replace(/\s/g, "").replace(/ヶ/g, "ケ").trim();
}