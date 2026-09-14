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