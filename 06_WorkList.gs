/**
 * ====================================================================
 * 作業リスト作成ロジック
 * ====================================================================
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

    const fiscalYear = (targetMonth <= 3) ? targetYear - 1 : targetYear;
    initializeWageData();

    const data = sheet.getDataRange().getValues();
    const dateHeaders = data[0];
    const timeHeaders = data[2];
    const output = [['拠点名', '該当日', '曜日', '曜日週数', '時間帯', '時給', '対応済', '', '日給', '勤務時間']];
    
    let totalReqH = 0, totalReqW = 0; 
    let totalExH = 0, totalExW = 0;   

    for (let i = 3; i < data.length; i++) {
      let loc = data[i][0];
      if (!loc) continue;
      if (loc === '千葉NT') loc = '千葉ニュータウン中央';
      let dept = data[i][1] || '小児科'; 

      for (let j = 2; j < data[i].length; j++) {
        const rawVal = data[i][j];
        if (!rawVal) continue;

        const valStr = String(rawVal).toLowerCase();
        const valNum = parseInt(valStr, 10);
        if (isNaN(valNum) || valNum < 1) continue;

        const date = new Date(dateHeaders[j]);
        date.setFullYear(targetYear); 

        const dateStr = Utilities.formatDate(date, "JST", "MM/dd");
        const holidayStr = Utilities.formatDate(date, "JST", "yyyy-MM-dd");
        const dayOfWeek = HOLIDAYS_LIST.includes(holidayStr) ? '祝' : DAYS_OF_WEEK_JP[date.getDay()];
        const weekNum = `第${Math.floor((date.getDate() - 1) / 7) + 1}週`;

        let timeSlot = "";
        let hours = valNum; 
        let isSplit17to20 = false;

        if (timeHeaders[j] === '午前') {
          if (valNum === 2) timeSlot = '10:00~12:00';
          else if (valNum === 3) timeSlot = '10:00~13:00';
          else if (valNum === 4) timeSlot = '09:00~13:00';
          else timeSlot = `午前(${valNum}h)`;
        } else if (timeHeaders[j] === '午後') {
          if (valNum === 1) { timeSlot = '17:00~20:00'; hours = 3; isSplit17to20 = true; } 
          else if (valNum === 2) { timeSlot = '15:00~17:00'; }
          else if (valNum === 3) { timeSlot = '15:00~18:00'; }
          else timeSlot = `午後(${valNum}h)`;
        } else if (timeHeaders[j] === '夜間') {
          if (valNum === 2) { timeSlot = '18:00~20:00'; }
          else if (valNum === 3) { timeSlot = '18:00~21:00'; }
          else timeSlot = `夜間(${valNum}h)`;
        }
        
        if (hours === 0) continue;

        const dayType = (dayOfWeek === '祝' || dayOfWeek === '日' || dayOfWeek === '土') ? 'hol' : 'wd';
        const clinicWages = getClinicWages(loc);
        let targetRates = null;

        if (clinicWages && clinicWages.length > 0) {
          let targetDeptData = clinicWages.find(w => w.department.includes(dept)) || clinicWages[0];
          targetRates = fiscalYear >= 2026 ? targetDeptData.rates.y2026 : targetDeptData.rates.y2025;
        }

        let wageDisplay = "";
        let dailyPay = 0;

        if (isSplit17to20) {
          let pmWage = targetRates ? (Number(targetRates[`${dayType}_pm`]) || 0) : 0;
          let ntWage = targetRates ? (Number(targetRates[`${dayType}_nt`]) || 0) : 0;
          wageDisplay = pmWage + " / " + ntWage; 
          dailyPay = (pmWage * 1) + (ntWage * 2); 
        } else {
          let timeType = 'am';
          if (timeHeaders[j] === '午後') timeType = 'pm';
          if (timeHeaders[j] === '夜間') timeType = 'nt';
          let wage = targetRates ? (Number(targetRates[`${dayType}_${timeType}`]) || 0) : 0;
          wageDisplay = wage;
          dailyPay = wage * hours;
        }

        totalReqH += hours;
        totalReqW += dailyPay;
        
        if (valStr.includes('f')) {
          totalExH += hours;
          totalExW += dailyPay;
        } else {
          output.push([loc, dateStr, dayOfWeek, weekNum, timeSlot, wageDisplay, '', '', dailyPay, hours]);
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
      sheet.getRange('K' + (summaryStartRow + 1)).setValue('今回募集枠数').setFontWeight('bold');
      sheet.getRange('L' + (summaryStartRow + 1)).setValue(output.length - 1);
      sheet.getRange('K' + (summaryStartRow + 2)).setValue('今回募集時間').setFontWeight('bold');
      sheet.getRange('L' + (summaryStartRow + 2)).setValue(totalReqH - totalExH);
      sheet.getRange('K' + (summaryStartRow + 3)).setValue('総額(募集分)').setFontWeight('bold');
      sheet.getRange('L' + (summaryStartRow + 3)).setValue(totalReqW - totalExW).setNumberFormat('"¥"#,##0');

      append2ndConsultationRequests(sheet);

      book.toast(`抽出完了：今回募集 ${totalReqH - totalExH}h (※充足済 ${totalExH}h除外)`, '完了', 5);
    } else {
      ui.alert('データがありません。');
    }
  } catch (e) {
    ui.alert('エラー: ' + e.message);
  }
}