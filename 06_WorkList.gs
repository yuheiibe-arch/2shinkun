/**
 * ====================================================================
 * 1. 作業リスト作成ロジック（Step1）
 * 【カレンダー1行目から読み取り ＆ 金額列誤爆防止 ＆ 午後1=1h】
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
    
    // 画像の通り、ローカルのカレンダーは1行目から始まっているため、決め打ちで取得
    const dateHeaders = data[0]; 
    const timeHeaders = data[2]; 
    const dataStartIdx = 3; // 実データは4行目(インデックス3)から

    const output = [['拠点名', '該当日', '曜日', '曜日週数', '時間帯', '時給', '対応済', '診療科', '日給', '勤務時間']];
    let totalReqH = 0, totalReqW = 0, totalExH = 0, totalExW = 0;   

    for (let i = dataStartIdx; i < data.length; i++) {
      let loc = data[i][0];
      
      // 空欄の行にぶつかったら、そこがカレンダーの終わりなので読み込みをストップ
      if (!loc) break; 
      
      if (loc === '千葉NT') loc = '千葉ニュータウン中央';
      
      let dept = data[i][1] ? String(data[i][1]).trim() : '小児科'; 

      for (let j = 2; j < data[i].length; j++) {
        // ★【安全装置】午前・午後・夜間 以外の列（右側にある金額列など）は絶対に無視する
        if (timeHeaders[j] !== '午前' && timeHeaders[j] !== '午後' && timeHeaders[j] !== '夜間') continue;

        const rawVal = data[i][j];
        if (!rawVal) continue;
        const valStr = String(rawVal).toLowerCase();
        
        const matchNum = valStr.match(/([0-9]+(\.[0-9]+)?)/);
        if (!matchNum) continue;
        const valNum = parseFloat(matchNum[1]);
        if (valNum <= 0) continue;

        const dateVal = dateHeaders[j];
        if (!dateVal) continue;
        const dateObj = new Date(dateVal);
        dateObj.setFullYear(targetYear);
        if (isNaN(dateObj.getTime())) continue;

        const dateStr = Utilities.formatDate(dateObj, "JST", "MM/dd");
        const holidayStr = Utilities.formatDate(dateObj, "JST", "yyyy-MM-dd");
        const dayOfWeek = HOLIDAYS_LIST.includes(holidayStr) ? '祝' : DAYS_OF_WEEK_JP[dateObj.getDay()];
        const weekNum = `第${Math.floor((dateObj.getDate() - 1) / 7) + 1}週`;

        let timeSlot = "";
        let hours = valNum; 

        if (timeHeaders[j] === '午前') {
          if (valNum === 2) { timeSlot = '10:00~12:00'; hours = 2; }
          else if (valNum === 3) { timeSlot = '10:00~13:00'; hours = 3; }
          else if (valNum === 4) { timeSlot = '09:00~13:00'; hours = 4; }
          else { timeSlot = `午前(${valNum}h)`; } 
        } else if (timeHeaders[j] === '午後') {
          // ★午後1を正しい時間（17:00-18:00 1h）に修正
          if (valNum === 1) { timeSlot = '17:00~18:00'; hours = 1; }
          else if (valNum === 2) { timeSlot = '15:00~17:00'; hours = 2; }
          else if (valNum === 3) { timeSlot = '15:00~18:00'; hours = 3; }
          else { timeSlot = `午後(${valNum}h)`; }
        } else if (timeHeaders[j] === '夜間') {
          if (valNum === 2) { timeSlot = '18:00~20:00'; hours = 2; }
          else if (valNum === 3) { timeSlot = '18:00~21:00'; hours = 3; }
          else { timeSlot = `夜間(${valNum}h)`; }
        }

        const dayType = (dayOfWeek === '祝' || dayOfWeek === '日' || dayOfWeek === '土') ? 'hol' : 'wd';
        const clinicWages = getClinicWages(loc);
        let targetRates = null;

        if (clinicWages && clinicWages.length > 0) {
          let targetDeptData = clinicWages.find(w => w.department.includes(dept)) || clinicWages[0];
          targetRates = fiscalYear >= 2026 ? targetDeptData.rates.y2026 : targetDeptData.rates.y2025;
        }

        let timeType = 'am';
        if (timeHeaders[j] === '午後') timeType = 'pm';
        if (timeHeaders[j] === '夜間') timeType = 'nt';
        
        let wage = targetRates ? (Number(targetRates[`${dayType}_${timeType}`]) || 0) : 0;
        let dailyPay = wage * hours;

        if (valStr.includes('f')) {
          totalExH += hours;
          totalExW += dailyPay;
        } else {
          totalReqH += hours;
          totalReqW += dailyPay;
          output.push([loc, dateStr, dayOfWeek, weekNum, timeSlot, wage, '', dept, dailyPay, hours]);
        }
      }
    }

    if (output.length > 1) {
      // 作業リストをカレンダーの下に追記
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
      sheet.getRange('L' + (summaryStartRow + 2)).setValue(totalReqH);
      sheet.getRange('K' + (summaryStartRow + 3)).setValue('総額(募集分)').setFontWeight('bold');
      sheet.getRange('L' + (summaryStartRow + 3)).setValue(totalReqW).setNumberFormat('"¥"#,##0');

      // ２診要望一覧へ転記（先ほど完成した最新ロジックが呼ばれます）
      append2ndConsultationRequests(sheet);

      book.toast(`作業リスト生成完了：今回募集 ${totalReqH}h`, '完了', 5);
    } else {
      ui.alert('データがありません。');
    }
  } catch (e) {
    ui.alert('エラー: ' + e.message);
  }
}