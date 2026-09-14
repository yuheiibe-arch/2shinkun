/**
 * ====================================================================
 * 募集リスト作成 ＆ 定期・単独の振り分けロジック
 * ====================================================================
 */

function createRecruitmentListFromActiveSheet() {
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

    const expectedCounts = { '日': 0, '月': 0, '火': 0, '水': 0, '木': 0, '金': 0, '土': 0, '祝': 0 };
    const daysInMonth = new Date(targetYear, targetMonth, 0).getDate();
    for (let d = 1; d <= daysInMonth; d++) {
      const dObj = new Date(targetYear, targetMonth - 1, d);
      const dStr = Utilities.formatDate(dObj, "JST", "yyyy-MM-dd");
      if (HOLIDAYS_LIST.includes(dStr)) expectedCounts['祝']++;
      else expectedCounts[DAYS_OF_WEEK_JP[dObj.getDay()]]++;
    }

    const data = sheet.getDataRange().getValues();
    let startIdx = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === '拠点名' && data[i][1] === '該当日') { startIdx = i; break; }
    }
    if (startIdx === -1) throw new Error('左側に「作業リスト」が見つかりません。先に作業リストを作成してください。');

    const START_ROW = startIdx + 1; 
    const calendarEndIdx = startIdx;
    const groups = {};
    const dateHeaders = data[0]; // ★修正: 1行目の日付ヘッダーを正しく取得
    const timeHeaders = data[2];
    const fiscalYear = (targetMonth <= 3) ? targetYear - 1 : targetYear;
    if (typeof _wageMasterMap2025 === 'undefined' || !_wageMasterMap2025) initializeWageData(); 

    // カレンダーから直接データ取得
    for (let i = 3; i < calendarEndIdx; i++) {
      let loc = data[i][0];
      if (!loc) continue;
      if (loc === '千葉NT') loc = '千葉ニュータウン中央';
      let rawDept = data[i][1] || '小児科'; 
      let dept = rawDept.includes('内科') ? '内科' : '小児科';

      for (let j = 2; j < data[i].length; j++) {
        const rawVal = data[i][j];
        if (!rawVal) continue;
        const valStr = String(rawVal).toLowerCase();
        if (valStr.includes('f')) continue; // f付きは無視

        const valNum = parseInt(valStr, 10);
        if (isNaN(valNum) || valNum < 1) continue;

        // ★修正: 列番号の計算ではなく、1行目の日付データを直接読み取る
        const dateVal = dateHeaders[j];
        if (!dateVal) continue;
        const dObj = new Date(dateVal);
        dObj.setFullYear(targetYear);
        if (isNaN(dObj.getTime())) continue;

        const dateStr = Utilities.formatDate(dObj, "JST", "yyyy-MM-dd");
        const isHoliday = HOLIDAYS_LIST.includes(dateStr);
        const dayOfWeek = DAYS_OF_WEEK_JP[dObj.getDay()];
        const groupDay = isHoliday ? '祝' : dayOfWeek;
        const weekNum = `第${Math.floor((dObj.getDate() - 1) / 7) + 1}週`;

        let timeSlot = "", start = "", end = "", hours = valNum;
        let isSplit17to20 = false;
        let timeTypeForRate = "";

        if (timeHeaders[j] === '午前') {
          timeTypeForRate = 'am';
          if (valNum === 2) { timeSlot = '10:00~12:00'; start = '10:00'; end = '12:00'; }
          else if (valNum === 3) { timeSlot = '10:00~13:00'; start = '10:00'; end = '13:00'; }
          else if (valNum === 4) { timeSlot = '09:00~13:00'; start = '09:00'; end = '13:00'; }
          else { timeSlot = `午前(${valNum}h)`; start = '09:00'; end = '12:00'; }
        } else if (timeHeaders[j] === '午後') {
          timeTypeForRate = 'pm';
          if (valNum === 1) { timeSlot = '17:00~20:00'; start = '17:00'; end = '20:00'; hours = 3; isSplit17to20 = true; } 
          else if (valNum === 2) { timeSlot = '15:00~17:00'; start = '15:00'; end = '17:00'; }
          else { timeSlot = '15:00~18:00'; start = '15:00'; end = '18:00'; }
        } else if (timeHeaders[j] === '夜間') {
          timeTypeForRate = 'nt';
          if (valNum === 2) { timeSlot = '18:00~20:00'; start = '18:00'; end = '20:00'; }
          else { timeSlot = '18:00~21:00'; start = '18:00'; end = '21:00'; }
        }
        
        if (hours === 0) continue;

        const dayTypeForRate = (dayOfWeek === '祝' || dayOfWeek === '日' || dayOfWeek === '土') ? 'hol' : 'wd';
        const clinicWages = getClinicWages(loc);
        let targetRates = null;
        if (clinicWages && clinicWages.length > 0) {
          let targetDeptData = clinicWages.find(w => w.department.includes(dept)) || clinicWages[0];
          targetRates = fiscalYear >= 2026 ? targetDeptData.rates.y2026 : targetDeptData.rates.y2025;
        }

        let wageStr = "";
        let dailyPay = 0;
        if (isSplit17to20) {
          let pmWage = targetRates ? (Number(targetRates[`${dayTypeForRate}_pm`]) || 0) : 0;
          let ntWage = targetRates ? (Number(targetRates[`${dayTypeForRate}_nt`]) || 0) : 0;
          wageStr = pmWage + " / " + ntWage; 
          dailyPay = (pmWage * 1) + (ntWage * 2); 
        } else {
          let wage = targetRates ? (Number(targetRates[`${dayTypeForRate}_${timeTypeForRate}`]) || 0) : 0;
          wageStr = wage;
          dailyPay = wage * hours;
        }

        const simpleTime = timeSlot.replace('~', '-');
        const key = `${loc}-${timeSlot}-${groupDay}`;

        if (!groups[key]) {
          groups[key] = {
            loc, dept, title: `${loc}／${simpleTime}／${groupDay}`,
            groupDay: groupDay, start, end, hours, wageStr, dailyPay, dates: []
          };
        }
        groups[key].dates.push({ date: dObj, dateStr: dateStr, display: `${Utilities.formatDate(dObj, "JST", "MM/dd")}(${dayOfWeek})`, weekNum: weekNum });
      }
    }

    const output = [];
    const singleShiftsForCSV = []; 
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
      let hasRecurring = false;
      locKeys[loc].forEach(g => { 
        const expected = expectedCounts[g.groupDay];
        g.isWeekly = (g.dates.length === expected && expected > 0 && g.groupDay !== '祝');
        if (g.isWeekly) hasRecurring = true; 
      });

      if (hasRecurring) output.push([`▼▼▼ ${loc} ▼▼▼`, '', '', '', '', '', '', '', '', '']);

      const dayRank = { '月':1, '火':2, '水':3, '木':4, '金':5, '土':6, '日':7, '祝':8 };
      locKeys[loc].sort((a, b) => (dayRank[a.groupDay] || 9) - (dayRank[b.groupDay] || 9));

      locKeys[loc].forEach(g => {
        if (g.isWeekly) {
          const count = g.dates.length;
          const subH = g.hours * count;
          const subC = g.dailyPay * count; 
          totalH += subH; totalCost += subC; totalCount += count;

          const pattern = g.dates.map(d => d.weekNum + d.display.slice(-3)).join('\n');
          const dateList = g.dates.map(d => d.display).join('\n');
          
          output.push([ `【毎週】${g.title}`, `${monthStart}～${monthEnd}`, g.start, g.end, pattern, dateList, g.wageStr, subH, subC, '' ]);
        } else {
          let targetLoc = g.loc;
          if (targetLoc.includes('亀有') || targetLoc.includes('北葛西')) targetLoc = `${targetLoc}（${g.dept}）`; 
          const idData = getClinicIdData(targetLoc); 
          g.dates.forEach(dInfo => {
            singleShiftsForCSV.push({
              loc: targetLoc, dept: g.dept, dateStr: dInfo.dateStr, start: g.start, end: g.end,
              wageStr: g.wageStr, clinicId: idData.clinicId, departmentCode: idData.departmentCode
            });
          });
        }
      });
    });

    const clearRows = sheet.getMaxRows() - START_ROW + 1;
    if (clearRows > 0) {
      sheet.getRange(START_ROW, 13, clearRows, 10).clearContent(); 
      sheet.getRange(START_ROW, 22, clearRows, 1).removeCheckboxes();
    }

    if (output.length > 0) {
      const header = [['シフトタイトル', '期間', '開始時間', '終了時間', '繰り返し曜日', '該当日', '時給', '募集時間', 'コスト', '対応済']];
      const footer = [['', '', '', '', '', '', '最終合計', totalH, totalCost, '']];
      const finalData = header.concat(output).concat(footer);

      const outRange = sheet.getRange(START_ROW, 13, finalData.length, 10);
      outRange.setValues(finalData);
      outRange.setVerticalAlignment('top');

      const summaryBox = [
        ['追加拠点数', sortedLocs.length], ['追加枠数', totalCount],
        ['総時間', totalH], ['総額', totalCost]
      ];
      const sRow = START_ROW + finalData.length + 2;
      sheet.getRange(sRow, 13, 4, 2).setValues(summaryBox).setNumberFormat('#,##0');
      sheet.getRange(sRow + 3, 14).setNumberFormat('"¥"#,##0');

      finalData.forEach((r, i) => {
        if (i > 0 && i < finalData.length - 1 && !r[0].startsWith('▼▼▼')) sheet.getRange(START_ROW + i, 22).insertCheckboxes();
      });
    }

    if (singleShiftsForCSV.length > 0) {
      generateChunkedCSVSheets(book, targetYear, targetMonth, singleShiftsForCSV);
    }

    book.toast('募集リスト（定期）と CSVシート（不定期/単独） の作成が完了しました。', '完了', 5);
  } catch (e) {
    ui.alert('エラー: ' + e.message);
  }
}