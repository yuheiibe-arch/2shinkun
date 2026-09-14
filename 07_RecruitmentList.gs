/**
 * ====================================================================
 * 2. 募集リスト作成 ＆ 定期・単独の振り分け（作業リストから読み込み）
 * 【北葛西20時制限 ＆ 内科2診アラート実装 版】
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
      if (data[i][0] === '拠点名' && data[i][1] === '該当日') {
        startIdx = i;
        break;
      }
    }
    if (startIdx === -1) throw new Error('左側に「作業リスト」が見つかりません。先に作業リストを作成してください。');

    const START_ROW = startIdx + 1; 
    const workData = data.slice(startIdx + 1); 
    const groups = {};
    
    // ★内科の2診シフトを検知・記録するためのセット
    const internalMedicineAlerts = new Set(); 

    workData.forEach(row => {
      const [loc, dateVal, dayOfWeekRaw, weekNumRaw, timeSlot, wageStr, , deptRaw, dailyPayVal, hoursVal] = row;
      if (!loc || !dateVal || !wageStr) return; 

      let dateObj;
      if (dateVal instanceof Date) {
        dateObj = new Date(dateVal);
        dateObj.setFullYear(targetYear);
      } else {
        dateObj = new Date(`${targetYear}/${dateVal}`);
      }
      if (isNaN(dateObj.getTime())) return;

      const dateStr = Utilities.formatDate(dateObj, "JST", "yyyy-MM-dd");
      const isHoliday = HOLIDAYS_LIST.includes(dateStr);
      const dayOfWeek = DAYS_OF_WEEK_JP[dateObj.getDay()];
      const groupDay = isHoliday ? '祝' : dayOfWeek;
      const weekNum = `第${Math.floor((dateObj.getDate() - 1) / 7) + 1}週`;

      const timeStr = String(timeSlot);
      let simpleTime = timeStr.replace('~', '-');
      
      let dept = deptRaw ? String(deptRaw).trim() : '小児科';

      // ★内科シフトの場合はアラート用に記録しておく
      if (dept.includes('内科')) {
        internalMedicineAlerts.add(`${loc} (${Utilities.formatDate(dateObj, "JST", "MM/dd")})`);
      }

      let start = '', end = '';
      let hours = Number(hoursVal) || 0;
      let dailyPay = Number(dailyPayVal) || 0;

      const m = timeStr.match(/(\d{2}:\d{2})~(\d{2}:\d{2})/);
      if (m) {
        start = m[1]; end = m[2]; 
      } else {
        const hourMatch = timeStr.match(/(\d+)h/);
        if (hourMatch) { start = '09:00'; end = `(${hourMatch[1]}h)`; }
      }

      // ★「北葛西」のみ「20:00営業終了」制限を適用（亀有は除外）
      if (loc.includes('北葛西')) {
        if (end === '21:00') {
          end = '20:00'; // 21時を20時で打ち切り
          simpleTime = simpleTime.replace('21:00', '20:00');
          // 18:00-21:00(3h)だった場合は2hに補正し、金額も2/3に調整
          if (start === '18:00' && hours === 3) {
            hours = 2;
            dailyPay = Math.floor((dailyPay / 3) * 2);
          }
        }
      }
      
      const key = `${loc}-${dept}-${simpleTime}-${groupDay}`;

      if (!groups[key]) {
        let titleName = loc;
        if (dept !== '小児科') titleName = `${loc}(${dept})`;
        
        groups[key] = {
          loc, dept, title: `${titleName}／${simpleTime}／${groupDay}`,
          groupDay: groupDay, start, end, 
          hours: hours,
          wageStr: wageStr, dailyPay: dailyPay,
          dates: []
        };
      }
      
      groups[key].dates.push({ 
        date: dateObj, 
        dateStr: dateStr, 
        display: `${Utilities.formatDate(dateObj, "JST", "MM/dd")}(${dayOfWeek})`, 
        weekNum: weekNum 
      });
    });

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

      if (hasRecurring) {
        output.push([`▼▼▼ ${loc} ▼▼▼`, '', '', '', '', '', '', '', '', '']);
      }

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
          
          output.push([
            `【毎週】${g.title}`, `${monthStart}～${monthEnd}`, 
            g.start, g.end, 
            pattern, dateList, 
            g.wageStr, subH, subC, '' 
          ]);
        } else {
          let clinicId = "";
          let deptCode = "";

          // 亀有・北葛西のハードコード処理
          if (g.loc.includes('亀有')) {
            clinicId = g.dept.includes('内科') ? '13' : '11';
            deptCode = clinicId;
          } else if (g.loc.includes('北葛西')) {
            clinicId = g.dept.includes('内科') ? '6' : '4';
            deptCode = clinicId;
          } else {
            // その他の動的取得
            let idData = getClinicIdData(`${g.loc}（${g.dept}）`);
            if (!idData || !idData.clinicId) {
              idData = getClinicIdData(g.loc);
            }
            clinicId = idData ? idData.clinicId : "";
            deptCode = idData ? idData.departmentCode : "";
          }

          g.dates.forEach(dInfo => {
            singleShiftsForCSV.push({
              loc: g.loc,
              dept: g.dept,
              dateStr: dInfo.dateStr,
              start: g.start,
              end: g.end,
              wageStr: g.wageStr, 
              clinicId: clinicId,
              departmentCode: deptCode
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
        ['UI出力拠点数', sortedLocs.length],
        ['UI出力枠数', totalCount],
        ['定期(UI)時間', totalH],
        ['定期総額', totalCost]
      ];
      const sRow = START_ROW + finalData.length + 2;
      sheet.getRange(sRow, 13, 4, 2).setValues(summaryBox).setNumberFormat('#,##0');
      sheet.getRange(sRow + 3, 14).setNumberFormat('"¥"#,##0');

      finalData.forEach((r, i) => {
        if (i > 0 && i < finalData.length - 1 && !r[0].startsWith('▼▼▼')) {
          sheet.getRange(START_ROW + i, 22).insertCheckboxes();
        }
      });
    }

    if (singleShiftsForCSV.length > 0) {
      generateChunkedCSVSheets(book, targetYear, targetMonth, singleShiftsForCSV);
    }

    // ★内科の2診が含まれていた場合のアラート表示
    if (internalMedicineAlerts.size > 0) {
      const alertMsg = Array.from(internalMedicineAlerts).slice(0, 15).join('\n') + 
                       (internalMedicineAlerts.size > 15 ? `\n...他 ${internalMedicineAlerts.size - 15} 件` : '');
      ui.alert(
        '⚠️【警告】内科の2診シフトが出力されています⚠️\n\n' +
        '内科の2診はほぼ発生しないはずですが、以下のシフトが「内科」として検知・出力されました。\n' +
        '意図した募集か必ずご確認ください。\n\n' +
        alertMsg
      );
    } else {
      book.toast('募集リスト（定期）と CSVシート（不定期/単独） の作成が完了しました。', '完了', 5);
    }

  } catch (e) {
    ui.alert('エラー: ' + e.message);
  }
}