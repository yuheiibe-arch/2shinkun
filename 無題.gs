function debugCheckExactSplit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const holidaySheet = ss.getSheetByName('祝日');
  
  const holidays = holidaySheet ? holidaySheet.getRange(1, 1, holidaySheet.getLastRow(), 1).getValues().flat().map(d => {
    const dt = new Date(d);
    return !isNaN(dt) ? Utilities.formatDate(dt, "JST", "yyyy-MM-dd") : null;
  }).filter(d => d) : [];

  const DAYS_OF_WEEK_JP = ['日', '月', '火', '水', '木', '金', '土'];
  let resultText = "【定期/CSV振り分け 完全ファクトチェック】\n\n";

  sheets.forEach(sheet => {
    const match = sheet.getName().match(/^(\d{4})\/(\d{1,2})$/);
    if (!match) return;

    const targetYear = parseInt(match[1], 10);
    const targetMonth = parseInt(match[2], 10);
    const expectedCounts = { '日': 0, '月': 0, '火': 0, '水': 0, '木': 0, '金': 0, '土': 0, '祝': 0 };
    const daysInMonth = new Date(targetYear, targetMonth, 0).getDate();
    
    for (let d = 1; d <= daysInMonth; d++) {
      const dObj = new Date(targetYear, targetMonth - 1, d);
      const dStr = Utilities.formatDate(dObj, "JST", "yyyy-MM-dd");
      if (holidays.includes(dStr)) expectedCounts['祝']++;
      else expectedCounts[DAYS_OF_WEEK_JP[dObj.getDay()]]++;
    }

    const data = sheet.getDataRange().getValues();
    let startIdx = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === '拠点名' && data[i][1] === '該当日') { startIdx = i; break; }
    }
    const calendarEndIdx = startIdx !== -1 ? startIdx : data.length;

    let totalHours = 0;
    let fHours = 0;
    let missingMasterHours = 0;

    const groups = {};
    const timeHeaders = data[2];

    // ★大元カレンダーから直接すべてを拾う
    for (let i = 3; i < calendarEndIdx; i++) {
      const loc = data[i][0];
      if (!loc) continue;

      for (let j = 2; j < data[i].length; j++) {
        const rawVal = data[i][j];
        if (!rawVal) continue;
        
        const valStr = String(rawVal).toLowerCase();
        const valNum = parseInt(valStr, 10);
        if (isNaN(valNum) || valNum < 1) continue;

        let h = valNum;
        if (timeHeaders[j] === '午後' && valNum === 1) h = 3; // 17:00-20:00対応

        totalHours += h;

        if (valStr.includes('f')) {
          fHours += h;
          continue; // f付きはここで終了
        }

        // fなし（今回募集）をグループ化して定期/CSVを判定する
        const dObj = new Date(targetYear, targetMonth - 1, j - 1);
        const dStr = Utilities.formatDate(dObj, "JST", "yyyy-MM-dd");
        const groupDay = holidays.includes(dStr) ? '祝' : DAYS_OF_WEEK_JP[dObj.getDay()];
        const key = `${loc}-${timeHeaders[j]}-${groupDay}`;

        if (!groups[key]) {
          groups[key] = { groupDay: groupDay, totalH: 0, count: 0 };
        }
        groups[key].totalH += h;
        groups[key].count++;
      }
    }

    let regularHours = 0;
    let csvHours = 0;

    // グループから定期（毎週）と単独（CSV）に振り分け
    Object.values(groups).forEach(g => {
      const expected = expectedCounts[g.groupDay];
      const isWeekly = (g.count === expected && expected > 0 && g.groupDay !== '祝');

      if (isWeekly) {
        regularHours += g.totalH;
      } else {
        csvHours += g.totalH;
      }
    });

    const actualHours = totalHours - fHours;
    const splitTotal = regularHours + csvHours;

    resultText += `■ ${sheet.getName()}\n`;
    resultText += `  総2診時間(カレンダー): ${totalHours} h\n`;
    resultText += `  ｆ付き(充足済除外)  : ${fHours} h\n`;
    resultText += `  --------------------------------\n`;
    resultText += `  募集対象(実質)      : ${actualHours} h\n`;
    resultText += `    ┣ UIリスト(定期)  : ${regularHours} h\n`;
    resultText += `    ┗ CSV出力(単独等) : ${csvHours} h\n`;
    
    if (actualHours !== splitTotal) {
      resultText += `  ★警告: 募集対象と振り分け合計に ${actualHours - splitTotal}h のズレがあります！\n\n`;
    } else {
      resultText += `  (内訳一致OK)\n\n`;
    }
  });

  console.log(resultText);
}