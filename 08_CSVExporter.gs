/**
 * ====================================================================
 * CSV分割出力・IDマスタ取得ロジック
 * 【連続シフトの時給分割(split)自動生成版】
 * ====================================================================
 */

let _localClinicIdCache = null;

function getClinicIdData(rawClinicName) {
  if (!_localClinicIdCache) {
    _localClinicIdCache = {};
    try {
      const masterUrl = 'https://docs.google.com/spreadsheets/d/14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs/edit';
      const masterSs = SpreadsheetApp.openByUrl(masterUrl);
      const masterSheet = masterSs.getSheetByName('拠点名');
      if (masterSheet) {
        const data = masterSheet.getDataRange().getValues();
        const headers = data[0];
        let officialCol = -1, idCol = -1, deptCodeCol = -1;
        const yureCols = [];
        headers.forEach((h, idx) => {
          const title = String(h).trim();
          if (title === '正規記載') officialCol = idx;
          if (title === 'クリニックNo') idCol = idx;
          if (title === 'departmentCode') deptCodeCol = idx;
          if (title.includes('表記揺れ')) yureCols.push(idx);
        });

        if (officialCol !== -1 && idCol !== -1) {
          const norm = (name) => name ? String(name).normalize('NFKC').replace(/\s/g, "").replace(/ヶ/g, "ケ").trim() : "";
          for (let i = 1; i < data.length; i++) {
            const officialName = data[i][officialCol];
            const clinicId = data[i][idCol];
            const deptCode = (deptCodeCol !== -1 && data[i][deptCodeCol]) ? data[i][deptCodeCol] : clinicId;
            if (!officialName || !clinicId) continue;
            
            const mapData = { clinicId: clinicId, departmentCode: deptCode };
            _localClinicIdCache[norm(officialName)] = mapData;
            yureCols.forEach(c => {
              if (data[i][c]) _localClinicIdCache[norm(data[i][c])] = mapData;
            });
          }
        }
      }
    } catch(e) {
      console.warn('独立IDマスタ取得エラー: ' + e.message);
    }
  }
  
  const cleanLoc = rawClinicName ? String(rawClinicName).normalize('NFKC').replace(/\s/g, "").replace(/ヶ/g, "ケ").trim() : "";
  return _localClinicIdCache[cleanLoc] || { clinicId: "", departmentCode: "" };
}

function generateChunkedCSVSheets(book, year, month, singleShifts) {
  const MAX_LIMIT = 900;
  const dailyMap = {};
  
  singleShifts.forEach(s => {
    const key = s.loc + "_" + s.dateStr;
    if (!dailyMap[key]) dailyMap[key] = [];
    dailyMap[key].push(s);
  });

  const rowsByDate = {};
  Object.values(dailyMap).forEach(dayShifts => {
    const targetDate = dayShifts[0].dateStr;
    if (!rowsByDate[targetDate]) rowsByDate[targetDate] = [];

    // ==============================================================
    // ★ 新処理：結合履歴(segments)から自動でCSVスプリットを生成
    // 決め打ちの時間を廃止し、どんな時間の組み合わせでも対応可能に
    // ==============================================================
    dayShifts.forEach(s => {
      let split1 = "", split2 = "", split3 = "";
      let wage1 = "", wage2 = "", wage3 = "";
      let finalWage = s.wageStr;

      if (s.segments && s.segments.length > 1) {
        split1 = s.segments[1].start;
        wage1 = s.segments[0].wage;
        if (s.segments.length > 2) {
           split2 = s.segments[2].start;
           wage2 = s.segments[1].wage;
           if (s.segments.length > 3) {
             split3 = s.segments[3].start;
             wage3 = s.segments[2].wage;
             finalWage = s.segments[3].wage;
           } else {
             finalWage = s.segments[2].wage;
           }
        } else {
           finalWage = s.segments[1].wage;
        }
      }

      rowsByDate[targetDate].push({
        departmentCode: s.departmentCode || s.clinicId, 
        email: "doctor@example.com", 
        date: targetDate,
        startTime: s.start, 
        splitDateTime1: split1, 
        splitDateTime2: split2, 
        splitDateTime3: split3,
        endTime: s.end, 
        splitHourlyWage1: wage1, 
        splitHourlyWage2: wage2, 
        splitHourlyWage3: wage3,
        hourlyWage: finalWage, 
        numberOfRooms: 1, 
        delete: "", 
        comment: `${s.dept}外来`
      });
    });
    // ==============================================================
  });

  const sortedDates = Object.keys(rowsByDate).sort();
  const chunks = [];
  let currentChunk = [];

  sortedDates.forEach(date => {
    const dailyRows = rowsByDate[date];
    if (currentChunk.length > 0 && (currentChunk.length + dailyRows.length) > MAX_LIMIT) {
      chunks.push(currentChunk);
      currentChunk = [];
    }
    currentChunk.push(...dailyRows);
  });
  if (currentChunk.length > 0) chunks.push(currentChunk);

  const header = [
    "departmentCode", "email", "date", "startTime", "splitDateTime1", "splitDateTime2", "splitDateTime3", "endTime",
    "splitHourlyWage1", "splitHourlyWage2", "splitHourlyWage3", "hourlyWage", "numberOfRooms", "delete", "comment",
    "isSplitDateTime1BreakTime", "isSplitDateTime2BreakTime", "isSplitDateTime3BreakTime", "isStartTimeBreakTime",
    "isPublished", "isDoubleRecruitment", 
    "doctorSubsidy1Title", "doctorSubsidy1Money", "doctorSubsidy2Title", "doctorSubsidy2Money", 
    "doctorSubsidy3Title", "doctorSubsidy3Money", "doctorSubsidy4Title", "doctorSubsidy4Money", 
    "doctorSubsidy5Title", "doctorSubsidy5Money"
  ];

  chunks.forEach((chunkData, index) => {
    const suffix = chunks.length > 1 ? `_Part${index + 1}` : "";
    const sheetName = `${year}.${('0' + month).slice(-2)}CSV${suffix}`;
    let targetSheet = book.getSheetByName(sheetName);
    
    if (targetSheet) {
      targetSheet.clearContents();
      targetSheet.showSheet(); 
    } else {
      const templateSheet = book.getSheetByName("CSV");
      if (templateSheet) {
        targetSheet = templateSheet.copyTo(book);
        targetSheet.setName(sheetName);
        targetSheet.clearContents();
        targetSheet.showSheet(); 
      } else {
        targetSheet = book.insertSheet(sheetName);
      }
    }

    const csvRows = [header];
    chunkData.forEach(r => {
      csvRows.push([
        r.departmentCode, r.email, r.date, r.startTime, r.splitDateTime1, r.splitDateTime2, r.splitDateTime3, r.endTime,
        r.splitHourlyWage1, r.splitHourlyWage2, r.splitHourlyWage3, r.hourlyWage, r.numberOfRooms, r.delete, r.comment,
        "", "", "", "FALSE", "TRUE", "TRUE", "", "", "", "", "", "", "", "", "", "" 
      ]);
    });

    targetSheet.getRange(1, 1, csvRows.length, header.length).setValues(csvRows);
  });
}