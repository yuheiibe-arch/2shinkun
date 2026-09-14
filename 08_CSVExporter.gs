/**
 * ====================================================================
 * CSV分割出力・IDマスタ取得ロジック
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
  const MAX_LIMIT = 190;
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

    const pm = dayShifts.find(s => s.start === '17:00' && s.end === '20:00');
    const ntTarget = dayShifts.find(s => s.start === '18:00' && (s.end === '20:00' || s.end === '21:00'));

    if (pm && ntTarget) {
      const wages = String(pm.wageStr).split('/');
      const w1 = wages[0] ? wages[0].trim() : "";
      const w2 = wages[1] ? wages[1].trim() : w1;
      
      rowsByDate[targetDate].push({
        departmentCode: pm.departmentCode || pm.clinicId, 
        email: "doctor@example.com", date: targetDate,
        startTime: "17:00", splitDateTime1: "18:00", splitDateTime2: "", splitDateTime3: "",
        endTime: ntTarget.end, splitHourlyWage1: w1, splitHourlyWage2: "", splitHourlyWage3: "",
        hourlyWage: w2, numberOfRooms: 1, delete: "", comment: `${pm.dept}外来`
      });
      dayShifts.forEach(s => {
        if (s !== pm && s !== ntTarget) {
          rowsByDate[targetDate].push({
            departmentCode: s.departmentCode || s.clinicId, email: "doctor@example.com", date: targetDate,
            startTime: s.start, splitDateTime1: "", splitDateTime2: "", splitDateTime3: "",
            endTime: s.end, splitHourlyWage1: "", splitHourlyWage2: "", splitHourlyWage3: "",
            hourlyWage: s.wageStr, numberOfRooms: 1, delete: "", comment: `${s.dept}外来`
          });
        }
      });
    } else {
      dayShifts.forEach(s => {
        rowsByDate[targetDate].push({
          departmentCode: s.departmentCode || s.clinicId, email: "doctor@example.com", date: targetDate,
          startTime: s.start, splitDateTime1: "", splitDateTime2: "", splitDateTime3: "",
          endTime: s.end, splitHourlyWage1: "", splitHourlyWage2: "", splitHourlyWage3: "",
          hourlyWage: s.wageStr, numberOfRooms: 1, delete: "", comment: `${s.dept}外来`
        });
      });
    }
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