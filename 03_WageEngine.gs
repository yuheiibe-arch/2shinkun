/**
 * ====================================================================
 * 【基幹エンジン】時給データ取得・検索モジュール
 * ====================================================================
 */

// ★追加：IDマスタキャッシュ用
let _idMasterCache = null;

function initializeWageData() {
  if (_wageMasterMap2025 && _wageData2026 && _idMasterCache) return;

  const masterSs = SpreadsheetApp.openByUrl(WAGE_CONFIG.MASTER_URL);
  
  _locationMasterMap = typeof fetchMasterData === 'function' ? fetchMasterData().map : {}; 

  // ★追加：拠点名マスタ（departmentCode, clinicId）のキャッシュ
  const nameSheet = masterSs.getSheetByName('拠点名');
  _idMasterCache = {};
  if (nameSheet) {
    const data = nameSheet.getDataRange().getValues();
    const headers = data[0];
    const officialCol = headers.indexOf('正規記載');
    const idCol = headers.indexOf('クリニックNo');
    const deptCodeCol = headers.indexOf('departmentCode');
    const yureCols = [];
    headers.forEach((h, idx) => { if (String(h).includes('表記揺れ')) yureCols.push(idx); });

    for (let i = 1; i < data.length; i++) {
      const officialName = data[i][officialCol];
      const clinicId = data[i][idCol];
      const deptCode = (deptCodeCol !== -1 && data[i][deptCodeCol]) ? data[i][deptCodeCol] : "";
      
      if (!officialName || !clinicId) continue;
      const mapData = { clinicId: clinicId, departmentCode: deptCode };
      
      _idMasterCache[_normalizeForSearch(officialName)] = mapData;
      yureCols.forEach(c => {
        const yureName = data[i][c];
        if (yureName) _idMasterCache[_normalizeForSearch(yureName)] = mapData;
      });
    }
  }

  // 2025年度データのキャッシュ
  const sheet2025 = masterSs.getSheetByName(WAGE_CONFIG.SHEET_2025);
  _wageMasterMap2025 = {};
  if (sheet2025) {
    const values = sheet2025.getRange('B2:O' + Math.max(2, sheet2025.getLastRow())).getValues();
    values.forEach(row => {
      const rawClinic = (row[0] || '').toString().trim();
      const clinic = _normalizeForSearch(rawClinic);
      const dept = (row[1] || '').toString().replace(/\s+/g, '');
      
      if (clinic) {
        _wageMasterMap2025[`${clinic}||${dept}`] = {
          wd_am: row[3], wd_pm: row[4], wd_nt: row[5],
          hol_am: row[11], hol_pm: row[12], hol_nt: row[13]
        };
      }
    });
  }

  // 2026年度データのキャッシュ
  const sheet2026 = masterSs.getSheetByName(WAGE_CONFIG.SHEET_2026);
  _wageData2026 = sheet2026 ? sheet2026.getDataRange().getValues() : [];
}

/**
 * 文字列の正規化・表記ブレ補正
 */
function _normalizeForSearch(rawName) {
  if (!rawName) return "";
  let name = rawName.toString();
  if (name.normalize) name = name.normalize('NFKC');
  
  name = name
    .replace(/[【】\[\]]/g, "") 
    .replace(/[\(（]?(内科|小児科|皮膚科|整形外科)[\)）]?/g, "") 
    .replace(/(病院|クリニック|診療所|モール)$/g, "") 
    .replace(/\s+/g, "") 
    .replace(/ヶ/g, "ケ") 
    .trim();

  if (name === "千葉NT") name = "千葉ニュータウン中央";

  return name;
}

// ★追加：IDデータ取得用関数
function getClinicIdData(rawClinicName) {
  if (!_idMasterCache) initializeWageData();
  const cleanLoc = _normalizeForSearch(rawClinicName);
  return _idMasterCache[cleanLoc] || { clinicId: "", departmentCode: "" };
}

function _findRate2025(formalName, deptName) {
  const exact = _wageMasterMap2025[`${formalName}||${deptName}`] || 
                _wageMasterMap2025[`${formalName}||共通`] || 
                _wageMasterMap2025[`${formalName}||`];
  if (exact) return exact;

  for (const key in _wageMasterMap2025) {
    const [c, d] = key.split("||");
    if (c.includes(formalName) || formalName.includes(c)) {
      if (deptName === "小児科" && (d.includes("小児") || c.includes("小児"))) return _wageMasterMap2025[key];
      if (deptName === "内科" && (!d.includes("小児") && !c.includes("小児"))) return _wageMasterMap2025[key];
    }
  }
  return null;
}

function _findRate2026(formalName, deptName) {
  let found = _wageData2026.find(row => {
    const c = _normalizeForSearch(row[1]);
    const d = (row[2] || "").toString();
    if (!c) return false;
    if (!c.includes(formalName) && !formalName.includes(c)) return false;

    if (deptName === "小児科") return c.includes("小児") || d.includes("小児");
    if (deptName === "内科") return (c.includes("内科") || d.includes("内科")) || (!c.includes("小児") && !d.includes("小児"));
    return false;
  });

  if (!found) {
    found = _wageData2026.find(row => {
      const c = _normalizeForSearch(row[1]);
      return c && c.includes(formalName);
    });
  }

  if (found) {
    return {
      wd_am: found[3], wd_pm: found[4], wd_nt: found[5],
      hol_am: found[6], hol_pm: found[7], hol_nt: found[8]
    };
  }
  return null;
}

function getClinicWages(rawClinicName) {
  if (!_wageMasterMap2025) initializeWageData();

  const cleanLoc = _normalizeForSearch(rawClinicName);
  const formalName = _locationMasterMap[cleanLoc] || cleanLoc;

  const targetDepts = (formalName.includes("北葛西") || formalName.includes("亀有")) 
                      ? ["内科", "小児科"] 
                      : ["小児科"];

  const results = [];
  targetDepts.forEach(dept => {
    const rate2025 = _findRate2025(formalName, dept);
    const rate2026 = _findRate2026(formalName, dept);

    results.push({
      clinicName: formalName,
      department: dept,
      rates: {
        y2025: rate2025 || null,
        y2026: rate2026 || null
      }
    });
  });

  return results;
}