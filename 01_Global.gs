/**
 * ====================================================================
 * 全体共通の変数・設定
 * ====================================================================
 */
const DAYS_OF_WEEK_JP = ['日', '月', '火', '水', '木', '金', '土'];

const WAGE_CONFIG = {
  MASTER_URL: 'https://docs.google.com/spreadsheets/d/1eqejNaKWSuHVnRwxaGT-RHgOnlsHkcXYQ5J32B8T_XM/edit',
  SHEET_2025: '2025年度（年間）',
  SHEET_2026: '2026年度（下期）'
};

// 時給データキャッシュ用
let _wageMasterMap2025 = null;
let _wageData2026 = null;
let _locationMasterMap = null;