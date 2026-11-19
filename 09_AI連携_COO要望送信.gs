/**
 * ============================================================================
 * 【2shinkun 追加用】09_AI連携_COO要望送信.gs
 * - 『２診要望一覧』シートのデータを暗号化して capsdoctor365@caps365.jp へ送信
 * - 作業リスト作成時の自動送信 ＆ メニューからの手動一括送信に対応
 * ============================================================================
 */

function get2shinkunBridgeConfig_() {
  return {
    RECIPIENT_EMAIL: 'capsdoctor365@caps365.jp',
    SUBJECT_COO_SYNC: '【AI連携暗号化JSON】COO2診要望リストパック',
    ATTACHMENT_FILENAME: 'coo_2nd_requests_pack.enc',
    DEFAULT_SECRET_KEY: 'CAPS_MNYS_AI_AGENT_SECRET_2026_KEY_#9876'
  };
}

/**
 * 🚀 メニュー・エディタから手動で『２診要望一覧』全件をcaps側へ即時送信する関数
 */
function syncAllCooRequestsToCapsManual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const count = sendEncryptedCooRequestsToCaps_(ss);
  SpreadsheetApp.getUi().alert(`『２診要望一覧』の要望データ（${count}枠）を暗号化してcaps側へ送信しました。`);
}

/**
 * 要望表作成.gs の末尾から自動で呼び出される暗号化送信本体
 */
function sendEncryptedCooRequestsToCaps_(ss) {
  const targetSheet = (ss || SpreadsheetApp.getActiveSpreadsheet()).getSheetByName('２診要望一覧');
  if (!targetSheet || targetSheet.getLastRow() < 2) return 0;

  const vals = targetSheet.getDataRange().getDisplayValues();
  const records = [];
  const targetMonthsSet = new Set();

  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    const locRaw = String(row[0] || '').trim();
    const dateRaw = String(row[1] || '').trim();
    const startStr = String(row[2] || '').trim();
    const endStr = String(row[3] || '').trim();
    const hoursRaw = String(row[4] || '').trim().toLowerCase();
    const clinicId = String(row[6] || '').trim();

    // サマリー行（A列・B列が空欄の「月間要望拠点数」等）はスキップ
    if (!locRaw || !dateRaw || !startStr || !endStr) continue;

    const mDate = dateRaw.replace(/[年月.\-]/g, '/').match(/(\d{4})\/(\d{1,2})\/(\d{1,2})/);
    if (!mDate) continue;
    const normDate = `${mDate[1]}/${('0' + mDate[2]).slice(-2)}/${('0' + mDate[3]).slice(-2)}`;
    const ym = normDate.slice(0, 7);

    const mNum = hoursRaw.match(/([0-9]+(?:\.[0-9]+)?)/);
    if (!mNum) continue;
    const hours = parseFloat(mNum[1]);
    if (hours <= 0) continue;

    const isF = hoursRaw.includes('f');
    const dept = locRaw.includes('内科') ? '内科' : '小児科';
    let cleanLoc = locRaw.replace(/[（(].*?[）)]/g, '').trim();
    if (cleanLoc === '千葉NT') cleanLoc = '千葉ニュータウン中央';

    targetMonthsSet.add(ym);
    records.push({
      ym: ym,
      d: normDate,
      c: cleanLoc,
      dp: dept,
      st: startStr.length === 4 ? '0' + startStr : startStr,
      et: endStr.length === 4 ? '0' + endStr : endStr,
      h: hours,
      f: isF ? 1 : 0,
      cid: clinicId
    });
  }

  if (records.length === 0) return 0;

  const cfg = get2shinkunBridgeConfig_();
  const generatedAt = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  const payload = {
    type: 'COO_2ND_REQUEST_SYNC',
    generatedAt: generatedAt,
    months: Array.from(targetMonthsSet).sort(),
    count: records.length,
    records: records
  };

  const secretKey = PropertiesService.getScriptProperties().getProperty('AI_BRIDGE_SECRET_KEY') || cfg.DEFAULT_SECRET_KEY;
  const encryptedPkg = encrypt2shinkunPayload_(JSON.stringify(payload), secretKey);
  const blob = Utilities.newBlob(JSON.stringify(encryptedPkg), 'application/octet-stream', cfg.ATTACHMENT_FILENAME);

  GmailApp.sendEmail(
    cfg.RECIPIENT_EMAIL,
    `${cfg.SUBJECT_COO_SYNC} (${generatedAt})`,
    `2shinkun『２診要望一覧』自動連携パック（対象月: ${payload.months.join(', ')} / 計 ${records.length}枠）`,
    { attachments: [blob] }
  );

  console.log(`✅ [AI連携] COO2診要望リスト（${records.length}枠）を暗号化して送信しました。`);
  return records.length;
}

function encrypt2shinkunPayload_(plainText, secretKey) {
  const textBlob = Utilities.newBlob(plainText, 'application/json', 'payload.json');
  const gzipBytes = Utilities.gzip(textBlob).getBytes();
  const ivStr = Utilities.getUuid().replace(/-/g, '').slice(0, 16);

  const output = new Array(gzipBytes.length);
  let keyStream = [];
  let blockIndex = 0;
  for (let i = 0; i < gzipBytes.length; i++) {
    if (i % 32 === 0) {
      keyStream = Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        `${secretKey}|${ivStr}|${blockIndex}`,
        Utilities.Charset.UTF_8
      );
      blockIndex++;
    }
    output[i] = (gzipBytes[i] ^ keyStream[i % 32]) & 0xff;
    if (output[i] > 127) output[i] -= 256;
  }

  const cipherBase64 = Utilities.base64Encode(output);
  const sigBytes = Utilities.computeHmacSha256Signature(ivStr + ':' + cipherBase64, secretKey);
  return { v: 1, iv: ivStr, data: cipherBase64, sig: Utilities.base64Encode(sigBytes) };
}