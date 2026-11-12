/**
 * @fileoverview 過去データ移行モジュール（全月タブ合体・返書ステータス4段階自動判定・BOM付UTF-8 CSV出力・kintone一括POST＆台帳反映）
 * サンテ溝上病院 地域連携・紹介返書管理＆第一報自動化システム (SANTE-BRIDGE)
 */

const PAST_MIGRATION_HEADERS = Object.freeze([
  '受診日',
  '患者ID',
  '患者氏名',
  '連携方向',
  '連携フェーズ',
  '受入・送出形態',
  '更新用キー',
  '連携先施設名_0',
  '連携先診療科',
  '紹介先診療科',
  '連携先医師名',
  '当院担当診療科',
  '担当医師',
  '返書日',
  '返書ステータス',
  '備考・連絡事項',
  '移行済レコードID',
]);

/**
 * 過去データ走査対象から除外すべきシート名かを判定する
 * @param {string} sheetName
 * @return {boolean}
 */
function isExcludedFromPastMigration(sheetName) {
  const name = String(sheetName || '').trim();
  if (!name) return true;

  const migrationSheetName =
    (SANTE_CONFIG.SHEETS && SANTE_CONFIG.SHEETS.PAST_MIGRATION) ||
    '【kintone一括登録用】紹介・逆紹介';

  if (
    name === migrationSheetName ||
    name === SANTE_CONFIG.SHEETS.INPUT ||
    name === SANTE_CONFIG.SHEETS.FACILITY_MASTER ||
    name.endsWith(SANTE_CONFIG.SHEETS.LEDGER_SUFFIX)
  ) {
    return true;
  }

  if (/^(集計|設定|マスタ|テンプレート|readme|log)/i.test(name)) {
    return true;
  }

  return false;
}

/**
 * 過去データの1行から連携方向・連携フェーズおよび返書ステータス（4段階）を自動判定する純粋関数
 *
 * 優先順:
 * 1. row[12]（問い合わせ）に「逆紹介」または「送出」を含む → 後方（逆紹介・送出） / 対象外
 * 2. 上記以外で、返書日が非空 または row[10] === '済' → 前方（紹介・受入） / 作成・送付完了
 * 3. 上記以外で、row[12] または row[13] が /(書かない|返書なし|不要|記載不要|中止|キャンセル|死亡|終診|閉院)/ にマッチ → 前方（紹介・受入） / 返書不要（医師判断）
 * 4. 上記以外 → 前方（紹介・受入） / 未作成
 *
 * @param {string|Date} rawReplyDate row[9] 返書日
 * @param {string} rawDoneFlag row[10] 済/未
 * @param {string} rawInquiry row[12] 問い合わせ
 * @param {string} rawRemarks row[13] 備考
 * @return {{direction: string, phase: string, replyStatus: string, replyDate: string}}
 */
function classifyPastRecordReplyStatus(rawReplyDate, rawDoneFlag, rawInquiry, rawRemarks) {
  const replyDate = formatDateYMD(rawReplyDate);
  const doneFlag = String(rawDoneFlag || '').trim();
  const inquiry = String(rawInquiry || '').trim();
  const remarks = String(rawRemarks || '').trim();

  // 1. 逆紹介・送出 → 後方 / 対象外
  if (inquiry.includes('逆紹介') || inquiry.includes('送出')) {
    return {
      direction: SANTE_CONFIG.PHASES.BACKWARD,
      phase: SANTE_CONFIG.PHASES.BACKWARD,
      replyStatus: SANTE_CONFIG.REPLY_STATUS.NOT_APPLICABLE,
      replyDate: replyDate,
    };
  }

  const forward = SANTE_CONFIG.PHASES.FORWARD;

  // 2. 返書日が非空 または 済/未が「済」 → 作成・送付完了
  if (replyDate !== '' || doneFlag === '済') {
    return {
      direction: forward,
      phase: forward,
      replyStatus: SANTE_CONFIG.REPLY_STATUS.COMPLETED,
      replyDate: replyDate,
    };
  }

  // 3. 問い合わせ・備考が不要系キーワードにマッチ → 返書不要（医師判断）
  const combinedNotes = `${inquiry} ${remarks}`;
  if (/(書かない|返書なし|不要|記載不要|中止|キャンセル|死亡|終診|閉院)/.test(combinedNotes)) {
    return {
      direction: forward,
      phase: forward,
      replyStatus: SANTE_CONFIG.REPLY_STATUS.NOT_REQUIRED,
      replyDate: '',
    };
  }

  // 4. 上記以外 → 未作成
  return {
    direction: forward,
    phase: forward,
    replyStatus: SANTE_CONFIG.REPLY_STATUS.UNCREATED,
    replyDate: '',
  };
}

/**
 * 過去データのM列（row[12] 問い合わせ）に除外キーワード（報告 → 照会 → 教示）が含まれるか判定する
 * ※ 先に当たった1理由のみを返し、同じ行を二重カウントしない。「逆紹介」は「報告」を含まないため除外されない。
 *
 * @param {string} rawInquiry
 * @return {'報告'|'照会'|'教示'|''} スキップ理由（除外対象外の場合は ''）
 */
function detectPastRowSkipReason(rawInquiry) {
  const inquiry = String(rawInquiry || '').trim();
  if (!inquiry) return '';
  if (inquiry.includes('報告')) return '報告';
  if (inquiry.includes('照会')) return '照会';
  if (inquiry.includes('教示')) return '教示';
  return '';
}

/**
 * 移行シート（【kintone一括登録用】紹介・逆紹介）の17列目に「移行済レコードID」ヘッダが存在することを保証する
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {number} 移行済レコードID列の1-based列番号（17）
 */
function ensureMigratedRecordIdColumn(sheet) {
  const migratedIdColNum = PAST_MIGRATION_HEADERS.length; // 17
  if (!sheet) return migratedIdColNum;
  const migratedIdHeader = PAST_MIGRATION_HEADERS[migratedIdColNum - 1];
  const currentHeader = String(sheet.getRange(1, migratedIdColNum).getValue() || '').trim();
  if (currentHeader !== migratedIdHeader) {
    sheet
      .getRange(1, migratedIdColNum)
      .setValue(migratedIdHeader)
      .setBackground('#0F4C5C')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }
  return migratedIdColNum;
}

/**
 * 過去月別タブの1行（0始まり row[0]〜row[13]）を移行シート用の17列データへ変換する純粋関数
 * @param {Array<*>} row
 * @param {Array<Object>} facilityMasterList
 * @return {{
 *   rowArray: Array<string>,
 *   replyStatus: string,
 *   phase: string
 * }|null}
 */
function transformPastSheetRow(row, facilityMasterList) {
  if (!row || !Array.isArray(row)) return null;

  const rawVisitDate = row[0];
  const visitDate = formatDateYMD(rawVisitDate);
  if (!visitDate || visitDate === '来院日' || !/^\d{4}-\d{2}-\d{2}$/.test(visitDate)) {
    return null;
  }

  const patientId = String(row[1] !== undefined && row[1] !== null ? row[1] : '').trim();
  const patientName = String(row[2] !== undefined && row[2] !== null ? row[2] : '').trim();
  if (!patientId && !patientName) {
    return null;
  }

  // row[12]: 問い合わせに「報告」「照会」「教示」が含まれる行は classifyPastRecordReplyStatus より前に除外
  const rawInquiry = String(row[12] || '').trim();
  if (detectPastRowSkipReason(rawInquiry)) {
    return null;
  }

  // row[3]: 紹介元病院名 → resolveFacility で名寄せ（完全一致／コア一致のみ更新用キーを採用）
  const rawFacilityName = String(row[3] || '').trim();
  const resolvedFacility = resolveFacility(rawFacilityName, facilityMasterList || []);
  const isConfirmedFacility =
    (resolvedFacility.matchType === 'exact' || resolvedFacility.matchType === 'core') &&
    Boolean(resolvedFacility.lookupKey);

  const lookupKey = isConfirmedFacility ? resolvedFacility.lookupKey : '';
  const facilityName0 = isConfirmedFacility
    ? resolvedFacility.officialName || rawFacilityName
    : rawFacilityName;

  // row[4]: 紹介元診療科 → 連携先診療科（自由記述）＋ 紹介先診療科（5択合致時のみ）
  const facilityDept = String(row[4] || '').trim();
  const referralDept = normalizeReferralDeptOption(facilityDept);

  // row[5]: 紹介元医師名（空白除去）
  const facilityDoctor = String(row[5] || '')
    .replace(/[\s\u3000]+/g, '')
    .trim();

  // row[6]: 受入・送出形態（外来 / 入院 / 検査）
  const visitType = normalizeVisitTypeValue(row[6]);

  // row[7], row[8]: 当院診療科・医師（現役11名・9診療科正規化）
  const rawRemarks = String(row[13] || '').trim();
  const ourNormalized = normalizeOurDeptAndDoctor(row[7], row[8], rawInquiry, rawRemarks);

  // row[9]: 返書日, row[10]: 済/未, row[12]: 問い合わせ, row[13]: 備考 → 4段階自動判定
  const statusInfo = classifyPastRecordReplyStatus(row[9], row[10], rawInquiry, rawRemarks);

  // 備考・連絡事項の組み立て（問い合わせ・備考・部分一致施設名退避・退職医/他科退避）
  const remarkParts = [];
  if (rawInquiry) remarkParts.push(rawInquiry);
  if (rawRemarks) remarkParts.push(rawRemarks);
  if (rawFacilityName && !isConfirmedFacility) {
    remarkParts.push(`施設名(マスタ未確定):${rawFacilityName}`);
  }
  if (ourNormalized.extraNote) {
    remarkParts.push(ourNormalized.extraNote);
  }
  const combinedRemarks = remarkParts.join(' / ');

  const rowArray = [
    visitDate,              // 0: 受診日
    patientId,              // 1: 患者ID
    patientName,            // 2: 患者氏名
    statusInfo.direction,   // 3: 連携方向
    statusInfo.phase,       // 4: 連携フェーズ
    visitType,              // 5: 受入・送出形態
    lookupKey,              // 6: 更新用キー
    facilityName0,          // 7: 連携先施設名_0
    facilityDept,           // 8: 連携先診療科
    referralDept,           // 9: 紹介先診療科
    facilityDoctor,         // 10: 連携先医師名
    ourNormalized.dept,     // 11: 当院担当診療科
    ourNormalized.doctor,   // 12: 担当医師
    statusInfo.replyDate,   // 13: 返書日
    statusInfo.replyStatus, // 14: 返書ステータス
    combinedRemarks,        // 15: 備考・連絡事項
    '',                     // 16: 移行済レコードID
  ];

  return {
    rowArray: rowArray,
    replyStatus: statusInfo.replyStatus,
    phase: statusInfo.phase,
  };
}

/**
 * 【タスク4-1】スプレッドシート内の過去月別タブ（「来院日」ヘッダーを持つシート）を走査・合体し、
 * 返書ステータス4段階自動判定と11名医師正規化を行った上で
 * 『【kintone一括登録用】紹介・逆紹介』シートを生成する
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ss
 * @return {{
 *   totalCount: number,
 *   countsByStatus: Object<string, number>,
 *   skippedByReason: Object<string, number>,
 *   skippedCount: number,
 *   processedSheets: Array<string>
 * }}
 */
function buildPastDataMigrationSheet(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  const facilityMasterList = loadFacilityMasterList(spreadsheet);
  const allSheets = spreadsheet.getSheets ? spreadsheet.getSheets() : [];

  const outputRows = [];
  const processedSheets = [];
  const countsByStatus = {
    [SANTE_CONFIG.REPLY_STATUS.COMPLETED]: 0,
    [SANTE_CONFIG.REPLY_STATUS.NOT_REQUIRED]: 0,
    [SANTE_CONFIG.REPLY_STATUS.UNCREATED]: 0,
    [SANTE_CONFIG.REPLY_STATUS.NOT_APPLICABLE]: 0,
  };
  const skippedByReason = {
    報告: 0,
    照会: 0,
    教示: 0,
  };

  allSheets.forEach((sheet) => {
    const sheetName = sheet.getName();
    if (isExcludedFromPastMigration(sheetName)) {
      return;
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 4) {
      return;
    }

    const numCols = Math.max(lastCol, 14);
    const values = sheet.getRange(1, 1, lastRow, numCols).getValues();

    // 先頭5行以内に「来院日」ヘッダーがあるか探す
    let headerRowIdx = -1;
    const maxHeaderScan = Math.min(values.length, 5);
    for (let r = 0; r < maxHeaderScan; r++) {
      const firstCells = values[r].slice(0, 4).map((v) => String(v || '').trim());
      if (firstCells.includes('来院日')) {
        headerRowIdx = r;
        break;
      }
    }

    if (headerRowIdx === -1) {
      return;
    }

    processedSheets.push(sheetName);

    for (let r = headerRowIdx + 1; r < values.length; r++) {
      const row = values[r];
      if (!row || !Array.isArray(row)) continue;
      const visitDate = formatDateYMD(row[0]);
      if (!visitDate || visitDate === '来院日' || !/^\d{4}-\d{2}-\d{2}$/.test(visitDate)) {
        continue;
      }
      const patientId = String(row[1] !== undefined && row[1] !== null ? row[1] : '').trim();
      const patientName = String(row[2] !== undefined && row[2] !== null ? row[2] : '').trim();
      if (!patientId && !patientName) {
        continue;
      }

      // M列（row[12] 問い合わせ）の除外判定（報告 → 照会 → 教示 の順で先に当たった1理由のみカウント）
      const skipReason = detectPastRowSkipReason(row[12]);
      if (skipReason) {
        skippedByReason[skipReason]++;
        continue;
      }

      const transformed = transformPastSheetRow(row, facilityMasterList);
      if (!transformed) continue;

      outputRows.push(transformed.rowArray);
      if (countsByStatus[transformed.replyStatus] !== undefined) {
        countsByStatus[transformed.replyStatus]++;
      } else {
        countsByStatus[transformed.replyStatus] = 1;
      }
    }
  });

  const migrationSheetName =
    (SANTE_CONFIG.SHEETS && SANTE_CONFIG.SHEETS.PAST_MIGRATION) ||
    '【kintone一括登録用】紹介・逆紹介';

  let outSheet = spreadsheet.getSheetByName(migrationSheetName);
  if (!outSheet) {
    outSheet = spreadsheet.insertSheet(migrationSheetName);
  } else if (outSheet.clear) {
    outSheet.clear();
  }

  const headerRange = outSheet.getRange(1, 1, 1, PAST_MIGRATION_HEADERS.length);
  headerRange
    .setValues([Array.from(PAST_MIGRATION_HEADERS)])
    .setBackground('#0F4C5C')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  outSheet.setFrozenRows(1);

  if (outputRows.length > 0) {
    // 患者ID列（2列目）と受診日列（1列目）・返書日列（14列目）・移行済レコードID列（17列目）を文字列書式に維持
    outSheet.getRange(2, 1, outputRows.length, 2).setNumberFormat('@');
    outSheet.getRange(2, 14, outputRows.length, 1).setNumberFormat('@');
    outSheet.getRange(2, PAST_MIGRATION_HEADERS.length, outputRows.length, 1).setNumberFormat('@');
    outSheet
      .getRange(2, 1, outputRows.length, PAST_MIGRATION_HEADERS.length)
      .setValues(outputRows);
  }

  const summaryText = [
    `対象シート数: ${processedSheets.length}タブ`,
    `抽出合計: ${outputRows.length}件`,
    `・作成・送付完了: ${countsByStatus[SANTE_CONFIG.REPLY_STATUS.COMPLETED]}件`,
    `・返書不要（医師判断）: ${countsByStatus[SANTE_CONFIG.REPLY_STATUS.NOT_REQUIRED]}件`,
    `・未作成: ${countsByStatus[SANTE_CONFIG.REPLY_STATUS.UNCREATED]}件`,
    `・対象外（逆紹介等）: ${countsByStatus[SANTE_CONFIG.REPLY_STATUS.NOT_APPLICABLE]}件`,
    `・スキップ（報告）: ${skippedByReason['報告']}件`,
    `・スキップ（照会）: ${skippedByReason['照会']}件`,
    `・スキップ（教示）: ${skippedByReason['教示']}件`,
  ].join('\n');

  try {
    if (SpreadsheetApp.getUi) {
      SpreadsheetApp.getUi().alert(
        '✅ 過去データ移行シート作成完了',
        `シート「${migrationSheetName}」を作成しました。\n\n${summaryText}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      spreadsheet.toast(summaryText.replace(/\n/g, ' / '), '✅ 過去データ整形完了', 8);
    }
  } catch (_) {
    spreadsheet.toast(summaryText.replace(/\n/g, ' / '), '✅ 過去データ整形完了', 8);
  }

  return {
    totalCount: outputRows.length,
    countsByStatus: countsByStatus,
    skippedByReason: skippedByReason,
    skippedCount:
      skippedByReason['報告'] + skippedByReason['照会'] + skippedByReason['教示'],
    processedSheets: processedSheets,
  };
}

/**
 * 2次元配列から BOM 付き UTF-8（先頭 \uFEFF）の CSV 文字列を生成する純粋関数
 * @param {Array<Array<*>>} matrix
 * @return {string}
 */
function buildMigrationCsvContent(matrix) {
  if (!matrix || matrix.length === 0) {
    return '\uFEFF';
  }

  const escapeCsvCell = (val) => {
    if (val === null || val === undefined) return '';
    const str = val instanceof Date ? formatDateYMD(val) : String(val);
    if (/[",\r\n]/.test(str)) {
      return `"${str.replace(/"/g, '""')}"`;
    }
    return str;
  };

  const lines = matrix.map((row) => row.map(escapeCsvCell).join(','));
  return '\uFEFF' + lines.join('\r\n');
}

/**
 * 【タスク4-2】『【kintone一括登録用】紹介・逆紹介』の内容を
 * BOM 付き UTF-8 の CSV として HtmlService モーダル経由でダウンロードする
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ss
 * @return {{csvContent: string, rowCount: number}}
 */
function downloadMigrationCsv(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  const migrationSheetName =
    (SANTE_CONFIG.SHEETS && SANTE_CONFIG.SHEETS.PAST_MIGRATION) ||
    '【kintone一括登録用】紹介・逆紹介';
  const sheet = spreadsheet.getSheetByName(migrationSheetName);

  if (!sheet || sheet.getLastRow() < 2) {
    throw new Error(
      `シート「${migrationSheetName}」にデータがありません。先に「1️⃣ 全月タブ合体＆返書ステータス自動判定シート作成」を実行してください。`
    );
  }

  ensureMigratedRecordIdColumn(sheet);

  const lastRow = sheet.getLastRow();
  const matrix = sheet
    .getRange(1, 1, lastRow, PAST_MIGRATION_HEADERS.length)
    .getValues();
  const csvContent = buildMigrationCsvContent(matrix);
  const base64Csv = Utilities.base64Encode(csvContent, Utilities.Charset.UTF_8);
  const fileName = `sante_past_migration_${formatDateYMD(new Date())}.csv`;

  const htmlString = [
    '<!DOCTYPE html>',
    '<html><head><meta charset="utf-8"></head>',
    '<body style="font-family: sans-serif; padding: 16px; text-align: center;">',
    `<p>全 <b>${lastRow - 1}件</b> の移行用CSV（BOM付UTF-8）を準備しました。</p>`,
    `<a id="dlLink" href="data:text/csv;charset=utf-8;base64,${base64Csv}" download="${fileName}" `,
    'style="display:inline-block;padding:10px 20px;background:#0F4C5C;color:#fff;text-decoration:none;border-radius:4px;font-weight:bold;">',
    '📥 CSVファイルをダウンロード</a>',
    '<script>document.getElementById("dlLink").click();</script>',
    '</body></html>',
  ].join('');

  if (typeof HtmlService !== 'undefined' && SpreadsheetApp.getUi) {
    const htmlOutput = HtmlService.createHtmlOutput(htmlString)
      .setWidth(420)
      .setHeight(180);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📥 移行用CSVダウンロード（BOM付UTF-8）');
  }

  return {
    csvContent: csvContent,
    rowCount: lastRow - 1,
  };
}

/**
 * 【タスク4-3】『【kintone一括登録用】紹介・逆紹介』のデータを読み取り、
 * ルックアップ自動コピー先の「連携先施設名_0」や空の日付・空のドロップダウン項目を除外した上で、
 * 100件単位で kintone アプリBへ一括 POST し、発番された ids を「移行済レコードID」列へ記録するとともに年度別連携台帳へも同時反映する
 * （既に「移行済レコードID」列に値がある行はスキップし、二重登録を防止する）
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ss
 * @return {{postedCount: number, ledgerAppendedCount: number, skippedCount: number, ids: Array<string>}}
 */
function sendPastDataToKintoneAppB(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  const migrationSheetName =
    (SANTE_CONFIG.SHEETS && SANTE_CONFIG.SHEETS.PAST_MIGRATION) ||
    '【kintone一括登録用】紹介・逆紹介';
  const sheet = spreadsheet.getSheetByName(migrationSheetName);

  if (!sheet || sheet.getLastRow() < 2) {
    throw new Error(
      `シート「${migrationSheetName}」にデータがありません。先に「1️⃣ 全月タブ合体＆返書ステータス自動判定シート作成」を実行してください。`
    );
  }

  const migratedIdColNum = ensureMigratedRecordIdColumn(sheet);
  const migratedIdColIdx = migratedIdColNum - 1; // 16

  const lastRow = sheet.getLastRow();
  const rows = sheet
    .getRange(2, 1, lastRow - 1, PAST_MIGRATION_HEADERS.length)
    .getValues();

  const nowStr = formatDateTimeNow();
  const migrationRecords = [];
  let skippedCount = 0;

  rows.forEach((row, rowIdx) => {
    const sheetRowNumber = rowIdx + 2;
    const visitDate = formatDateYMD(row[0]);
    const patientId = String(row[1] || '').trim();
    const patientName = String(row[2] || '').trim();
    if (!visitDate && !patientId && !patientName) return;

    // 既に「移行済レコードID」が記録されている行は POST・台帳追記ともにスキップ
    const existingMigratedId = String(row[migratedIdColIdx] || '').trim();
    if (existingMigratedId !== '') {
      skippedCount++;
      return;
    }

    const phaseVal = String(row[4] || row[3] || SANTE_CONFIG.PHASES.FORWARD).trim();
    const visitTypeVal = String(row[5] || '').trim();
    const lookupKeyVal = String(row[6] || '').trim();
    const displayFacilityName = String(row[7] || '').trim(); // 連携先施設名_0（台帳記録用）
    const facilityDeptVal = String(row[8] || '').trim();
    const referralDeptVal = String(row[9] || '').trim();
    const facilityDoctorVal = String(row[10] || '').trim();
    const ourDeptVal = String(row[11] || '').trim();
    const ourDoctorVal = String(row[12] || '').trim();
    const replyDateVal = formatDateYMD(row[13]);
    const replyStatusVal = String(row[14] || '').trim();
    const remarksVal = String(row[15] || '').trim();

    migrationRecords.push({
      _sheetRowNumber: sheetRowNumber,
      registeredAt: nowStr,
      visitDate: visitDate,
      patientId: patientId,
      patientName: patientName,
      phase: phaseVal,
      visitType: visitTypeVal,
      lookupKey: lookupKeyVal,
      // 連携先施設名_0 はルックアップコピー先のため kintone POST ペイロードからは除外（空文字）
      facilityName: '',
      _ledgerFacilityName: displayFacilityName,
      facilityDept: facilityDeptVal,
      referralDept: referralDeptVal,
      facilityDoctor: facilityDoctorVal,
      ourDept: ourDeptVal,
      ourDoctor: ourDoctorVal,
      replyStatus: replyStatusVal,
      replyDate: replyDateVal,
      firstReportUrl: '',
      faxStatus: '', // 空のドロップダウン項目は送信しない
      remarks: remarksVal,
    });
  });

  if (migrationRecords.length === 0) {
    const skipMsg = `新規移行対象は 0件 です（移行済スキップ: ${skippedCount}件）。`;
    spreadsheet.toast(skipMsg, 'ℹ️ 過去データ一括移行', 8);
    return {
      postedCount: 0,
      ledgerAppendedCount: 0,
      skippedCount: skippedCount,
      ids: [],
    };
  }

  let postedCount = 0;
  let ledgerAppendedCount = 0;
  const allIssuedIds = [];

  try {
    // 100件チャンク成功ごとに移行シートの「移行済レコードID」列への書き込みと年度別台帳追記を確定する
    postRecordsToKintoneBatch(migrationRecords, {
      onChunkSuccess: (chunk, chunkIds) => {
        chunk.forEach((rec, idx) => {
          const recId = chunkIds[idx] || '';
          sheet.getRange(rec._sheetRowNumber, migratedIdColNum).setValue(recId);
        });

        const chunkLedgerRecords = chunk.map((rec, idx) => ({
          kintoneRecordId: chunkIds[idx] || '',
          registeredAt: rec.registeredAt,
          visitDate: rec.visitDate,
          patientId: rec.patientId,
          patientName: rec.patientName,
          phase: rec.phase,
          visitType: rec.visitType,
          lookupKey: rec.lookupKey,
          facilityName: rec._ledgerFacilityName,
          facilityDept: rec.facilityDept,
          facilityDoctor: rec.facilityDoctor,
          ourDept: rec.ourDept,
          ourDoctor: rec.ourDoctor,
          replyStatus: rec.replyStatus,
          replyDate: rec.replyDate,
          firstReportUrl: '',
          faxStatus: SANTE_CONFIG.FAX_STATUS.NOT_REQUIRED,
          remarks: rec.remarks,
        }));

        const appendedMeta = appendRecordsToLedger(chunkLedgerRecords, spreadsheet);
        postedCount += chunk.length;
        ledgerAppendedCount += appendedMeta.length;
        chunkIds.forEach((id) => allIssuedIds.push(id));
      },
    });

    const doneMsg = `過去データ ${postedCount}件 を kintone アプリBへ一括登録し、年度別連携台帳へ ${ledgerAppendedCount}件 反映しました（移行済スキップ: ${skippedCount}件）。`;
    spreadsheet.toast(doneMsg, '🎉 過去データ一括移行完了', 8);

    return {
      postedCount: postedCount,
      ledgerAppendedCount: ledgerAppendedCount,
      skippedCount: skippedCount,
      ids: allIssuedIds,
    };
  } catch (err) {
    notifySlackError(
      '過去データkintone一括移行 (sendPastDataToKintoneAppB) 異常終了',
      err,
      {
        targetRecordCount: String(migrationRecords.length),
        completedBeforeErrorCount: String(postedCount),
        skippedCount: String(skippedCount),
      },
      { includeStack: false }
    );
    spreadsheet.toast(`過去データ移行エラー: ${err.message}`, '🚨 移行中断', 10);
    throw err;
  }
}
