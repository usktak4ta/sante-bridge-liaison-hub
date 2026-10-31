/**
 * @fileoverview SANTE-BRIDGE メインコントローラー・UIトリガー・一括処理オーケストレーター
 * サンテ溝上病院 地域連携・紹介返書管理＆第一報自動化システム
 */

/**
 * スプレッドシート起動時トリガー: カスタムメニューを構築する
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏥 SANTE-BRIDGE 連携メニュー')
    .addItem('🚀 一括実行（第一報生成・kintone連携・台帳反映）', 'executeBatchProcess')
    .addItem('📝 第一報ドキュメントのみ先行生成', 'executeFirstReportGenerationOnly')
    .addItem('🔍 全列の5重ロック再照合', 'reRunFiveLockMatchingForAllColumns')
    .addSeparator()
    .addItem('📢 未返書リストを今すぐ通知（LINE WORKS / Slack）', 'sendUnrepliedRemindersManual')
    .addItem('🧹 送信完了・処理済列のクリア整理', 'clearCompletedColumns')
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu('⚙️ 初期設定・メンテナンス')
        .addItem('📋 3タブ初期構築・フォーマット整備', 'setupAllSheets')
        .addItem('⏰ 未返書リマインド定期トリガー設定（毎週月・木 9:00）', 'setupReminderScheduleTrigger')
        .addItem('📄 第一報Googleドキュメント標準テンプレート生成', 'handleCreateDefaultTemplateMenu')
        .addItem('🔑 スクリプトプロパティ初期化', 'initScriptPropertiesTemplate')
        .addItem('🔔 Slack疎通テスト通知', 'testSlackWebhookConnection')
    )
    .addToUi();
}

/**
 * セル編集時トリガー:
 * `紹介・逆紹介入力` シートの `C2:C3` にGemini抽出JSONがペーストされた際、
 * C列以降へ横展開し、施設マスター名寄せ＆年度別台帳との5重ロック自動照合を実行する
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== SANTE_CONFIG.SHEETS.INPUT) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  const numRows = e.range.getNumRows();
  const numCols = e.range.getNumColumns();

  // C2:C3 (row 2〜3, col 3) との交差判定
  const intersectsPasteRow = row <= 3 && row + numRows - 1 >= 2;
  const intersectsPasteCol = col <= 3 && col + numCols - 1 >= 3;
  if (!intersectsPasteRow || !intersectsPasteCol) return;

  const pasteRange = sheet.getRange(SANTE_CONFIG.INPUT_LAYOUT.PASTE_RANGE);
  const pasteValues = pasteRange.getValues();
  const rawInput = String(pasteValues[0][0] || pasteValues[1][0] || '').trim();

  if (!rawInput) return;
  if (!rawInput.includes('{') && !rawInput.includes('[')) return;

  try {
    expandJsonAndMatchLedger(rawInput, sheet.getParent());
    pasteRange.clearContent();
  } catch (err) {
    Logger.log(`JSON展開エラー: ${err.message}\n${err.stack || ''}`);
    notifySlackError('C2セル JSON貼付・横展開エラー', err);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `JSON解析エラー: ${err.message}`,
      '🚨 SANTE-BRIDGE エラー',
      10
    );
  }
}

/**
 * Gemini (Gem) が出力したJSON文字列を安全にパースし、患者データ配列として返す
 * Markdownのコードブロック (```json ... ```) や単一オブジェクト・ラッパーオブジェクトにも対応
 *
 * @param {string} rawText
 * @return {Array<Object>}
 */
function parseGeminiJsonInput(rawText) {
  let cleaned = String(rawText || '').trim();

  // ```json ... ``` のコードフェンスを除去
  cleaned = cleaned
    .replace(/^```(?:json)?\s*/i, '')
    .replace(/\s*```$/i, '')
    .trim();

  let parsed;
  try {
    parsed = JSON.parse(cleaned);
  } catch (parseErr) {
    throw new Error(`有効なJSON形式ではありません (${parseErr.message})`);
  }

  if (Array.isArray(parsed)) {
    return parsed;
  }

  if (parsed && typeof parsed === 'object') {
    const candidateKeys = ['records', 'patients', 'data', 'items', 'referrals'];
    for (let i = 0; i < candidateKeys.length; i++) {
      const k = candidateKeys[i];
      if (Array.isArray(parsed[k])) {
        return parsed[k];
      }
    }
    return [parsed];
  }

  throw new Error('JSONデータから患者レコードを抽出できませんでした。');
}

/**
 * 連携フェーズ文字列を4つの正規選択肢のいずれかに正規化する
 * @param {string} rawPhase
 * @return {string}
 */
function normalizePhaseValue(rawPhase) {
  const s = String(rawPhase || '').trim();
  const phases = SANTE_CONFIG.PHASES;

  if (s === phases.FORWARD || s === phases.BACKWARD || s === phases.REPLY || s === phases.OTHER) {
    return s;
  }
  if (s.includes('逆紹介') || s.includes('後方') || s.includes('送出')) {
    return phases.BACKWARD;
  }
  if (s.includes('返書') || s.includes('経過') || s.includes('結果報告')) {
    return phases.REPLY;
  }
  if (s.includes('その他') || s.includes('介護') || s.includes('対診')) {
    return phases.OTHER;
  }
  if (s.includes('前方') || s.includes('紹介') || s.includes('受入')) {
    return phases.FORWARD;
  }
  return phases.FORWARD;
}

/**
 * 受入・送出形態を「外来」または「入院」に正規化する
 * @param {string} rawType
 * @return {string}
 */
function normalizeVisitTypeValue(rawType) {
  const s = String(rawType || '').trim();
  if (s.includes('入院')) {
    return SANTE_CONFIG.VISIT_TYPES.INPATIENT;
  }
  return SANTE_CONFIG.VISIT_TYPES.OUTPATIENT;
}

/**
 * JSONオブジェクトの1要素から、キー揺れを吸収して正規化済みの患者オブジェクトを構築する
 * @param {Object} rawItem
 * @return {{
 *   visitDate: string,
 *   patientId: string,
 *   patientName: string,
 *   phase: string,
 *   visitType: string,
 *   facilityName: string,
 *   facilityDept: string,
 *   facilityDoctor: string,
 *   ourDept: string,
 *   ourDoctor: string,
 *   needFirstReport: boolean|null,
 *   remarks: string
 * }}
 */
function normalizeRawJsonRecord(rawItem) {
  const pick = (keys) => {
    for (let i = 0; i < keys.length; i++) {
      const val = rawItem[keys[i]];
      if (val !== undefined && val !== null && String(val).trim() !== '') {
        return String(val).trim();
      }
    }
    return '';
  };

  const visitDateRaw = pick(['受診日', 'visitDate', 'visit_date', 'date', '紹介日', '記入日', '日付']);
  const visitDate = formatDateYMD(visitDateRaw) || formatDateYMD(new Date());
  const patientId = pick(['患者ID', 'patientId', 'patient_id', 'id', '患者番号', 'カルテ番号']);
  const patientName = pick(['患者氏名', 'patientName', 'patient_name', 'name', '氏名']);
  const phase = normalizePhaseValue(
    pick(['連携フェーズ', 'phase', '区分', '種別', '紹介区分', 'フェーズ'])
  );
  const visitType = normalizeVisitTypeValue(
    pick(['受入・送出形態', '受入形態', '送出形態', '形態', 'visitType', 'visit_type', '入院外来区分'])
  );
  const facilityName = pick([
    '連携先施設名',
    '施設名',
    '医療機関名',
    '紹介元医療機関',
    '紹介先医療機関',
    'facilityName',
    'facility_name',
  ]);
  const facilityDept = pick([
    '連携先診療科',
    '相手先診療科',
    '紹介元診療科',
    '紹介先診療科',
    '診療科',
    'facilityDepartment',
    'facility_dept',
  ]);
  const facilityDoctor = pick([
    '連携先医師名',
    '相手先医師名',
    '紹介元医師名',
    '紹介元医師',
    '紹介先医師名',
    '紹介先医師',
    '医師名',
    'facilityDoctor',
    'facility_doctor',
  ]);
  const ourDept = pick([
    '当院担当診療科',
    '当院診療科',
    '担当診療科',
    'ourDepartment',
    'our_dept',
  ]);
  const ourDoctor = pick([
    '担当医師',
    '当院担当医師',
    '当院担当医',
    '当院医師名',
    'ourDoctor',
    'our_doctor',
  ]);
  const remarks = pick(['備考・連絡事項', '備考', '連絡事項', 'remarks', 'notes', 'memo']);

  let needFirstReport = null;
  if (typeof rawItem['📝 第一報作成'] === 'boolean') {
    needFirstReport = rawItem['📝 第一報作成'];
  } else if (typeof rawItem.needFirstReport === 'boolean') {
    needFirstReport = rawItem.needFirstReport;
  }

  return {
    visitDate,
    patientId,
    patientName,
    phase,
    visitType,
    facilityName,
    facilityDept,
    facilityDoctor,
    ourDept,
    ourDoctor,
    needFirstReport,
    remarks,
  };
}

/**
 * 紹介・逆紹介入力シート上で次に書き込み可能な空き開始列番号（C列=3以降）を取得する
 * @param {GoogleAppsScript.Spreadsheet.Sheet} inputSheet
 * @return {number} 1-based 列番号（3以上）
 */
function findNextAvailableColumn(inputSheet) {
  const startCol = SANTE_CONFIG.INPUT_LAYOUT.DATA_START_COL; // 3 (C列)
  const lastCol = inputSheet.getLastColumn();
  if (lastCol < startCol) {
    return startCol;
  }

  const rowStart = SANTE_CONFIG.INPUT_LAYOUT.ROW_START; // 5
  const rowCount = 10; // 5行目〜14行目（基本患者情報領域）で空き列を判定
  const matrix = inputSheet
    .getRange(rowStart, startCol, rowCount, lastCol - startCol + 1)
    .getValues();

  let rightmostUsedOffset = -1;
  for (let colOffset = 0; colOffset <= lastCol - startCol; colOffset++) {
    const hasData = matrix.some((row) => String(row[colOffset] || '').trim() !== '');
    if (hasData) {
      rightmostUsedOffset = colOffset;
    }
  }

  return startCol + rightmostUsedOffset + 1;
}

/**
 * 指定された列範囲（C列以降）に入力規則（プルダウン・チェックボックス）と書式を適用する
 * @param {GoogleAppsScript.Spreadsheet.Sheet} inputSheet
 * @param {number} startCol
 * @param {number} numCols
 */
function applyInputColumnValidations(inputSheet, startCol, numCols) {
  if (numCols <= 0) return;
  const rows = SANTE_CONFIG.INPUT_LAYOUT.ROWS;

  // 5行目（受診日）をテキスト書式（YYYY-MM-DD維持）に設定
  inputSheet.getRange(rows.VISIT_DATE, startCol, 1, numCols).setNumberFormat('@');
  // 6行目（患者ID）をテキスト書式（先頭ゼロ維持）に設定
  inputSheet.getRange(rows.PATIENT_ID, startCol, 1, numCols).setNumberFormat('@');

  // 8行目: 連携フェーズ プルダウン
  const phaseRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      [
        SANTE_CONFIG.PHASES.FORWARD,
        SANTE_CONFIG.PHASES.BACKWARD,
        SANTE_CONFIG.PHASES.REPLY,
        SANTE_CONFIG.PHASES.OTHER,
      ],
      true
    )
    .setAllowInvalid(true)
    .build();
  inputSheet.getRange(rows.PHASE, startCol, 1, numCols).setDataValidation(phaseRule);

  // 9行目: 受入・送出形態 プルダウン
  const visitTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      [SANTE_CONFIG.VISIT_TYPES.OUTPATIENT, SANTE_CONFIG.VISIT_TYPES.INPATIENT],
      true
    )
    .setAllowInvalid(true)
    .build();
  inputSheet.getRange(rows.VISIT_TYPE, startCol, 1, numCols).setDataValidation(visitTypeRule);

  // 15行目: 📝 第一報作成 チェックボックス
  inputSheet.getRange(rows.NEED_FIRST_REPORT, startCol, 1, numCols).insertCheckboxes();

  // 17行目: ✅ FAX送信完了 チェックボックス
  inputSheet.getRange(rows.FAX_SENT, startCol, 1, numCols).insertCheckboxes();

  // 罫線・列幅の整備
  const dataBlock = inputSheet.getRange(
    SANTE_CONFIG.INPUT_LAYOUT.STATUS_ROW,
    startCol,
    SANTE_CONFIG.INPUT_LAYOUT.ROW_END - SANTE_CONFIG.INPUT_LAYOUT.STATUS_ROW + 1,
    numCols
  );
  dataBlock.setBorder(true, true, true, true, true, true, '#d9d9d9', SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * 【要件A】JSON文字列をパースしてC列以降に横展開し、
 * 施設マスター名寄せ＆年度別台帳との5重ロック自動照合を一括実行する
 *
 * @param {string} rawJsonText
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ss
 * @return {{expandedCount: number, matchedCount: number, startCol: number}}
 */
function expandJsonAndMatchLedger(rawJsonText, ss) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = spreadsheet.getSheetByName(SANTE_CONFIG.SHEETS.INPUT);
  if (!inputSheet) {
    throw new Error(`シート「${SANTE_CONFIG.SHEETS.INPUT}」が存在しません。初期設定を実行してください。`);
  }

  const rawRecords = parseGeminiJsonInput(rawJsonText);
  if (rawRecords.length === 0) {
    return { expandedCount: 0, matchedCount: 0, startCol: 3 };
  }

  const facilityMasterList = loadFacilityMasterList(spreadsheet);
  const startCol = findNextAvailableColumn(inputSheet);
  const numCols = rawRecords.length;

  // 4行目（ステータス行）＋ 5〜19行目（15行分）の列マトリクスを構築
  const statusRowValues = [];
  const rowCount = SANTE_CONFIG.INPUT_LAYOUT.ROW_END - SANTE_CONFIG.INPUT_LAYOUT.ROW_START + 1; // 15
  const columnMatrix = Array.from({ length: rowCount }, () => new Array(numCols).fill(''));
  const targetNotes = new Array(numCols).fill('');
  const facilityNotes = new Array(numCols).fill('');

  let matchedCount = 0;

  rawRecords.forEach((rawItem, colIdx) => {
    const rec = normalizeRawJsonRecord(rawItem);

    // 施設マスターによる名寄せ（正式名称・更新用キー・FAX番号の取得）
    const resolvedFacility = resolveFacility(rec.facilityName, facilityMasterList);
    const officialFacilityName =
      resolvedFacility.matchType === 'partial'
        ? rec.facilityName
        : resolvedFacility.officialName || rec.facilityName;
    if (resolvedFacility.matchType === 'partial') {
      const candidateNames = resolvedFacility.candidates.map((c) => c.officialName).join('／');
      facilityNotes[colIdx] = `施設名が部分一致のみ: 候補 ${candidateNames}。マスター別名に登録するか正式名称を入力してください`;
    }

    // 15行目: 前方 かつ 入院 の場合は初期値 TRUE、外来等は FALSE（手動チェック可）
    const defaultNeedReport =
      rec.phase === SANTE_CONFIG.PHASES.FORWARD &&
      rec.visitType === SANTE_CONFIG.VISIT_TYPES.INPATIENT;
    const needFirstReport =
      rec.needFirstReport !== null ? rec.needFirstReport : defaultNeedReport;

    // 18行目: 連携フェーズが「返書」または「後方」の場合、5重ロック自動照合を実行
    let targetRecordInfo = '';
    const isReplyOrBackward =
      rec.phase === SANTE_CONFIG.PHASES.REPLY || rec.phase === SANTE_CONFIG.PHASES.BACKWARD;

    if (isReplyOrBackward) {
      const matchResult = findFiveLockTarget(
        {
          visitDate: rec.visitDate,
          patientId: rec.patientId,
          patientName: rec.patientName,
          phase: rec.phase,
          facilityName: officialFacilityName,
          facilityDoctor: rec.facilityDoctor,
          lookupKey: resolvedFacility.lookupKey,
        },
        facilityMasterList,
        spreadsheet
      );

      if (matchResult.matched) {
        targetRecordInfo = matchResult.formattedTargetInfo;
        matchedCount++;
        targetNotes[colIdx] = `5重ロック照合一致: ${matchResult.ledgerSheetName} 行#${matchResult.ledgerRowNumber}`;
      } else {
        targetRecordInfo = '';
        targetNotes[colIdx] = '5重ロック照合: 該当する未返書の前方紹介レコードが見つかりませんでした。';
      }
    }

    statusRowValues.push(SANTE_CONFIG.COLUMN_SYNC_STATUS.PENDING);

    // 5行目〜19行目（index 0〜14）へのマッピング
    columnMatrix[0][colIdx] = rec.visitDate;           // 5行目: 受診日
    columnMatrix[1][colIdx] = rec.patientId;           // 6行目: 患者ID
    columnMatrix[2][colIdx] = rec.patientName;         // 7行目: 患者氏名
    columnMatrix[3][colIdx] = rec.phase;               // 8行目: 連携フェーズ
    columnMatrix[4][colIdx] = rec.visitType;           // 9行目: 受入・送出形態
    columnMatrix[5][colIdx] = officialFacilityName;    // 10行目: 連携先施設名（名寄せ後正式名称）
    columnMatrix[6][colIdx] = rec.facilityDept;        // 11行目: 連携先診療科
    columnMatrix[7][colIdx] = rec.facilityDoctor;      // 12行目: 連携先医師名
    columnMatrix[8][colIdx] = rec.ourDept;             // 13行目: 当院担当診療科
    columnMatrix[9][colIdx] = rec.ourDoctor;           // 14行目: 担当医師
    columnMatrix[10][colIdx] = needFirstReport;        // 15行目: 📝 第一報作成
    columnMatrix[11][colIdx] = '';                     // 16行目: 🔗 報告書ドキュメント
    columnMatrix[12][colIdx] = false;                  // 17行目: ✅ FAX送信完了
    columnMatrix[13][colIdx] = targetRecordInfo;       // 18行目: 🎯 返書消込ターゲットID
    columnMatrix[14][colIdx] = rec.remarks;            // 19行目: 備考・連絡事項
  });

  // 入力規則とチェックボックスを設定した上で値を一括書き込み
  applyInputColumnValidations(inputSheet, startCol, numCols);

  inputSheet
    .getRange(SANTE_CONFIG.INPUT_LAYOUT.STATUS_ROW, startCol, 1, numCols)
    .setValues([statusRowValues])
    .setHorizontalAlignment('center')
    .setFontWeight('bold')
    .setBackground('#fff2cc');

  inputSheet
    .getRange(SANTE_CONFIG.INPUT_LAYOUT.ROW_START, startCol, rowCount, numCols)
    .setValues(columnMatrix);

  // 10行目に部分一致候補の補足メモを付与（一致時は空文字でクリア）
  inputSheet
    .getRange(SANTE_CONFIG.INPUT_LAYOUT.ROWS.FACILITY_NAME, startCol, 1, numCols)
    .setNotes([facilityNotes]);

  // 18行目に照合結果の補足メモを付与
  inputSheet
    .getRange(SANTE_CONFIG.INPUT_LAYOUT.ROWS.TARGET_RECORD_INFO, startCol, 1, numCols)
    .setNotes([targetNotes]);

  spreadsheet.toast(
    `${numCols}件の患者データを横展開しました（5重ロック自動照合ヒット: ${matchedCount}件）`,
    '✅ SANTE-BRIDGE 展開完了',
    6
  );

  return {
    expandedCount: numCols,
    matchedCount: matchedCount,
    startCol: startCol,
  };
}

/**
 * 紹介・逆紹介入力シートのC列以降の全有効列データを読み取る
 * @param {GoogleAppsScript.Spreadsheet.Sheet} inputSheet
 * @return {Array<{
 *   colNumber: number,
 *   syncStatus: string,
 *   visitDate: string,
 *   patientId: string,
 *   patientName: string,
 *   phase: string,
 *   visitType: string,
 *   facilityName: string,
 *   facilityDept: string,
 *   facilityDoctor: string,
 *   ourDept: string,
 *   ourDoctor: string,
 *   needFirstReport: boolean,
 *   reportDocUrl: string,
 *   faxSent: boolean,
 *   targetRecordInfo: string,
 *   remarks: string
 * }>}
 */
function readAllInputColumns(inputSheet) {
  const startCol = SANTE_CONFIG.INPUT_LAYOUT.DATA_START_COL;
  const lastCol = inputSheet.getLastColumn();
  if (lastCol < startCol) return [];

  const numCols = lastCol - startCol + 1;
  const statusRow = SANTE_CONFIG.INPUT_LAYOUT.STATUS_ROW;
  const totalRows = SANTE_CONFIG.INPUT_LAYOUT.ROW_END - statusRow + 1; // 4行目〜19行目 (16行)

  const values = inputSheet.getRange(statusRow, startCol, totalRows, numCols).getValues();
  const columns = [];

  for (let c = 0; c < numCols; c++) {
    const visitDate = formatDateYMD(values[1][c]);
    const patientId = String(values[2][c] || '').trim();
    const patientName = String(values[3][c] || '').trim();
    const phase = String(values[4][c] || '').trim();
    const facilityName = String(values[6][c] || '').trim();

    // 完全に空の列はスキップ
    if (!visitDate && !patientId && !patientName && !phase && !facilityName) {
      continue;
    }

    columns.push({
      colNumber: startCol + c,
      syncStatus: String(values[0][c] || '').trim(),
      visitDate: visitDate || formatDateYMD(new Date()),
      patientId: patientId,
      patientName: patientName,
      phase: phase ? normalizePhaseValue(phase) : SANTE_CONFIG.PHASES.FORWARD,
      visitType: normalizeVisitTypeValue(values[5][c]),
      facilityName: facilityName,
      facilityDept: String(values[7][c] || '').trim(),
      facilityDoctor: String(values[8][c] || '').trim(),
      ourDept: String(values[9][c] || '').trim(),
      ourDoctor: String(values[10][c] || '').trim(),
      needFirstReport: values[11][c] === true || String(values[11][c]).toUpperCase() === 'TRUE',
      reportDocUrl: String(values[12][c] || '').trim(),
      faxSent: values[13][c] === true || String(values[13][c]).toUpperCase() === 'TRUE',
      targetRecordInfo: String(values[14][c] || '').trim(),
      remarks: String(values[15][c] || '').trim(),
    });
  }

  return columns;
}

/**
 * 【要件B: 実行ボタンによる一括処理】
 * 1. 第一報（受診・入院報告書）Googleドキュメント生成
 * 2. kintone（アプリB / ID: 195）への一括POST（新規登録）と年度別台帳追記
 * 3. kintone（アプリB）への一括PUT（返書日消込）と年度別台帳更新
 *
 * @return {{
 *   docsGenerated: number,
 *   kintonePosted: number,
 *   kintonePutUpdated: number,
 *   ledgerUpdated: number
 * }}
 */
function executeBatchProcess() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '現在別のプロセスが一括処理を実行中です。しばらく待ってから再実行してください。',
      '⏳ 処理中',
      5
    );
    return { docsGenerated: 0, kintonePosted: 0, kintonePutUpdated: 0, ledgerUpdated: 0 };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName(SANTE_CONFIG.SHEETS.INPUT);
  if (!inputSheet) {
    lock.releaseLock();
    throw new Error(`シート「${SANTE_CONFIG.SHEETS.INPUT}」が見つかりません。`);
  }

  const rows = SANTE_CONFIG.INPUT_LAYOUT.ROWS;
  const statusRow = SANTE_CONFIG.INPUT_LAYOUT.STATUS_ROW;
  const facilityMasterList = loadFacilityMasterList(ss);
  const columns = readAllInputColumns(inputSheet);

  if (columns.length === 0) {
    lock.releaseLock();
    ss.toast('処理対象となる患者データ列がありません。', 'ℹ️ SANTE-BRIDGE', 5);
    return { docsGenerated: 0, kintonePosted: 0, kintonePutUpdated: 0, ledgerUpdated: 0 };
  }

  let docsGenerated = 0;
  let kintonePosted = 0;
  let kintonePutUpdated = 0;
  let ledgerUpdated = 0;

  try {
    // =========================================================================
    // Step 1: 第一報（受診・入院報告書）Googleドキュメント生成
    // （15行目 `📝 第一報作成` が TRUE かつ 16行目が空欄の列）
    // =========================================================================
    columns.forEach((colData) => {
      if (colData.needFirstReport && !colData.reportDocUrl) {
        const docResult = generateFirstReportDoc(colData, facilityMasterList);
        colData.reportDocUrl = docResult.docUrl;

        inputSheet.getRange(rows.REPORT_DOC_URL, colData.colNumber).setValue(docResult.docUrl);
        if (colData.syncStatus !== SANTE_CONFIG.COLUMN_SYNC_STATUS.SYNCED) {
          inputSheet
            .getRange(statusRow, colData.colNumber)
            .setValue(SANTE_CONFIG.COLUMN_SYNC_STATUS.DOC_READY);
        }
        docsGenerated++;
      }
    });

    // =========================================================================
    // Step 2: kintone（アプリB / ID: 195）への一括POST（新規登録）と台帳追記
    // （連携フェーズが「前方」または「後方」の未連携列。「返書」単体はPOSTしない）
    // =========================================================================
    const postCandidateCols = columns.filter((colData) => {
      if (colData.syncStatus === SANTE_CONFIG.COLUMN_SYNC_STATUS.SYNCED) {
        return false;
      }
      return (
        colData.phase === SANTE_CONFIG.PHASES.FORWARD ||
        colData.phase === SANTE_CONFIG.PHASES.BACKWARD
      );
    });

    if (postCandidateCols.length > 0) {
      const nowStr = formatDateTimeNow();
      const postRecordsPayload = postCandidateCols.map((colData) => {
        const resolved = resolveFacility(colData.facilityName, facilityMasterList);
        const officialFacilityName =
          resolved.matchType === 'partial'
            ? colData.facilityName
            : resolved.officialName || colData.facilityName;

        // 施設マスター名寄せ結果で入力シート10行目も最新化
        const facilityCell = inputSheet.getRange(rows.FACILITY_NAME, colData.colNumber);
        if (resolved.matchType === 'partial') {
          const candidateNames = resolved.candidates.map((c) => c.officialName).join('／');
          facilityCell.setNote(
            `施設名が部分一致のみ: 候補 ${candidateNames}。マスター別名に登録するか正式名称を入力してください`
          );
        } else {
          if (resolved.matchType === 'exact' || resolved.matchType === 'core') {
            facilityCell.clearNote();
          }
          if (officialFacilityName !== colData.facilityName) {
            facilityCell.setValue(officialFacilityName);
            colData.facilityName = officialFacilityName;
          }
        }

        const isForward = colData.phase === SANTE_CONFIG.PHASES.FORWARD;
        const replyStatus = isForward
          ? SANTE_CONFIG.REPLY_STATUS.UNCREATED
          : SANTE_CONFIG.REPLY_STATUS.NOT_APPLICABLE;

        let faxStatus = SANTE_CONFIG.FAX_STATUS.NOT_REQUIRED;
        if (colData.faxSent) {
          faxStatus = SANTE_CONFIG.FAX_STATUS.SENT;
        } else if (colData.needFirstReport) {
          faxStatus = SANTE_CONFIG.FAX_STATUS.UNSENT;
        }

        return {
          _colData: colData,
          registeredAt: nowStr,
          visitDate: colData.visitDate,
          patientId: colData.patientId,
          patientName: colData.patientName,
          phase: colData.phase,
          visitType: colData.visitType,
          lookupKey: resolved.lookupKey,
          facilityName: officialFacilityName,
          facilityDept: colData.facilityDept,
          facilityDoctor: colData.facilityDoctor,
          ourDept: colData.ourDept,
          ourDoctor: colData.ourDoctor,
          replyStatus: replyStatus,
          replyDate: '',
          firstReportUrl: colData.reportDocUrl,
          faxStatus: faxStatus,
          remarks: colData.remarks,
        };
      });

      // kintone `/k/v1/records.json` (POST) 一括登録
      const postResponse = postRecordsToKintoneBatch(postRecordsPayload);
      const issuedIds = postResponse.ids;

      // 発番された kintoneレコードID を付与して年度別台帳の最下行へ一括追記
      const ledgerRowsToAppend = postRecordsPayload.map((item, idx) => ({
        ...item,
        kintoneRecordId: issuedIds[idx] || '',
      }));

      appendRecordsToLedger(ledgerRowsToAppend, ss);
      kintonePosted = ledgerRowsToAppend.length;

      // 入力シートのステータスを「✅ kintone反映済」に更新
      postRecordsPayload.forEach((item, idx) => {
        const colNum = item._colData.colNumber;
        const recId = issuedIds[idx] || '';
        inputSheet
          .getRange(statusRow, colNum)
          .setValue(SANTE_CONFIG.COLUMN_SYNC_STATUS.SYNCED)
          .setBackground('#d9ead3');
        inputSheet
          .getRange(rows.VISIT_DATE, colNum)
          .setNote(`kintone登録済: レコードID #${recId} (${nowStr})`);
        item._colData.syncStatus = SANTE_CONFIG.COLUMN_SYNC_STATUS.SYNCED;
      });
    }

    // =========================================================================
    // Step 3: kintone（アプリB）への一括PUT（返書日消込）と年度別台帳更新
    // （18行目に `kintone#<レコードID>` が入っている列）
    // ※ 連携フェーズが純粋な「返書（経過・結果報告）」の場合はStep 2をスキップし、本Step 3のみ実行される
    // =========================================================================
    const putCandidateItems = [];
    columns.forEach((colData) => {
      if (!colData.targetRecordInfo) return;

      // 既に返書消込が完了している純粋返書列は二重PUTしない
      const targetCellNote = inputSheet
        .getRange(rows.TARGET_RECORD_INFO, colData.colNumber)
        .getNote();
      if (targetCellNote && targetCellNote.includes('[消込完了]')) {
        return;
      }

      const parsedTarget = parseTargetRecordInfo(colData.targetRecordInfo);
      if (!parsedTarget || !parsedTarget.kintoneRecordId) return;

      const replyDate = colData.visitDate || formatDateYMD(new Date());
      putCandidateItems.push({
        colData: colData,
        kintoneRecordId: parsedTarget.kintoneRecordId,
        ledgerRowNumber: parsedTarget.ledgerRowNumber,
        fiscalYear: parsedTarget.fiscalYear,
        replyDate: replyDate,
      });
    });

    if (putCandidateItems.length > 0) {
      // kintone `/k/v1/records.json` (PUT) で返書日・返書ステータスを一括更新
      const putResult = putReplyDatesToKintoneBatch(
        putCandidateItems.map((item) => ({
          kintoneRecordId: item.kintoneRecordId,
          replyDate: item.replyDate,
          replyStatus: SANTE_CONFIG.REPLY_STATUS.COMPLETED,
        }))
      );
      kintonePutUpdated = putResult.updatedIds.length;

      // 年度別台帳タブの該当行の返書ステータスを "作成・送付完了"、返書日を今回の日付に更新
      ledgerUpdated = updateLedgerReplyRecords(putCandidateItems, ss);

      // 入力シート18行目のメモと4行目ステータスを更新
      const nowStr = formatDateTimeNow();
      putCandidateItems.forEach((item) => {
        const colNum = item.colData.colNumber;
        const targetCell = inputSheet.getRange(rows.TARGET_RECORD_INFO, colNum);
        const prevNote = targetCell.getNote();
        targetCell.setNote(
          `${prevNote ? prevNote + '\n' : ''}[消込完了] kintone#${item.kintoneRecordId} / 返書日:${item.replyDate} (${nowStr})`
        );

        inputSheet
          .getRange(statusRow, colNum)
          .setValue(SANTE_CONFIG.COLUMN_SYNC_STATUS.SYNCED)
          .setBackground('#d9ead3');
      });
    }

    // 完了通知（トースト ＆ Slack通知）
    const summaryMsg = `第一報生成: ${docsGenerated}件 / kintone新規POST: ${kintonePosted}件 / 返書消込PUT: ${kintonePutUpdated}件 (台帳消込: ${ledgerUpdated}件)`;
    ss.toast(summaryMsg, '🎉 SANTE-BRIDGE 一括処理完了', 8);

    if (docsGenerated > 0 || kintonePosted > 0 || kintonePutUpdated > 0) {
      notifySlackMilestone('地域連携バッチ処理完了', summaryMsg, {
        '第一報ドキュメント生成': `${docsGenerated}件`,
        'kintone新規登録(POST)': `${kintonePosted}件`,
        'kintone返書消込(PUT)': `${kintonePutUpdated}件`,
        '台帳ステータス更新': `${ledgerUpdated}件`,
      });
    }

    return {
      docsGenerated,
      kintonePosted,
      kintonePutUpdated,
      ledgerUpdated,
    };
  } catch (err) {
    notifySlackError('一括処理 (executeBatchProcess) 異常終了', err);
    ss.toast(`エラーが発生しました: ${err.message}`, '🚨 処理中断', 10);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/**
 * メニュー用補助関数: 第一報ドキュメントのみ先行生成する
 */
function executeFirstReportGenerationOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName(SANTE_CONFIG.SHEETS.INPUT);
  if (!inputSheet) return;

  const facilityMasterList = loadFacilityMasterList(ss);
  const columns = readAllInputColumns(inputSheet);
  const rows = SANTE_CONFIG.INPUT_LAYOUT.ROWS;
  let count = 0;

  columns.forEach((colData) => {
    if (colData.needFirstReport && !colData.reportDocUrl) {
      const res = generateFirstReportDoc(colData, facilityMasterList);
      inputSheet.getRange(rows.REPORT_DOC_URL, colData.colNumber).setValue(res.docUrl);
      count++;
    }
  });

  ss.toast(`第一報Googleドキュメントを ${count}件 生成しました。`, '📝 第一報生成完了', 6);
}

/**
 * メニュー用補助関数: 現在入力シートに展開されている全列に対して5重ロック照合を再実行する
 */
function reRunFiveLockMatchingForAllColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName(SANTE_CONFIG.SHEETS.INPUT);
  if (!inputSheet) return;

  const facilityMasterList = loadFacilityMasterList(ss);
  const columns = readAllInputColumns(inputSheet);
  const rows = SANTE_CONFIG.INPUT_LAYOUT.ROWS;
  let matchedCount = 0;

  columns.forEach((colData) => {
    const resolved = resolveFacility(colData.facilityName, facilityMasterList);
    const facilityCell = inputSheet.getRange(rows.FACILITY_NAME, colData.colNumber);
    if (resolved.matchType === 'partial') {
      const candidateNames = resolved.candidates.map((c) => c.officialName).join('／');
      facilityCell.setNote(
        `施設名が部分一致のみ: 候補 ${candidateNames}。マスター別名に登録するか正式名称を入力してください`
      );
    } else {
      if (resolved.matchType === 'exact' || resolved.matchType === 'core') {
        facilityCell.clearNote();
      }
      if (resolved.officialName && resolved.officialName !== colData.facilityName) {
        facilityCell.setValue(resolved.officialName);
        colData.facilityName = resolved.officialName;
      }
    }

    const isReplyOrBackward =
      colData.phase === SANTE_CONFIG.PHASES.REPLY ||
      colData.phase === SANTE_CONFIG.PHASES.BACKWARD;

    if (isReplyOrBackward) {
      const matchResult = findFiveLockTarget(
        {
          visitDate: colData.visitDate,
          patientId: colData.patientId,
          patientName: colData.patientName,
          phase: colData.phase,
          facilityName: colData.facilityName,
          facilityDoctor: colData.facilityDoctor,
          lookupKey: resolved.lookupKey,
        },
        facilityMasterList,
        ss
      );

      const targetCell = inputSheet.getRange(rows.TARGET_RECORD_INFO, colData.colNumber);
      if (matchResult.matched) {
        targetCell.setValue(matchResult.formattedTargetInfo);
        targetCell.setNote(
          `5重ロック再照合一致: ${matchResult.ledgerSheetName} 行#${matchResult.ledgerRowNumber}`
        );
        matchedCount++;
      }
    }
  });

  ss.toast(`全列の再照合が完了しました（ヒット: ${matchedCount}件）`, '🔍 5重ロック再照合完了', 6);
}

/**
 * 【要件B-4: 送信完了列のクリア機能】
 * 17行目 `✅ FAX送信完了` がチェックされた列、または
 * 15行目 `📝 第一報作成` が FALSE かつ kintone反映が完了した列を安全にクリアし、
 * 未完了の列を左詰め（C列開始）に整理する
 *
 * @return {{clearedCount: number, remainingCount: number}}
 */
function clearCompletedColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName(SANTE_CONFIG.SHEETS.INPUT);
  if (!inputSheet) return { clearedCount: 0, remainingCount: 0 };

  const startCol = SANTE_CONFIG.INPUT_LAYOUT.DATA_START_COL;
  const lastCol = inputSheet.getLastColumn();
  if (lastCol < startCol) {
    ss.toast('クリア対象の列はありません。', '🧹 列整理', 4);
    return { clearedCount: 0, remainingCount: 0 };
  }

  const columns = readAllInputColumns(inputSheet);
  const statusRow = SANTE_CONFIG.INPUT_LAYOUT.STATUS_ROW;
  const rowEnd = SANTE_CONFIG.INPUT_LAYOUT.ROW_END;
  const totalRows = rowEnd - statusRow + 1; // 4〜19行目 (16行)

  const remainingColumns = [];
  let clearedCount = 0;

  columns.forEach((colData) => {
    const isSynced = colData.syncStatus === SANTE_CONFIG.COLUMN_SYNC_STATUS.SYNCED;
    // FAX送信完了チェック済（かつkintone連携済、またはFAX完了優先）または（第一報不要 かつ kintone反映済）
    const canClear =
      (colData.faxSent && isSynced) ||
      (colData.faxSent && colData.phase === SANTE_CONFIG.PHASES.OTHER) ||
      (!colData.needFirstReport && isSynced);

    if (canClear) {
      clearedCount++;
    } else {
      const colRange = inputSheet.getRange(statusRow, colData.colNumber, totalRows, 1);
      remainingColumns.push({
        values: colRange.getValues(),
        notes: colRange.getNotes(),
        syncStatus: colData.syncStatus,
      });
    }
  });

  if (clearedCount === 0) {
    ss.toast(
      'クリア条件（kintone反映済 かつ FAX送信完了 または 第一報不要）を満たす列がありません。',
      'ℹ️ 列整理',
      6
    );
    return { clearedCount: 0, remainingCount: columns.length };
  }

  // C列以降の既存データ領域を一旦クリア
  const numExistingCols = lastCol - startCol + 1;
  const clearRange = inputSheet.getRange(statusRow, startCol, totalRows, numExistingCols);
  clearRange.clearContent();
  clearRange.clearNote();
  clearRange.clearDataValidations();
  clearRange.setBackground(null);

  // 未完了の列が残っている場合はC列から左詰めで再配置
  if (remainingColumns.length > 0) {
    const remCount = remainingColumns.length;
    applyInputColumnValidations(inputSheet, startCol, remCount);

    const newValues = Array.from({ length: totalRows }, () => new Array(remCount).fill(''));
    const newNotes = Array.from({ length: totalRows }, () => new Array(remCount).fill(''));

    remainingColumns.forEach((colObj, cIdx) => {
      for (let r = 0; r < totalRows; r++) {
        newValues[r][cIdx] = colObj.values[r][0];
        newNotes[r][cIdx] = colObj.notes[r][0];
      }
    });

    const targetRange = inputSheet.getRange(statusRow, startCol, totalRows, remCount);
    targetRange.setValues(newValues);
    targetRange.setNotes(newNotes);

    // 4行目の背景色を復元
    for (let cIdx = 0; cIdx < remCount; cIdx++) {
      const cell = inputSheet.getRange(statusRow, startCol + cIdx);
      cell.setHorizontalAlignment('center').setFontWeight('bold');
      if (remainingColumns[cIdx].syncStatus === SANTE_CONFIG.COLUMN_SYNC_STATUS.SYNCED) {
        cell.setBackground('#d9ead3');
      } else {
        cell.setBackground('#fff2cc');
      }
    }
  }

  ss.toast(
    `完了済みの ${clearedCount}列 をクリアし、未完了 ${remainingColumns.length}列 を左詰め整理しました。`,
    '🧹 列クリア完了',
    6
  );

  return {
    clearedCount: clearedCount,
    remainingCount: remainingColumns.length,
  };
}

/**
 * 3タブ構成（①紹介・逆紹介入力、②当年度_連携台帳、③施設マスター）を初期構築・フォーマット整備する
 */
function setupAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. タブ①: 紹介・逆紹介入力
  let inputSheet = ss.getSheetByName(SANTE_CONFIG.SHEETS.INPUT);
  if (!inputSheet) {
    inputSheet = ss.insertSheet(SANTE_CONFIG.SHEETS.INPUT, 0);
  }

  // B2:B3 ラベルと C2:C3 JSONペースト領域の設定
  inputSheet.getRange('B2').setValue('📥 Gemini JSON貼付セル').setFontWeight('bold');
  inputSheet.getRange('B3').setValue('（C2に貼付で自動展開）').setFontColor('#666666').setFontSize(9);

  const pasteRange = inputSheet.getRange(SANTE_CONFIG.INPUT_LAYOUT.PASTE_RANGE);
  pasteRange
    .setBackground('#e8f0fe')
    .setBorder(true, true, true, true, false, false, '#1a73e8', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // B4: ステータスラベル
  inputSheet
    .getRange(SANTE_CONFIG.INPUT_LAYOUT.STATUS_ROW, SANTE_CONFIG.INPUT_LAYOUT.LABEL_COL)
    .setValue('🔄 連携ステータス')
    .setBackground('#1f4e78')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  // B5:B19 項目ラベルの設定
  const labels = SANTE_CONFIG.INPUT_LAYOUT.LABELS.map((l) => [l]);
  const labelRange = inputSheet.getRange(
    SANTE_CONFIG.INPUT_LAYOUT.ROW_START,
    SANTE_CONFIG.INPUT_LAYOUT.LABEL_COL,
    labels.length,
    1
  );
  labelRange
    .setValues(labels)
    .setBackground('#d9e1f2')
    .setFontWeight('bold')
    .setBorder(true, true, true, true, true, true, '#8ea9db', SpreadsheetApp.BorderStyle.SOLID);

  inputSheet.setColumnWidth(1, 25);  // A列（余白）
  inputSheet.setColumnWidth(2, 210); // B列（項目ラベル）
  inputSheet.setFrozenColumns(2);    // A〜B列（項目ラベル列）を固定
  inputSheet.setFrozenRows(4);       // 1〜4行目（JSON貼付枠 C2:C3 ＆ ステータス行）を上部に常時固定

  // 2. タブ②: 当年度_連携台帳（例: 2026年度_連携台帳）の生成・UI固定・条件付き書式適用
  const currentFy = getFiscalYear();
  getOrCreateLedgerSheet(currentFy, ss, true);
  const prevLedger = ss.getSheetByName(getLedgerSheetName(currentFy - 1));
  if (prevLedger) {
    applyLedgerSheetFormatting(prevLedger);
  }

  // 3. タブ③: 施設マスター
  let masterSheet = ss.getSheetByName(SANTE_CONFIG.SHEETS.FACILITY_MASTER);
  if (!masterSheet) {
    masterSheet = ss.insertSheet(SANTE_CONFIG.SHEETS.FACILITY_MASTER);
    const masterHeaders = SANTE_CONFIG.FACILITY_MASTER_COLUMNS.HEADERS;
    const headerRange = masterSheet.getRange(1, 1, 1, masterHeaders.length);
    headerRange
      .setValues([masterHeaders])
      .setBackground('#274e13')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    masterSheet.setFrozenRows(1);
  }

  ss.toast('3タブ構成（紹介・逆紹介入力 / 年度別連携台帳 / 施設マスター）の整備が完了しました。', '📋 初期構築完了', 6);
}

/**
 * メニュー用: 第一報テンプレートGoogleドキュメントを自動作成しURLを通知する
 */
function handleCreateDefaultTemplateMenu() {
  const res = createDefaultReportTemplateDoc();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `テンプレートを作成しました: ${res.docUrl}`,
    '📄 第一報テンプレート作成完了',
    8
  );
}

/**
 * メニュー用: Slack Webhookの疎通テストを実行する
 */
function testSlackWebhookConnection() {
  const ok = sendSlackNotification(
    'Slack連携 疎通テスト',
    'SANTE-BRIDGE (サンテ溝上病院 地域連携システム) からのWebhook通知テストが正常に完了しました。',
    'SUCCESS',
    {
      実行者: Session.getActiveUser().getEmail() || '地域連携室スタッフ',
      対象年度台帳: getLedgerSheetName(),
    }
  );
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ok) {
    ss.toast('Slackチャンネルへテスト通知を送信しました。', '🔔 Slackテスト成功', 5);
  } else {
    ss.toast(
      'Slack通知に失敗しました。スクリプトプロパティ「SLACK_WEBHOOK_URL」を確認してください。',
      '⚠️ Slackテスト未送信',
      8
    );
  }
}
