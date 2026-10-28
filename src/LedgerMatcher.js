/**
 * @fileoverview 施設マスター名寄せ・年度別台帳管理・5重ロック自動照合エンジン
 * サンテ溝上病院 地域連携・紹介返書管理＆第一報自動化システム (SANTE-BRIDGE)
 */

/**
 * 文字列の空白（半角・全角）をすべて除去し、全角英数字を半角へ正規化する
 * @param {string|number|null|undefined} str
 * @return {string}
 */
function normalizeCompactString(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, (s) => String.fromCharCode(s.charCodeAt(0) - 0xfee0))
    .replace(/[\s\u3000]+/g, '')
    .trim();
}

/**
 * 施設名比較用に法人格接頭辞や空白を除去したコア名称を返す
 * @param {string} name
 * @return {string}
 */
function normalizeFacilityCoreName(name) {
  const compact = normalizeCompactString(name);
  if (!compact) return '';
  return compact.replace(
    /^(医療法人社団|医療法人財団|社会医療法人|医療法人|独立行政法人国立病院機構|独立行政法人|地方独立行政法人|公益財団法人|公益社団法人|一般社団法人|一般財団法人|社会福祉法人|特定医療法人)/,
    ''
  );
}

/**
 * 施設マスターシートから全レコードを読み込み、メモリ上の検索リストとして返す
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ss
 * @return {Array<{
 *   lookupKey: string,
 *   officialName: string,
 *   address: string,
 *   faxNumber: string,
 *   aliases: Array<string>,
 *   normalizedCandidates: Array<string>,
 *   coreCandidates: Array<string>
 * }>}
 */
function loadFacilityMasterList(ss) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SANTE_CONFIG.SHEETS.FACILITY_MASTER);
  if (!sheet) {
    Logger.log(`警告: シート「${SANTE_CONFIG.SHEETS.FACILITY_MASTER}」が見つかりません。`);
    return [];
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 2) {
    return [];
  }

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const idx = SANTE_CONFIG.FACILITY_MASTER_COLUMNS.INDEX;
  const masterList = [];

  values.forEach((row) => {
    const lookupKey = String(row[idx.LOOKUP_KEY] || '').trim();
    const officialName = String(row[idx.OFFICIAL_NAME] || '').trim();
    if (!officialName) return;

    const address = String(row[idx.ADDRESS] || '').trim();
    const faxNumber = String(row[idx.FAX_NUMBER] || '').trim();

    // E列以降のすべてのセル（およびセル内のカンマ・読点区切り）を別名リストとして展開
    const rawAliases = [];
    for (let c = idx.ALIAS_START; c < row.length; c++) {
      const cellVal = String(row[c] || '').trim();
      if (cellVal) {
        cellVal.split(/[、,，／\/]+/).forEach((part) => {
          const trimmed = part.trim();
          if (trimmed) rawAliases.push(trimmed);
        });
      }
    }

    const allNames = [officialName, ...rawAliases];
    const normalizedCandidates = allNames
      .map((n) => normalizeCompactString(n))
      .filter(Boolean);
    const coreCandidates = allNames
      .map((n) => normalizeFacilityCoreName(n))
      .filter(Boolean);

    masterList.push({
      lookupKey,
      officialName,
      address,
      faxNumber,
      aliases: rawAliases,
      normalizedCandidates,
      coreCandidates,
    });
  });

  return masterList;
}

/**
 * 連携先施設名を施設マスターと照合し、正式名称・更新用キー・FAX番号へ解決（名寄せ）する
 * @param {string} rawFacilityName OCR抽出または手入力された施設名
 * @param {Array<Object>=} facilityMasterList 事前ロード済みの施設マスター配列（省略時は自動取得）
 * @return {{
 *   lookupKey: string,
 *   officialName: string,
 *   address: string,
 *   faxNumber: string,
 *   isMatched: boolean
 * }}
 */
function resolveFacility(rawFacilityName, facilityMasterList) {
  const trimmedInput = String(rawFacilityName || '').trim();
  if (!trimmedInput) {
    return {
      lookupKey: '',
      officialName: '',
      address: '',
      faxNumber: '',
      isMatched: false,
    };
  }

  const list = facilityMasterList || loadFacilityMasterList();
  const normInput = normalizeCompactString(trimmedInput);
  const coreInput = normalizeFacilityCoreName(trimmedInput);

  // Step 1: 完全一致（正式名称 または 略称・別名リスト）
  for (let i = 0; i < list.length; i++) {
    const item = list[i];
    if (item.normalizedCandidates.includes(normInput)) {
      return {
        lookupKey: item.lookupKey,
        officialName: item.officialName,
        address: item.address,
        faxNumber: item.faxNumber,
        isMatched: true,
      };
    }
  }

  // Step 2: 法人格を除いたコア名称での完全一致
  if (coreInput) {
    for (let i = 0; i < list.length; i++) {
      const item = list[i];
      if (item.coreCandidates.includes(coreInput)) {
        return {
          lookupKey: item.lookupKey,
          officialName: item.officialName,
          address: item.address,
          faxNumber: item.faxNumber,
          isMatched: true,
        };
      }
    }
  }

  // Step 3: 部分一致（2文字以上の場合のみ）
  if (coreInput.length >= 2) {
    for (let i = 0; i < list.length; i++) {
      const item = list[i];
      const matched = item.coreCandidates.some(
        (cand) => cand.length >= 2 && (cand.includes(coreInput) || coreInput.includes(cand))
      );
      if (matched) {
        return {
          lookupKey: item.lookupKey,
          officialName: item.officialName,
          address: item.address,
          faxNumber: item.faxNumber,
          isMatched: true,
        };
      }
    }
  }

  // 未ヒット時は元の施設名を維持して返却
  return {
    lookupKey: '',
    officialName: trimmedInput,
    address: '',
    faxNumber: '',
    isMatched: false,
  };
}

/**
 * 指定された会計年度の台帳シートを取得する（存在しない場合は新規作成・ヘッダー初期化する）
 * @param {number=} fiscalYear 会計年度（例: 2026）
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ss
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateLedgerSheet(fiscalYear, ss) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = getLedgerSheetName(fiscalYear);
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    const headers = SANTE_CONFIG.LEDGER_COLUMNS.HEADERS;
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange
      .setBackground('#1f4e78')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * 照合対象となる年度別台帳シート群（当年度＋存在する場合は前年度）を取得する
 * @param {string|Date=} baseDate
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ss
 * @return {Array<{sheet: GoogleAppsScript.Spreadsheet.Sheet, fiscalYear: number, sheetName: string}>}
 */
function getCandidateLedgerSheets(baseDate, ss) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  const currentFy = getFiscalYear(baseDate);
  const prevFy = currentFy - 1;
  const result = [];

  // 当年度台帳（存在しない場合は自動生成）
  const currentSheet = getOrCreateLedgerSheet(currentFy, spreadsheet);
  result.push({
    sheet: currentSheet,
    fiscalYear: currentFy,
    sheetName: currentSheet.getName(),
  });

  // 前年度台帳（シートが存在する場合のみ参照対象に追加）
  const prevSheetName = getLedgerSheetName(prevFy);
  const prevSheet = spreadsheet.getSheetByName(prevSheetName);
  if (prevSheet) {
    result.push({
      sheet: prevSheet,
      fiscalYear: prevFy,
      sheetName: prevSheetName,
    });
  }

  return result;
}

/**
 * 【ロック1: 患者特定】
 * 患者IDが一致（ID空欄時はスペース除去後の患者氏名が一致）するかを判定する
 * @param {string} inputPatientId
 * @param {string} inputPatientName
 * @param {string} ledgerPatientId
 * @param {string} ledgerPatientName
 * @return {boolean}
 */
function isLock1PatientMatched(inputPatientId, inputPatientName, ledgerPatientId, ledgerPatientName) {
  const normInputId = normalizeCompactString(inputPatientId);
  const normLedgerId = normalizeCompactString(ledgerPatientId);

  // 両方に患者IDが存在する場合は患者IDで判定（先頭ゼロ有無の揺れも許容）
  if (normInputId !== '' && normLedgerId !== '') {
    if (normInputId === normLedgerId) return true;
    const strippedInput = normInputId.replace(/^0+/, '');
    const strippedLedger = normLedgerId.replace(/^0+/, '');
    return strippedInput !== '' && strippedInput === strippedLedger;
  }

  // いずれかの患者IDが空欄の場合はスペース除去後の患者氏名で判定
  const normInputName = normalizeCompactString(inputPatientName);
  const normLedgerName = normalizeCompactString(ledgerPatientName);
  if (!normInputName || !normLedgerName) return false;

  return normInputName === normLedgerName;
}

/**
 * 【ロック2: 施設特定】
 * 名寄せ後の連携先施設名が一致（または部分一致）するかを判定する
 * @param {{lookupKey: string, officialName: string}} inputFacility
 * @param {string} ledgerLookupKey
 * @param {string} ledgerFacilityName
 * @param {Array<Object>} facilityMasterList
 * @return {boolean}
 */
function isLock2FacilityMatched(inputFacility, ledgerLookupKey, ledgerFacilityName, facilityMasterList) {
  const normLedgerKey = normalizeCompactString(ledgerLookupKey);
  const normInputKey = normalizeCompactString(inputFacility.lookupKey);

  // 更新用キーが双方に存在し一致すれば確定一致
  if (normInputKey && normLedgerKey && normInputKey === normLedgerKey) {
    return true;
  }

  // 台帳側の施設名も施設マスターで名寄せ解決した上で比較
  const ledgerResolved = resolveFacility(ledgerFacilityName, facilityMasterList);
  if (
    normInputKey &&
    ledgerResolved.lookupKey &&
    normInputKey === normalizeCompactString(ledgerResolved.lookupKey)
  ) {
    return true;
  }

  const coreInput = normalizeFacilityCoreName(inputFacility.officialName);
  const coreLedger = normalizeFacilityCoreName(ledgerResolved.officialName || ledgerFacilityName);

  if (!coreInput || !coreLedger) return false;
  if (coreInput === coreLedger) return true;

  // 部分一致（includes）
  return coreInput.includes(coreLedger) || coreLedger.includes(coreInput);
}

/**
 * 医師名文字列をスペース除去後、`・` `、` `,` `／` `/` で分割し、敬称を除去した配列にする
 * @param {string} rawDoctorStr
 * @return {Array<string>}
 */
function splitAndNormalizeDoctors(rawDoctorStr) {
  const compact = normalizeCompactString(rawDoctorStr);
  if (!compact) return [];

  return compact
    .split(/[・、,，／\/]+/)
    .map((token) => {
      // 末尾の敬称・役職（「先生」「医師」「Dr.」等）を除去したコア医師名も抽出
      const stripped = token
        .replace(/^(Dr\.?|Ｄｒ\.?)/i, '')
        .replace(/(先生|医師|院長|副院長|部長|医長|科長|教授|准教授)$/, '');
      return stripped || token;
    })
    .filter(Boolean);
}

/**
 * 【ロック3: 医師名の「含む」部分一致・連名対応】
 * 台帳側の連携先医師名と今回抽出された連携先医師名について、スペース除去後、
 * `・` `、` `,` `／` `/` で分割した配列同士を比較し、いずれかの医師名が部分一致（includes）していれば一致とみなす
 * @param {string} inputDoctorStr
 * @param {string} ledgerDoctorStr
 * @return {boolean}
 */
function isLock3DoctorMatched(inputDoctorStr, ledgerDoctorStr) {
  const inputDoctors = splitAndNormalizeDoctors(inputDoctorStr);
  const ledgerDoctors = splitAndNormalizeDoctors(ledgerDoctorStr);

  if (inputDoctors.length === 0 || ledgerDoctors.length === 0) {
    return false;
  }

  return ledgerDoctors.some((ledgerDoc) =>
    inputDoctors.some(
      (inputDoc) => ledgerDoc.includes(inputDoc) || inputDoc.includes(ledgerDoc)
    )
  );
}

/**
 * 【ロック4: 未返書状態】
 * 台帳側の返書ステータスが「未作成」（または返書日が空欄）であり、
 * かつ連携フェーズが「前方（紹介・受入）」であるかを判定する
 * @param {string} ledgerPhase
 * @param {string} ledgerReplyStatus
 * @param {string|Date} ledgerReplyDate
 * @return {boolean}
 */
function isLock4UnrepliedForward(ledgerPhase, ledgerReplyStatus, ledgerReplyDate) {
  const phaseStr = String(ledgerPhase || '').trim();
  const isForward =
    phaseStr === SANTE_CONFIG.PHASES.FORWARD || phaseStr.startsWith('前方');
  if (!isForward) {
    return false;
  }

  const statusStr = String(ledgerReplyStatus || '').trim();
  const replyDateStr = formatDateYMD(ledgerReplyDate);

  // 既に「作成・送付完了」かつ返書日が入っている場合は対象外
  if (statusStr === SANTE_CONFIG.REPLY_STATUS.COMPLETED && replyDateStr !== '') {
    return false;
  }

  return statusStr === SANTE_CONFIG.REPLY_STATUS.UNCREATED || replyDateStr === '';
}

/**
 * 18行目（🎯 返書消込ターゲットID）の表示文字列を生成する
 * フォーマット: `kintone#<レコードID> / 台帳#<行番号> (<受診日>受診分)`
 * @param {string|number} kintoneRecordId
 * @param {number} ledgerRowNumber
 * @param {string} visitDate
 * @return {string}
 */
function formatTargetRecordInfo(kintoneRecordId, ledgerRowNumber, visitDate) {
  const recId = String(kintoneRecordId || '').trim();
  const dateStr = formatDateYMD(visitDate) || '日付不明';
  return `kintone#${recId} / 台帳#${ledgerRowNumber} (${dateStr}受診分)`;
}

/**
 * 18行目（🎯 返書消込ターゲットID）の文字列をパースし、消込に必要な識別子を抽出する
 * @param {string} targetInfoStr
 * @return {{
 *   kintoneRecordId: string,
 *   ledgerRowNumber: number,
 *   visitDate: string,
 *   fiscalYear: number
 * }|null}
 */
function parseTargetRecordInfo(targetInfoStr) {
  const str = String(targetInfoStr || '').trim();
  if (!str) return null;

  const kintoneMatch = str.match(/kintone#(\d+)/i);
  const ledgerMatch = str.match(/台帳#(\d+)/);
  const dateMatch = str.match(/\((\d{4}-\d{2}-\d{2})受診分\)/);

  if (!kintoneMatch) {
    return null;
  }

  const visitDate = dateMatch ? dateMatch[1] : '';
  const fiscalYear = visitDate ? getFiscalYear(visitDate) : getFiscalYear();

  return {
    kintoneRecordId: kintoneMatch[1],
    ledgerRowNumber: ledgerMatch ? parseInt(ledgerMatch[1], 10) : 0,
    visitDate: visitDate,
    fiscalYear: fiscalYear,
  };
}

/**
 * 【要件A: 5重ロック自動照合】
 * 連携フェーズが「返書（経過・結果報告）」または「後方（逆紹介・送出）」のデータについて、
 * 年度別台帳（当年度＋必要に応じて前年度）から5重ロック条件をすべて満たす最新受診日の行を検索する
 *
 * @param {{
 *   visitDate: string,
 *   patientId: string,
 *   patientName: string,
 *   phase: string,
 *   facilityName: string,
 *   facilityDoctor: string,
 *   lookupKey?: string
 * }} inputRecord
 * @param {Array<Object>=} facilityMasterList
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ss
 * @return {{
 *   matched: boolean,
 *   kintoneRecordId: string,
 *   ledgerRowNumber: number,
 *   ledgerSheetName: string,
 *   visitDate: string,
 *   formattedTargetInfo: string
 * }}
 */
function findFiveLockTarget(inputRecord, facilityMasterList, ss) {
  const emptyResult = {
    matched: false,
    kintoneRecordId: '',
    ledgerRowNumber: 0,
    ledgerSheetName: '',
    visitDate: '',
    formattedTargetInfo: '',
  };

  const phase = String(inputRecord.phase || '').trim();
  const isTargetPhase =
    phase === SANTE_CONFIG.PHASES.REPLY ||
    phase === SANTE_CONFIG.PHASES.BACKWARD ||
    phase.startsWith('返書') ||
    phase.startsWith('後方');

  if (!isTargetPhase) {
    return emptyResult;
  }

  const masterList = facilityMasterList || loadFacilityMasterList(ss);
  const resolvedInputFacility = resolveFacility(inputRecord.facilityName, masterList);
  if (inputRecord.lookupKey && !resolvedInputFacility.lookupKey) {
    resolvedInputFacility.lookupKey = inputRecord.lookupKey;
  }

  const candidateSheets = getCandidateLedgerSheets(inputRecord.visitDate, ss);
  const idx = SANTE_CONFIG.LEDGER_COLUMNS.INDEX;
  const matchedCandidates = [];

  candidateSheets.forEach(({ sheet, sheetName }) => {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const values = sheet
      .getRange(2, 1, lastRow - 1, SANTE_CONFIG.LEDGER_COLUMNS.HEADERS.length)
      .getValues();

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const sheetRowNumber = i + 2; // 2行目開始

      const ledgerKintoneId = String(row[idx.KINTONE_RECORD_ID] || '').trim();
      const ledgerVisitDate = formatDateYMD(row[idx.VISIT_DATE]);
      const ledgerPatientId = String(row[idx.PATIENT_ID] || '').trim();
      const ledgerPatientName = String(row[idx.PATIENT_NAME] || '').trim();
      const ledgerPhase = String(row[idx.PHASE] || '').trim();
      const ledgerLookupKey = String(row[idx.LOOKUP_KEY] || '').trim();
      const ledgerFacilityName = String(row[idx.FACILITY_NAME] || '').trim();
      const ledgerFacilityDoctor = String(row[idx.FACILITY_DOCTOR] || '').trim();
      const ledgerReplyStatus = String(row[idx.REPLY_STATUS] || '').trim();
      const ledgerReplyDate = row[idx.REPLY_DATE];

      // ロック4（未返書状態 ＆ 前方紹介レコード）を先に判定して高速フィルタリング
      if (!isLock4UnrepliedForward(ledgerPhase, ledgerReplyStatus, ledgerReplyDate)) {
        continue;
      }

      // ロック1（患者特定: 患者ID一致 または スペース除去後の患者氏名一致）
      if (
        !isLock1PatientMatched(
          inputRecord.patientId,
          inputRecord.patientName,
          ledgerPatientId,
          ledgerPatientName
        )
      ) {
        continue;
      }

      // ロック2（施設特定: 名寄せ後の連携先施設名が一致または部分一致）
      if (
        !isLock2FacilityMatched(
          resolvedInputFacility,
          ledgerLookupKey,
          ledgerFacilityName,
          masterList
        )
      ) {
        continue;
      }

      // ロック3（医師名の「含む」部分一致・連名対応）
      if (!isLock3DoctorMatched(inputRecord.facilityDoctor, ledgerFacilityDoctor)) {
        continue;
      }

      matchedCandidates.push({
        kintoneRecordId: ledgerKintoneId,
        ledgerRowNumber: sheetRowNumber,
        ledgerSheetName: sheetName,
        visitDate: ledgerVisitDate,
      });
    }
  });

  if (matchedCandidates.length === 0) {
    return emptyResult;
  }

  // ロック5（最新優先）: 受診日が最も新しい行を優先（同日の場合は行番号が大きい方を優先）
  matchedCandidates.sort((a, b) => {
    if (a.visitDate !== b.visitDate) {
      return b.visitDate.localeCompare(a.visitDate);
    }
    return b.ledgerRowNumber - a.ledgerRowNumber;
  });

  const best = matchedCandidates[0];
  return {
    matched: true,
    kintoneRecordId: best.kintoneRecordId,
    ledgerRowNumber: best.ledgerRowNumber,
    ledgerSheetName: best.ledgerSheetName,
    visitDate: best.visitDate,
    formattedTargetInfo: formatTargetRecordInfo(
      best.kintoneRecordId,
      best.ledgerRowNumber,
      best.visitDate
    ),
  };
}

/**
 * 年度別台帳タブの最下行へ新規登録レコード群を一括追記する
 * @param {Array<{
 *   kintoneRecordId: string,
 *   registeredAt: string,
 *   visitDate: string,
 *   patientId: string,
 *   patientName: string,
 *   phase: string,
 *   visitType: string,
 *   lookupKey: string,
 *   facilityName: string,
 *   facilityDept: string,
 *   facilityDoctor: string,
 *   ourDept: string,
 *   ourDoctor: string,
 *   replyStatus: string,
 *   replyDate: string,
 *   firstReportUrl: string,
 *   faxStatus: string,
 *   remarks: string
 * }>} recordsToAppend
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ss
 * @return {Array<{sheetName: string, rowNumber: number, kintoneRecordId: string}>}
 */
function appendRecordsToLedger(recordsToAppend, ss) {
  if (!recordsToAppend || recordsToAppend.length === 0) {
    return [];
  }

  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  // 会計年度ごとにグループ化して一括追記
  const groupedByFy = {};
  recordsToAppend.forEach((rec) => {
    const fy = getFiscalYear(rec.visitDate);
    if (!groupedByFy[fy]) {
      groupedByFy[fy] = [];
    }
    groupedByFy[fy].push(rec);
  });

  const appendedMeta = [];

  Object.keys(groupedByFy).forEach((fyKey) => {
    const fy = parseInt(fyKey, 10);
    const sheet = getOrCreateLedgerSheet(fy, spreadsheet);
    const sheetName = sheet.getName();
    const startRow = sheet.getLastRow() + 1;
    const group = groupedByFy[fyKey];

    const rowValues = group.map((rec, idx) => {
      appendedMeta.push({
        sheetName: sheetName,
        rowNumber: startRow + idx,
        kintoneRecordId: rec.kintoneRecordId,
      });

      return [
        rec.kintoneRecordId,
        rec.registeredAt || formatDateTimeNow(),
        rec.visitDate,
        rec.patientId,
        rec.patientName,
        rec.phase,
        rec.visitType,
        rec.lookupKey,
        rec.facilityName,
        rec.facilityDept,
        rec.facilityDoctor,
        rec.ourDept,
        rec.ourDoctor,
        rec.replyStatus,
        rec.replyDate,
        rec.firstReportUrl,
        rec.faxStatus,
        rec.remarks,
      ];
    });

    sheet
      .getRange(startRow, 1, rowValues.length, SANTE_CONFIG.LEDGER_COLUMNS.HEADERS.length)
      .setValues(rowValues);
  });

  return appendedMeta;
}

/**
 * 年度別台帳の指定行（またはkintoneレコードID一致行）の返書ステータス・返書日を更新する
 * @param {Array<{
 *   kintoneRecordId: string,
 *   ledgerRowNumber: number,
 *   fiscalYear: number,
 *   replyDate: string
 * }>} updates
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ss
 * @return {number} 更新成功行数
 */
function updateLedgerReplyRecords(updates, ss) {
  if (!updates || updates.length === 0) return 0;

  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  const colNum = SANTE_CONFIG.LEDGER_COLUMNS.COL_NUM;
  const idx = SANTE_CONFIG.LEDGER_COLUMNS.INDEX;
  let updatedCount = 0;

  updates.forEach((upd) => {
    const sheet = getOrCreateLedgerSheet(upd.fiscalYear, spreadsheet);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    let targetRow = upd.ledgerRowNumber;

    // 指定行番号のkintoneレコードIDが一致するか検証し、ズレていればkintoneレコードIDで再検索
    if (targetRow >= 2 && targetRow <= lastRow) {
      const idAtRow = String(
        sheet.getRange(targetRow, colNum.KINTONE_RECORD_ID).getValue() || ''
      ).trim();
      if (idAtRow !== String(upd.kintoneRecordId).trim()) {
        targetRow = 0;
      }
    } else {
      targetRow = 0;
    }

    if (targetRow === 0 && upd.kintoneRecordId) {
      const idValues = sheet
        .getRange(2, colNum.KINTONE_RECORD_ID, lastRow - 1, 1)
        .getValues();
      for (let i = 0; i < idValues.length; i++) {
        if (String(idValues[i][0]).trim() === String(upd.kintoneRecordId).trim()) {
          targetRow = i + 2;
          break;
        }
      }
    }

    if (targetRow >= 2) {
      // N列(14): 返書ステータス = "作成・送付完了", O列(15): 返書日 = 今回の日付
      sheet
        .getRange(targetRow, colNum.REPLY_STATUS, 1, 2)
        .setValues([[SANTE_CONFIG.REPLY_STATUS.COMPLETED, upd.replyDate]]);
      updatedCount++;
    } else {
      Logger.log(
        `警告: 台帳上で消込対象レコードが見つかりませんでした (kintone#${upd.kintoneRecordId}, 行#${upd.ledgerRowNumber})`
      );
    }
  });

  return updatedCount;
}
