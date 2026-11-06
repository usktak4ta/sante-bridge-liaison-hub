/**
 * @fileoverview kintone REST API クライアントモジュール（アプリB / App ID: 195 連携・Phase 2 強化版）
 * - ルックアップ参照元（施設マスター）APIトークンのカンマ連結対応
 * - 連携方向（DIRECTION）フィールド互換対応
 * - 未定義フィールドエラー（GAIA_FC01等）検知時の自動除外リトライ（Auto-Recovery）機能搭載
 */

/**
 * kintone API 通信設定を検証し、エンドポイントURLとヘッダーを構築する
 * ルックアップフィールド（更新用キー）参照用のマスタートークンが設定されている場合はカンマ連結する
 *
 * @return {{baseUrl: string, appId: string, headers: Object<string, string>}}
 */
function getKintoneRequestContext() {
  const secret = getSecretConfig();
  if (!secret.kintoneDomain) {
    throw new Error('スクリプトプロパティ「KINTONE_DOMAIN」が設定されていません。');
  }
  if (!secret.kintoneApiToken) {
    throw new Error('スクリプトプロパティ「KINTONE_API_TOKEN」が設定されていません。');
  }

  // ドメイン文字列に https:// が含まれている場合も安全に正規化
  const cleanDomain = secret.kintoneDomain
    .replace(/^https?:\/\//i, '')
    .replace(/\/+$/, '');

  // アプリBのトークンと施設マスター参照トークンをカンマ区切りで連結（重複排除）
  const rawTokens = [secret.kintoneApiToken, secret.kintoneApiTokenMaster]
    .filter(Boolean)
    .join(',')
    .split(',')
    .map((t) => t.trim())
    .filter(Boolean);
  const combinedToken = Array.from(new Set(rawTokens)).join(',');

  return {
    baseUrl: `https://${cleanDomain}/k/v1/records.json`,
    appId: secret.kintoneAppId || SANTE_CONFIG.KINTONE.DEFAULT_APP_ID,
    headers: {
      'X-Cybozu-API-Token': combinedToken,
    },
  };
}

/**
 * 診療科文字列を kintone アプリBの「紹介先診療科」5択
 * （脳神経内科外科 / 整形外科 / 内科 / 血管外科 / リハビリテーション科）へ正規化する。
 * 5択に正規化できない場合は空文字を返し、選択肢不正エラー（GAIA_FO01）を防止する。
 *
 * @param {string} rawDept
 * @return {string} 5択のいずれか、または空文字 ''
 */
function normalizeReferralDeptOption(rawDept) {
  const s = normalizeCompactString(rawDept);
  if (!s) return '';

  const validOptions = SANTE_CONFIG.KINTONE.REFERRAL_DEPT_OPTIONS || [
    '脳神経内科外科',
    '整形外科',
    '内科',
    '血管外科',
    'リハビリテーション科',
  ];

  if (validOptions.includes(s)) {
    return s;
  }

  if (
    s === '脳神経内科' ||
    s === '脳神経外科' ||
    s === '脳神経内科・外科' ||
    s === '脳神経外科・内科' ||
    s === '脳外' ||
    s === '脳内'
  ) {
    return '脳神経内科外科';
  }
  if (s === '整形') {
    return '整形外科';
  }
  if (s === 'リハ科' || s === 'リハビリ科' || s === 'リハビリテーション') {
    return 'リハビリテーション科';
  }

  return '';
}

/**
 * 当院担当診療科および担当医師名を正規化する（現役医師11名・9診療科対応）
 * 退職医（野田明生・佐藤恭一）や選択肢外の医師・診療科はドロップダウン値から除外し、
 * 備考欄への退避テキスト（extraNote）として返すことで GAIA_FO01 エラーを防止する
 *
 * @param {string} rawDept
 * @param {string} rawDoc
 * @param {string=} inquiryCol
 * @param {string=} remarksCol
 * @return {{dept: string, doctor: string, extraNote: string}}
 */
function normalizeOurDeptAndDoctor(rawDept, rawDoc, inquiryCol, remarksCol) {
  let doc = String(rawDoc || '').replace(/[\s\u3000]+/g, '').replace(/(先生|医師|Dr\.?)$/i, '');
  let dept = String(rawDept || '').trim();
  let extraNote = '';
  // よくある誤字補正
  if (doc === '大江陽介' || doc === '大江健一') doc = '大江健介';
  if (doc === '田中準') doc = '田中淳';
  if (doc === '佐々木威浩' || (!doc && String(inquiryCol || '').includes('佐々木'))) doc = '佐々木威治';
  // 医師名から診療科の確定または補完（現役11名）
  if (['溝上泰一朗', '尤郁偉', '劉軒'].includes(doc)) dept = '脳神経外科';
  else if (['上床武史'].includes(doc)) dept = '脳血管内科';
  else if (['田中淳'].includes(doc)) dept = '脳神経内科';
  else if (['奈田慎一'].includes(doc)) dept = '血管外科';
  else if (['吉原智仁', '泉政寛', '佐々木威治'].includes(doc)) dept = '整形外科';
  else if (['大江健介'].includes(doc)) dept = dept.includes('循環器') ? '循環器内科' : '内科（一般）';
  else if (['村山賢一郎'].includes(doc)) dept = dept.includes('消化器') ? '消化器内科' : '内科（一般）';
  // 退職医（野田明生・佐藤恭一）は診療科だけ補完し、医師名は後段で備考へ退避
  else if (doc === '野田明生') dept = '整形外科';
  else if (doc === '佐藤恭一') dept = dept.includes('循環器') ? '循環器内科' : '内科（一般）';
  else {
    if (dept.includes('脳外')) dept = '脳神経外科';
    else if (dept.includes('脳内')) dept = '脳神経内科';
    else if (dept.includes('血管内科')) dept = '脳血管内科';
    else if (dept.includes('血管外科') || dept.includes('血外')) dept = '血管外科';
    else if (dept.includes('整形') || String(inquiryCol || '').includes('整形')) dept = '整形外科';
    else if (dept.includes('循環器')) dept = '循環器内科';
    else if (dept.includes('消化器') || (dept === '外科' && String(remarksCol || '').includes('胃カメラ'))) dept = '消化器内科';
    else if (dept.includes('リハ')) dept = 'リハビリテーション科';
    else if (dept.includes('内科') || dept === '外科') dept = '内科（一般）';
  }
  const validDepts = SANTE_CONFIG.KINTONE.OUR_DEPT_OPTIONS;
  const validDoctors = SANTE_CONFIG.KINTONE.OUR_DOCTOR_OPTIONS;
  if (dept && !validDepts.includes(dept)) {
    extraNote += `当院科:${dept} `;
    dept = '';
  }
  if (doc && !validDoctors.includes(doc)) {
    extraNote += `旧/他担当医:${doc}`;
    doc = '';
  }
  return { dept: dept, doctor: doc, extraNote: extraNote.trim() };
}

/**
 * 入力シートの1列分のデータから kintone POST 用の record オブジェクトを構築する
 * @param {{
 *   visitDate: string,
 *   patientId: string,
 *   patientName: string,
 *   phase: string,
 *   visitType: string,
 *   lookupKey: string,
 *   facilityName: string,
 *   facilityDept: string,
 *   referralDept?: string,
 *   facilityDoctor: string,
 *   ourDept: string,
 *   ourDoctor: string,
 *   inquiry?: string,
 *   replyStatus: string,
 *   replyDate: string,
 *   firstReportUrl: string,
 *   faxStatus: string,
 *   remarks: string
 * }} item
 * @param {Set<string>=} excludedFieldsSet Auto-Recoveryで除外された未定義フィールドコードのSet
 * @return {Object} kintone record ペイロード
 */
function buildKintonePostRecordPayload(item, excludedFieldsSet) {
  const fc = SANTE_CONFIG.KINTONE.FIELD_CODES;
  const excluded = excludedFieldsSet || new Set();
  const record = {};

  const setField = (fieldCode, value) => {
    if (!fieldCode || excluded.has(fieldCode)) return;
    record[fieldCode] = { value: value };
  };

  if (item.visitDate) {
    setField(fc.VISIT_DATE, formatDateYMD(item.visitDate));
  }
  setField(fc.PATIENT_ID, String(item.patientId || ''));
  setField(fc.PATIENT_NAME, String(item.patientName || ''));

  // 連携フェーズ および 既存アプリB互換の「連携方向」の両方にセット（未定義片方はAuto-Recoveryで自動除外）
  const phaseVal = String(item.phase || '').trim();
  if (phaseVal) {
    setField(fc.PHASE, phaseVal);
  }
  if (
    phaseVal === SANTE_CONFIG.PHASES.FORWARD ||
    phaseVal === SANTE_CONFIG.PHASES.BACKWARD ||
    phaseVal.startsWith('前方') ||
    phaseVal.startsWith('後方')
  ) {
    setField(fc.DIRECTION, phaseVal);
  }

  const visitTypeVal = String(item.visitType || '').trim();
  if (visitTypeVal) {
    setField(fc.VISIT_TYPE, visitTypeVal);
  }

  // ルックアップキーが存在する場合は更新用キーのみをセットし、施設名はルックアップの自動コピーに任せる
  // ルックアップキーが空のときだけ自由入力として施設名をセットする
  const lookupKeyStr = String(item.lookupKey || '').trim();
  if (lookupKeyStr) {
    setField(fc.LOOKUP_KEY, lookupKeyStr);
  } else if (item.facilityName) {
    setField(fc.FACILITY_NAME, String(item.facilityName));
  }

  const facilityDeptStr = String(item.facilityDept || '');
  setField(fc.FACILITY_DEPT, facilityDeptStr);
  const normalizedReferralDept = normalizeReferralDeptOption(
    item.referralDept !== undefined ? item.referralDept : facilityDeptStr
  );
  if (normalizedReferralDept) {
    setField(fc.REFERRAL_DEPT, normalizedReferralDept);
  }

  setField(fc.FACILITY_DOCTOR, String(item.facilityDoctor || ''));

  // 当院担当診療科・担当医師の11名正規化（空のドロップダウン値は送信せず、退職医・選択肢外は備考へ退避）
  const rawRemarks = String(item.remarks || '').trim();
  const normalizedOur = normalizeOurDeptAndDoctor(
    item.ourDept,
    item.ourDoctor,
    item.inquiry || '',
    rawRemarks
  );
  if (normalizedOur.dept) {
    setField(fc.OUR_DEPT, normalizedOur.dept);
  }
  if (normalizedOur.doctor) {
    setField(fc.OUR_DOCTOR, normalizedOur.doctor);
  }

  const replyStatusVal = String(item.replyStatus || '').trim();
  if (replyStatusVal) {
    setField(fc.REPLY_STATUS, replyStatusVal);
  }

  if (item.replyDate) {
    setField(fc.REPLY_DATE, formatDateYMD(item.replyDate));
  }
  setField(fc.FIRST_REPORT_URL, String(item.firstReportUrl || ''));
  const faxStatusVal = String(item.faxStatus || '').trim();
  if (faxStatusVal) {
    setField(fc.FAX_STATUS, faxStatusVal);
  }

  let finalRemarks = rawRemarks;
  if (normalizedOur.extraNote && !finalRemarks.includes(normalizedOur.extraNote)) {
    finalRemarks = finalRemarks
      ? `${finalRemarks} / ${normalizedOur.extraNote}`
      : normalizedOur.extraNote;
  }
  setField(fc.REMARKS, finalRemarks);

  return record;
}

/**
 * kintone APIエラーレスポンス（GAIA_FC01等）を解析し、
 * アプリB側に存在しない（またはルックアップコピー先等で直接書き込み不可の）フィールドコードを抽出する
 *
 * @param {string} responseText kintoneエラーレスポンスJSON文字列
 * @param {Array<string>} candidateFieldCodes 現在送信中のフィールドコード一覧
 * @param {Set<string>} alreadyExcluded 既に除外済みのフィールドコードSet
 * @return {Array<string>} 新たに除外すべきフィールドコード配列
 */
function extractMissingFieldCodesFromError(responseText, candidateFieldCodes, alreadyExcluded) {
  const detected = new Set();
  if (!responseText) return [];

  let parsed = null;
  try {
    parsed = JSON.parse(responseText);
  } catch (e) {
    // JSON以外のレスポンスの場合は空配列
    return [];
  }

  const code = String(parsed.code || '');
  const message = String(parsed.message || '');
  const isFieldErrorCode =
    code === 'GAIA_FC01' ||
    code === 'CB_VA01' ||
    message.includes('フィールド') ||
    message.toLowerCase().includes('field');

  if (!isFieldErrorCode) {
    return [];
  }

  // 1. エラーメッセージ内の括弧 `（フィールド名）` または `(fieldCode)` から抽出
  const bracketRegex = /[（(]([^）)]+)[）)]/g;
  let match;
  while ((match = bracketRegex.exec(message)) !== null) {
    const candidate = match[1].trim();
    if (candidate && !alreadyExcluded.has(candidate) && candidateFieldCodes.includes(candidate)) {
      detected.add(candidate);
    }
  }

  // 2. parsed.errors オブジェクトのキー（例: "records[0].第一報URL.value" や "record.返書ステータス.value"）から抽出
  if (parsed.errors && typeof parsed.errors === 'object') {
    Object.keys(parsed.errors).forEach((errKey) => {
      const errDetail = parsed.errors[errKey];
      const errMsgs = (errDetail && Array.isArray(errDetail.messages) ? errDetail.messages : []).join(' ');
      const isRecoverableFieldErr =
        code === 'GAIA_FC01' ||
        errMsgs.includes('見つかりません') ||
        errMsgs.includes('存在しません') ||
        errMsgs.includes('指定できません') ||
        errMsgs.toLowerCase().includes('not found');

      const keyMatch = errKey.match(/^records?(?:\[\d+\])?\.([^.]+)(?:\.value)?$/);
      if (keyMatch && isRecoverableFieldErr) {
        const fcName = keyMatch[1].trim();
        if (fcName && !alreadyExcluded.has(fcName)) {
          detected.add(fcName);
        }
      }
    });
  }

  // 3. GAIA_FC01 で括弧抽出できなかった場合のフォールバック: 拡張フィールド群がレスポンス本文に含まれているか確認
  if (detected.size === 0 && code === 'GAIA_FC01') {
    candidateFieldCodes.forEach((fcName) => {
      if (!alreadyExcluded.has(fcName) && responseText.includes(fcName)) {
        detected.add(fcName);
      }
    });
  }

  return Array.from(detected);
}

/**
 * Auto-Recovery機能付き kintone `/k/v1/records.json` リクエスト送信共通ヘルパー
 * 未定義フィールドエラー（GAIA_FC01等）発生時、該当フィールドを自動除外してリトライ送信する
 *
 * @param {'post'|'put'} method
 * @param {{baseUrl: string, appId: string, headers: Object<string, string>}} ctx
 * @param {function(Set<string>): Array<Object>} buildRecordsFn excludedFieldsSetを受け取りrecords配列を返す関数
 * @param {Set<string>} sharedExcludedFields チャンク間で共有する除外フィールドSet
 * @param {string} errorContextTitle Slackエラー通知用タイトル
 * @param {Object} errorMeta Slackエラー通知用メタデータ
 * @return {Object} パース済みのkintoneレスポンスオブジェクト
 */
function sendKintoneRequestWithAutoRecovery(
  method,
  ctx,
  buildRecordsFn,
  sharedExcludedFields,
  errorContextTitle,
  errorMeta
) {
  const maxRetries = SANTE_CONFIG.KINTONE.MAX_AUTO_RECOVERY_RETRIES || 8;
  const allCandidateCodes = Object.values(SANTE_CONFIG.KINTONE.FIELD_CODES);

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    const recordsPayload = buildRecordsFn(sharedExcludedFields);
    const body = {
      app: Number(ctx.appId) || ctx.appId,
      records: recordsPayload,
    };

    let response;
    try {
      response = UrlFetchApp.fetch(ctx.baseUrl, {
        method: method,
        contentType: 'application/json',
        headers: ctx.headers,
        payload: JSON.stringify(body),
        muteHttpExceptions: true,
      });
    } catch (networkErr) {
      notifySlackError(
        `${errorContextTitle} ネットワーク通信失敗`,
        'kintone API通信中にネットワークエラーが発生しました',
        errorMeta,
        { includeStack: false }
      );
      throw new Error(`${errorContextTitle} 通信中にネットワークエラーが発生しました`);
    }

    const statusCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (statusCode >= 200 && statusCode < 300) {
      return JSON.parse(responseText);
    }

    // エラー発生時: 未定義フィールドによるエラーか判定し、該当フィールドを除外してリトライ
    const newlyDetectedFields = extractMissingFieldCodesFromError(
      responseText,
      allCandidateCodes,
      sharedExcludedFields
    );

    if (newlyDetectedFields.length > 0 && attempt < maxRetries) {
      newlyDetectedFields.forEach((fc) => sharedExcludedFields.add(fc));
      Logger.log(
        `[kintone Auto-Recovery] 未定義フィールドを検知したためペイロードから除外してリトライします (${attempt + 1}/${maxRetries}): ${newlyDetectedFields.join(', ')}`
      );
      continue;
    }

    let errorCode = 'UNKNOWN';
    const errorFieldCodes = [];
    try {
      const parsedErr = JSON.parse(responseText);
      if (parsedErr && parsedErr.code) {
        errorCode = String(parsedErr.code);
      }
      if (parsedErr && parsedErr.errors && typeof parsedErr.errors === 'object') {
        Object.keys(parsedErr.errors).forEach((errKey) => {
          const keyMatch = errKey.match(/^records?(?:\[\d+\])?\.([^.]+)(?:\.value)?$/);
          const fieldName = keyMatch ? keyMatch[1].trim() : errKey.trim();
          if (fieldName && !errorFieldCodes.includes(fieldName)) {
            errorFieldCodes.push(fieldName);
          }
        });
      }
    } catch (_) {
      // レスポンス本文やレコード値はSlackや例外メッセージへ出さない
    }

    const safeErrorSummary =
      errorFieldCodes.length > 0
        ? `HTTP ${statusCode} / code: ${errorCode} / fields: ${errorFieldCodes.join(', ')}`
        : `HTTP ${statusCode} / code: ${errorCode}`;

    notifySlackError(
      `${errorContextTitle} APIエラー (HTTP ${statusCode})`,
      safeErrorSummary,
      {
        ...errorMeta,
        httpStatus: String(statusCode),
        errorCode: errorCode,
        errorFields: errorFieldCodes.join(', ') || 'なし',
        excludedFields: Array.from(sharedExcludedFields).join(', ') || 'なし',
      },
      { includeStack: false }
    );
    throw new Error(`${errorContextTitle} APIエラー (HTTP ${statusCode}, code: ${errorCode})`);
  }

  throw new Error(`${errorContextTitle}: Auto-Recoveryのリトライ上限(${maxRetries}回)を超過しました。`);
}

/**
 * 【要件B-2】kintone アプリB（ID: 195）へ複数レコードを一括POST（新規登録）する
 * 100件を超える場合も自動でチャンク分割して送信し、未定義フィールドがあれば自動除外リトライした上で
 * 発番された全レコードID配列を返す
 *
 * @param {Array<Object>} recordsData 登録対象のレコードオブジェクト配列
 * @return {{ids: Array<string>, revisions: Array<string>, excludedFields: Array<string>}}
 */
function postRecordsToKintoneBatch(recordsData) {
  if (!recordsData || recordsData.length === 0) {
    return { ids: [], revisions: [], excludedFields: [] };
  }

  const ctx = getKintoneRequestContext();
  const batchLimit = SANTE_CONFIG.KINTONE.BATCH_LIMIT;
  const allIds = [];
  const allRevisions = [];
  const sharedExcludedFields = new Set();

  for (let offset = 0; offset < recordsData.length; offset += batchLimit) {
    const chunk = recordsData.slice(offset, offset + batchLimit);

    const parsed = sendKintoneRequestWithAutoRecovery(
      'post',
      ctx,
      (excludedSet) => chunk.map((item) => buildKintonePostRecordPayload(item, excludedSet)),
      sharedExcludedFields,
      'kintone 一括POST',
      {
        appId: ctx.appId,
        chunkSize: chunk.length,
      }
    );

    if (Array.isArray(parsed.ids)) {
      parsed.ids.forEach((id) => allIds.push(String(id)));
    }
    if (Array.isArray(parsed.revisions)) {
      parsed.revisions.forEach((rev) => allRevisions.push(String(rev)));
    }
  }

  return {
    ids: allIds,
    revisions: allRevisions,
    excludedFields: Array.from(sharedExcludedFields),
  };
}

/**
 * 【要件B-3】kintone アプリB（ID: 195）の既存レコードに対して「返書日」および「返書ステータス」を一括PUT更新する
 * アプリB側に「返書ステータス」フィールドが存在しない場合も Auto-Recovery により「返書日」のみで安全に更新を完遂する
 *
 * @param {Array<{
 *   kintoneRecordId: string,
 *   replyDate: string,
 *   replyStatus?: string
 * }>} putTargets
 * @return {{updatedIds: Array<string>, excludedFields: Array<string>}}
 */
function putReplyDatesToKintoneBatch(putTargets) {
  if (!putTargets || putTargets.length === 0) {
    return { updatedIds: [], excludedFields: [] };
  }

  const ctx = getKintoneRequestContext();
  const fc = SANTE_CONFIG.KINTONE.FIELD_CODES;
  const batchLimit = SANTE_CONFIG.KINTONE.BATCH_LIMIT;

  // 同一バッチ内で同じ kintoneRecordId が重複すると kintone PUT API がエラーになるため重複排除
  const dedupMap = new Map();
  putTargets.forEach((t) => {
    const idStr = String(t.kintoneRecordId || '').trim();
    if (!idStr) return;
    dedupMap.set(idStr, {
      kintoneRecordId: idStr,
      replyDate: formatDateYMD(t.replyDate) || formatDateYMD(new Date()),
      replyStatus: t.replyStatus || SANTE_CONFIG.REPLY_STATUS.COMPLETED,
    });
  });

  const uniqueTargets = Array.from(dedupMap.values());
  if (uniqueTargets.length === 0) {
    return { updatedIds: [], excludedFields: [] };
  }

  const updatedIds = [];
  const sharedExcludedFields = new Set();

  for (let offset = 0; offset < uniqueTargets.length; offset += batchLimit) {
    const chunk = uniqueTargets.slice(offset, offset + batchLimit);

    const parsed = sendKintoneRequestWithAutoRecovery(
      'put',
      ctx,
      (excludedSet) =>
        chunk.map((target) => {
          const recObj = {};
          if (!excludedSet.has(fc.REPLY_DATE)) {
            recObj[fc.REPLY_DATE] = { value: target.replyDate };
          }
          if (!excludedSet.has(fc.REPLY_STATUS)) {
            recObj[fc.REPLY_STATUS] = { value: target.replyStatus };
          }
          return {
            id: Number(target.kintoneRecordId) || target.kintoneRecordId,
            record: recObj,
          };
        }),
      sharedExcludedFields,
      'kintone 一括PUT（返書消込）',
      {
        appId: ctx.appId,
        targetIds: chunk.map((c) => c.kintoneRecordId).join(', '),
      }
    );

    if (Array.isArray(parsed.records)) {
      parsed.records.forEach((r) => updatedIds.push(String(r.id)));
    } else {
      chunk.forEach((c) => updatedIds.push(c.kintoneRecordId));
    }
  }

  return {
    updatedIds,
    excludedFields: Array.from(sharedExcludedFields),
  };
}
