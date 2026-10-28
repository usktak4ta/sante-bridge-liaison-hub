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
 *   facilityDoctor: string,
 *   ourDept: string,
 *   ourDoctor: string,
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
  const phaseVal = String(item.phase || '');
  setField(fc.PHASE, phaseVal);
  if (
    phaseVal === SANTE_CONFIG.PHASES.FORWARD ||
    phaseVal === SANTE_CONFIG.PHASES.BACKWARD ||
    phaseVal.startsWith('前方') ||
    phaseVal.startsWith('後方')
  ) {
    setField(fc.DIRECTION, phaseVal);
  }

  setField(fc.VISIT_TYPE, String(item.visitType || ''));

  // ルックアップキーが存在する場合は更新用キーをセット、あわせて施設名もセット
  if (item.lookupKey) {
    setField(fc.LOOKUP_KEY, String(item.lookupKey));
  }
  if (item.facilityName && !item.lookupKey) {
    setField(fc.FACILITY_NAME, String(item.facilityName));
  } else if (item.facilityName && fc.FACILITY_NAME !== fc.LOOKUP_KEY) {
    setField(fc.FACILITY_NAME, String(item.facilityName));
  }

  setField(fc.FACILITY_DEPT, String(item.facilityDept || ''));
  setField(fc.FACILITY_DOCTOR, String(item.facilityDoctor || ''));
  setField(fc.OUR_DEPT, String(item.ourDept || ''));
  setField(fc.OUR_DOCTOR, String(item.ourDoctor || ''));
  setField(fc.REPLY_STATUS, String(item.replyStatus || ''));

  if (item.replyDate) {
    setField(fc.REPLY_DATE, formatDateYMD(item.replyDate));
  }
  setField(fc.FIRST_REPORT_URL, String(item.firstReportUrl || ''));
  setField(fc.FAX_STATUS, String(item.faxStatus || ''));
  setField(fc.REMARKS, String(item.remarks || ''));

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
      notifySlackError(`${errorContextTitle} ネットワーク通信失敗`, networkErr, errorMeta);
      throw new Error(`${errorContextTitle} 通信中にネットワークエラーが発生しました: ${networkErr.message}`);
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

    notifySlackError(`${errorContextTitle} APIエラー (HTTP ${statusCode})`, responseText, {
      ...errorMeta,
      excludedFields: Array.from(sharedExcludedFields).join(', ') || 'なし',
    });
    throw new Error(`${errorContextTitle} APIエラー (HTTP ${statusCode}): ${responseText}`);
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
