/**
 * @fileoverview kintone REST API クライアントモジュール（アプリB / App ID: 195 連携）
 * サンテ溝上病院 地域連携・紹介返書管理＆第一報自動化システム (SANTE-BRIDGE)
 */

/**
 * kintone API 通信設定を検証し、エンドポイントURLとヘッダーを構築する
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

  return {
    baseUrl: `https://${cleanDomain}/k/v1/records.json`,
    appId: secret.kintoneAppId || SANTE_CONFIG.KINTONE.DEFAULT_APP_ID,
    headers: {
      'X-Cybozu-API-Token': secret.kintoneApiToken,
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
 * @return {Object} kintone record ペイロード
 */
function buildKintonePostRecordPayload(item) {
  const fc = SANTE_CONFIG.KINTONE.FIELD_CODES;
  const record = {};

  if (item.visitDate) {
    record[fc.VISIT_DATE] = { value: formatDateYMD(item.visitDate) };
  }
  record[fc.PATIENT_ID] = { value: String(item.patientId || '') };
  record[fc.PATIENT_NAME] = { value: String(item.patientName || '') };
  record[fc.PHASE] = { value: String(item.phase || '') };
  record[fc.VISIT_TYPE] = { value: String(item.visitType || '') };

  // ルックアップキーが存在する場合は更新用キーをセット、あわせて施設名もセット
  if (item.lookupKey) {
    record[fc.LOOKUP_KEY] = { value: String(item.lookupKey) };
  }
  if (item.facilityName && !item.lookupKey) {
    record[fc.FACILITY_NAME] = { value: String(item.facilityName) };
  } else if (item.facilityName && fc.FACILITY_NAME !== fc.LOOKUP_KEY) {
    // ルックアップ自動コピー対象でない構成にも対応できるよう施設名を保持（ルックアップコピー先の場合はkintone側設定に従う）
    record[fc.FACILITY_NAME] = { value: String(item.facilityName) };
  }

  record[fc.FACILITY_DEPT] = { value: String(item.facilityDept || '') };
  record[fc.FACILITY_DOCTOR] = { value: String(item.facilityDoctor || '') };
  record[fc.OUR_DEPT] = { value: String(item.ourDept || '') };
  record[fc.OUR_DOCTOR] = { value: String(item.ourDoctor || '') };
  record[fc.REPLY_STATUS] = { value: String(item.replyStatus || '') };

  if (item.replyDate) {
    record[fc.REPLY_DATE] = { value: formatDateYMD(item.replyDate) };
  }
  record[fc.FIRST_REPORT_URL] = { value: String(item.firstReportUrl || '') };
  record[fc.FAX_STATUS] = { value: String(item.faxStatus || '') };
  record[fc.REMARKS] = { value: String(item.remarks || '') };

  return record;
}

/**
 * 【要件B-2】kintone アプリB（ID: 195）へ複数レコードを一括POST（新規登録）する
 * 100件を超える場合も自動でチャンク分割して送信し、発番された全レコードID配列を返す
 *
 * @param {Array<Object>} recordsData 登録対象のレコードオブジェクト配列
 * @return {{ids: Array<string>, revisions: Array<string>}}
 */
function postRecordsToKintoneBatch(recordsData) {
  if (!recordsData || recordsData.length === 0) {
    return { ids: [], revisions: [] };
  }

  const ctx = getKintoneRequestContext();
  const batchLimit = SANTE_CONFIG.KINTONE.BATCH_LIMIT;
  const allIds = [];
  const allRevisions = [];

  for (let offset = 0; offset < recordsData.length; offset += batchLimit) {
    const chunk = recordsData.slice(offset, offset + batchLimit);
    const kintoneRecords = chunk.map((item) => buildKintonePostRecordPayload(item));

    const body = {
      app: Number(ctx.appId) || ctx.appId,
      records: kintoneRecords,
    };

    let response;
    try {
      response = UrlFetchApp.fetch(ctx.baseUrl, {
        method: 'post',
        contentType: 'application/json',
        headers: ctx.headers,
        payload: JSON.stringify(body),
        muteHttpExceptions: true,
      });
    } catch (networkErr) {
      notifySlackError('kintone 一括POST ネットワーク通信失敗', networkErr, {
        appId: ctx.appId,
        chunkSize: chunk.length,
      });
      throw new Error(`kintone POST通信中にネットワークエラーが発生しました: ${networkErr.message}`);
    }

    const statusCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (statusCode < 200 || statusCode >= 300) {
      notifySlackError(`kintone 一括POST APIエラー (HTTP ${statusCode})`, responseText, {
        appId: ctx.appId,
        chunkSize: chunk.length,
      });
      throw new Error(`kintone POST APIエラー (HTTP ${statusCode}): ${responseText}`);
    }

    const parsed = JSON.parse(responseText);
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
  };
}

/**
 * 【要件B-3】kintone アプリB（ID: 195）の既存レコードに対して「返書日」および「返書ステータス」を一括PUT更新する
 *
 * @param {Array<{
 *   kintoneRecordId: string,
 *   replyDate: string,
 *   replyStatus?: string
 * }>} putTargets
 * @return {{updatedIds: Array<string>}}
 */
function putReplyDatesToKintoneBatch(putTargets) {
  if (!putTargets || putTargets.length === 0) {
    return { updatedIds: [] };
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
    return { updatedIds: [] };
  }

  const updatedIds = [];

  for (let offset = 0; offset < uniqueTargets.length; offset += batchLimit) {
    const chunk = uniqueTargets.slice(offset, offset + batchLimit);
    const recordsPayload = chunk.map((target) => ({
      id: Number(target.kintoneRecordId) || target.kintoneRecordId,
      record: {
        [fc.REPLY_DATE]: {
          value: target.replyDate,
        },
        [fc.REPLY_STATUS]: {
          value: target.replyStatus,
        },
      },
    }));

    const body = {
      app: Number(ctx.appId) || ctx.appId,
      records: recordsPayload,
    };

    let response;
    try {
      response = UrlFetchApp.fetch(ctx.baseUrl, {
        method: 'put',
        contentType: 'application/json',
        headers: ctx.headers,
        payload: JSON.stringify(body),
        muteHttpExceptions: true,
      });
    } catch (networkErr) {
      notifySlackError('kintone 一括PUT（返書消込）ネットワーク通信失敗', networkErr, {
        appId: ctx.appId,
        targetIds: chunk.map((c) => c.kintoneRecordId).join(', '),
      });
      throw new Error(`kintone PUT通信中にネットワークエラーが発生しました: ${networkErr.message}`);
    }

    const statusCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (statusCode < 200 || statusCode >= 300) {
      notifySlackError(`kintone 一括PUT（返書消込）APIエラー (HTTP ${statusCode})`, responseText, {
        appId: ctx.appId,
        targetIds: chunk.map((c) => c.kintoneRecordId).join(', '),
      });
      throw new Error(`kintone PUT APIエラー (HTTP ${statusCode}): ${responseText}`);
    }

    const parsed = JSON.parse(responseText);
    if (Array.isArray(parsed.records)) {
      parsed.records.forEach((r) => updatedIds.push(String(r.id)));
    } else {
      chunk.forEach((c) => updatedIds.push(c.kintoneRecordId));
    }
  }

  return { updatedIds };
}
