/**
 * @fileoverview 未返書リスト抽出 ＆ LINE WORKS / Slack 定期リマインド通知モジュール (Phase 2)
 * サンテ溝上病院 地域連携・紹介返書管理＆第一報自動化システム (SANTE-BRIDGE)
 *
 * kintone APIを消費せず、スプレッドシートの「年度別連携台帳（当年度＋前年度）」を直接高速スキャンし、
 * 前方紹介の未返書案件を担当医師別にグルーピング・経過日数バッジ付与の上、
 * LINE WORKS Bot（プレーンテキスト形式）および Slack へ通知する。
 */

const REMINDER_BADGE = Object.freeze({
  URGENT_14D: '🚨', // 14日以上経過（至急・最優先フォロー）
  WARN_7D: '⚠️',    // 7日〜13日経過（医師への作成依頼タイミング）
  RECENT: '📌',     // 7日未満（直近受診・経過観察）
});

const REMINDER_TRIGGER_HANDLER = 'runScheduledUnrepliedReminder';

/**
 * 基準日と受診日（YYYY-MM-DD）から経過日数（日単位）を算出する
 * @param {string|Date} visitDateInput
 * @param {Date=} baseDateInput
 * @return {number} 経過日数（0以上）
 */
function calculateElapsedDays(visitDateInput, baseDateInput) {
  const visitYmd = formatDateYMD(visitDateInput);
  if (!visitYmd) return 0;

  const baseYmd = formatDateYMD(baseDateInput || new Date());
  const visitParts = visitYmd.split('-').map((n) => parseInt(n, 10));
  const baseParts = baseYmd.split('-').map((n) => parseInt(n, 10));

  if (visitParts.length < 3 || baseParts.length < 3) return 0;

  const visitUtc = Date.UTC(visitParts[0], visitParts[1] - 1, visitParts[2]);
  const baseUtc = Date.UTC(baseParts[0], baseParts[1] - 1, baseParts[2]);
  const diffDays = Math.floor((baseUtc - visitUtc) / (1000 * 60 * 60 * 24));
  return diffDays >= 0 ? diffDays : 0;
}

/**
 * 経過日数に応じた3段階の優先度バッジを返す
 * - 14日以上経過: 🚨（至急・最優先フォロー）
 * - 7日〜13日経過: ⚠️（医師への作成依頼タイミング）
 * - 7日未満: 📌（直近受診・経過観察）
 *
 * @param {number} elapsedDays
 * @return {string}
 */
function getElapsedBadge(elapsedDays) {
  if (elapsedDays >= 14) return REMINDER_BADGE.URGENT_14D;
  if (elapsedDays >= 7) return REMINDER_BADGE.WARN_7D;
  return REMINDER_BADGE.RECENT;
}

/**
 * YYYY-MM-DD 文字列を MM/dd 形式に変換する
 * @param {string} ymd
 * @return {string}
 */
function formatShortMonthDay(ymd) {
  const norm = formatDateYMD(ymd);
  const m = norm.match(/^\d{4}-(\d{2})-(\d{2})$/);
  if (!m) return norm || '--/--';
  return `${m[1]}/${m[2]}`;
}

/**
 * 日付を「YYYY/MM/dd(曜)」形式で返す
 * @param {Date=} dateObj
 * @return {string}
 */
function formatDateWithJapaneseWeekday(dateObj) {
  const d = dateObj || new Date();
  const ymdSlash = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
  const weekdays = ['日', '月', '火', '水', '木', '金', '土'];
  const jstStr = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  const parts = jstStr.split('-').map((n) => parseInt(n, 10));
  const localDate = new Date(parts[0], parts[1] - 1, parts[2]);
  const wd = weekdays[localDate.getDay()] || '';
  return `${ymdSlash}(${wd})`;
}

/**
 * 医師名末尾の「先生」重複を防ぐためコア医師名を整形する
 * @param {string} rawDoctor
 * @param {string} fallbackText
 * @return {string}
 */
function cleanDoctorDisplayName(rawDoctor, fallbackText) {
  const trimmed = String(rawDoctor || '')
    .replace(/(先生|医師)+$/g, '')
    .trim();
  return trimmed || fallbackText;
}

/**
 * 当年度および前年度の連携台帳シートを走査し、未返書の前方紹介レコードをすべて抽出する
 *
 * 抽出条件:
 * - F列 `連携フェーズ` が `"前方"` で始まる
 * - O列 `返書日` が空欄
 * - N列 `返書ステータス` が `"作成・送付完了"`・`"対象外"` ではない
 *
 * @param {Date=} baseDate
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ss
 * @return {{
 *   totalCount: number,
 *   over14DaysCount: number,
 *   over7DaysCount: number,
 *   under7DaysCount: number,
 *   groupedByDoctor: Array<{
 *     doctorName: string,
 *     items: Array<{
 *       sheetName: string,
 *       rowNumber: number,
 *       kintoneRecordId: string,
 *       visitDate: string,
 *       visitDateShort: string,
 *       elapsedDays: number,
 *       badge: string,
 *       visitType: string,
 *       patientId: string,
 *       patientName: string,
 *       facilityName: string,
 *       facilityDoctor: string,
 *       ourDept: string,
 *       ourDoctor: string
 *     }>
 *   }>
 * }}
 */
function extractUnrepliedReferrals(baseDate, ss) {
  const nowDate = baseDate || new Date();
  const candidateSheets = getCandidateLedgerSheets(nowDate, ss);
  const idx = SANTE_CONFIG.LEDGER_COLUMNS.INDEX;

  const extractedItems = [];
  let over14DaysCount = 0;
  let over7DaysCount = 0;
  let under7DaysCount = 0;

  candidateSheets.forEach(({ sheet, sheetName }) => {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const values = sheet
      .getRange(2, 1, lastRow - 1, SANTE_CONFIG.LEDGER_COLUMNS.HEADERS.length)
      .getValues();

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const phaseStr = String(row[idx.PHASE] || '').trim();
      const replyDateStr = formatDateYMD(row[idx.REPLY_DATE]);
      const replyStatusStr = String(row[idx.REPLY_STATUS] || '').trim();

      // 条件1: F列 連携フェーズが "前方" で始まる
      if (!phaseStr.startsWith('前方')) {
        continue;
      }

      // 条件2: O列 返書日が空欄
      if (replyDateStr !== '') {
        continue;
      }

      // 条件3: N列 返書ステータスが "作成・送付完了"・"対象外" ではない
      if (
        replyStatusStr === SANTE_CONFIG.REPLY_STATUS.COMPLETED ||
        replyStatusStr.startsWith(SANTE_CONFIG.REPLY_STATUS.NOT_APPLICABLE)
      ) {
        continue;
      }

      const visitDate = formatDateYMD(row[idx.VISIT_DATE]);
      if (!visitDate) continue;

      const elapsedDays = calculateElapsedDays(visitDate, nowDate);
      const badge = getElapsedBadge(elapsedDays);

      if (elapsedDays >= 14) {
        over14DaysCount++;
      } else if (elapsedDays >= 7) {
        over7DaysCount++;
      } else {
        under7DaysCount++;
      }

      const rawOurDoctor = String(row[idx.OUR_DOCTOR] || '').trim();
      const ourDoctor = rawOurDoctor ? cleanDoctorDisplayName(rawOurDoctor, '担当医未定') : '担当医未定';

      extractedItems.push({
        sheetName: sheetName,
        rowNumber: i + 2,
        kintoneRecordId: String(row[idx.KINTONE_RECORD_ID] || '').trim(),
        visitDate: visitDate,
        visitDateShort: formatShortMonthDay(visitDate),
        elapsedDays: elapsedDays,
        badge: badge,
        visitType: String(row[idx.VISIT_TYPE] || SANTE_CONFIG.VISIT_TYPES.OUTPATIENT).trim(),
        patientId: String(row[idx.PATIENT_ID] || '未登録').trim() || '未登録',
        patientName: String(row[idx.PATIENT_NAME] || '氏名未入力').trim() || '氏名未入力',
        facilityName: String(row[idx.FACILITY_NAME] || '施設名未入力').trim() || '施設名未入力',
        facilityDoctor: cleanDoctorDisplayName(row[idx.FACILITY_DOCTOR], '担当医'),
        ourDept: String(row[idx.OUR_DEPT] || '').trim(),
        ourDoctor: ourDoctor,
      });
    }
  });

  // 当院担当医師（M列）ごとにグルーピング
  const doctorMap = new Map();
  extractedItems.forEach((item) => {
    if (!doctorMap.has(item.ourDoctor)) {
      doctorMap.set(item.ourDoctor, []);
    }
    doctorMap.get(item.ourDoctor).push(item);
  });

  // 各医師グループ内を経過日数が長い順（古い受診日順）にソート
  const groupedByDoctor = [];
  doctorMap.forEach((items, doctorName) => {
    items.sort((a, b) => {
      if (b.elapsedDays !== a.elapsedDays) {
        return b.elapsedDays - a.elapsedDays;
      }
      return a.visitDate.localeCompare(b.visitDate);
    });
    groupedByDoctor.push({
      doctorName: doctorName,
      items: items,
    });
  });

  // 医師グループの並び順: 最長経過日数が大きい医師順（「担当医未定」は末尾）
  groupedByDoctor.sort((a, b) => {
    if (a.doctorName === '担当医未定') return 1;
    if (b.doctorName === '担当医未定') return -1;
    const maxA = a.items.length > 0 ? a.items[0].elapsedDays : 0;
    const maxB = b.items.length > 0 ? b.items[0].elapsedDays : 0;
    if (maxB !== maxA) return maxB - maxA;
    return b.items.length - a.items.length;
  });

  return {
    totalCount: extractedItems.length,
    over14DaysCount: over14DaysCount,
    over7DaysCount: over7DaysCount,
    under7DaysCount: under7DaysCount,
    groupedByDoctor: groupedByDoctor,
  };
}

/**
 * LINE WORKS向けにMarkdown記号（* 等）を一切含まないプレーンテキスト＋絵文字・罫線形式で
 * 未返書リマインドメッセージを構築する
 *
 * @param {ReturnType<typeof extractUnrepliedReferrals>} reportData
 * @param {Date=} baseDate
 * @return {string}
 */
function buildUnrepliedReminderPlainText(reportData, baseDate) {
  const dateHeader = formatDateWithJapaneseWeekday(baseDate || new Date());
  const total = reportData.totalCount;
  const over14 = reportData.over14DaysCount;
  const over7Total = reportData.over14DaysCount + reportData.over7DaysCount;

  const lines = [
    '━━━━━━━━━━━━━━━━━━━━',
    '📋 【地域連携室】紹介患者 返書未作成リマインド',
    `集計日: ${dateHeader}`,
    `未作成合計: ${total}件（うち14日以上経過: ${over14}件 / 7日以上経過: ${over7Total}件）`,
    '━━━━━━━━━━━━━━━━━━━━',
  ];

  if (total === 0) {
    lines.push('');
    lines.push('✅ 現在、未返書の前方紹介案件はありません。');
    lines.push('━━━━━━━━━━━━━━━━━━━━');
    return lines.join('\n');
  }

  lines.push('🚨:14日以上経過(至急) ⚠️:7〜13日経過 📌:7日未満');

  reportData.groupedByDoctor.forEach((group) => {
    lines.push('');
    lines.push('────────────────────');
    lines.push(`■ ${group.doctorName} 先生（${group.items.length}件）`);
    lines.push('────────────────────');

    group.items.forEach((item, idx) => {
      if (idx > 0) {
        lines.push('');
      }
      lines.push(
        `${item.badge} ${item.visitDateShort}受診(${item.elapsedDays}日経過) [${item.visitType}]`
      );
      lines.push(`ID:${item.patientId} ${item.patientName} 様`);
      lines.push(`紹介元: ${item.facilityName} ${item.facilityDoctor}先生`);
    });
  });

  lines.push('');
  lines.push('━━━━━━━━━━━━━━━━━━━━');
  lines.push('💡 作成済みの返書（経過・結果報告）または逆紹介状は、');
  lines.push('スキャンしてGemini JSONをC2セルへ貼り付けると');
  lines.push('5重ロック照合によりkintone・連携台帳から自動消込されます。');
  lines.push('━━━━━━━━━━━━━━━━━━━━');

  return lines.join('\n');
}

/**
 * Base64URL エンコードヘルパー（LINE WORKS JWT署名用）
 * @param {string|Array<number>} input
 * @return {string}
 */
function base64UrlEncode(input) {
  const encoded =
    typeof input === 'string'
      ? Utilities.base64EncodeWebSafe(input, Utilities.Charset.UTF_8)
      : Utilities.base64EncodeWebSafe(input);
  return encoded.replace(/=+$/, '');
}

/**
 * LINE WORKS API 2.0 サービスアカウント認証（JWT RS256）でアクセストークンを取得する
 * @param {{
 *   lwClientId: string,
 *   lwClientSecret: string,
 *   lwServiceAccount: string,
 *   lwPrivateKey: string
 * }} secret
 * @return {string} accessToken
 */
function getLineWorksAccessToken(secret) {
  const nowSec = Math.floor(Date.now() / 1000);
  const header = { alg: 'RS256', typ: 'JWT' };
  const payload = {
    iss: secret.lwClientId,
    sub: secret.lwServiceAccount,
    iat: nowSec,
    exp: nowSec + 3600,
  };

  const signingInput = `${base64UrlEncode(JSON.stringify(header))}.${base64UrlEncode(
    JSON.stringify(payload)
  )}`;

  // PropertiesServiceに1行で保存された秘密鍵の `\n` エスケープを実際の改行コードへ復元
  const normalizedPrivateKey = secret.lwPrivateKey.replace(/\\n/g, '\n').trim();
  const signatureBytes = Utilities.computeRsaSha256Signature(signingInput, normalizedPrivateKey);
  const jwtAssertion = `${signingInput}.${base64UrlEncode(signatureBytes)}`;

  const tokenUrl = 'https://auth.worksmobile.com/oauth2/v2.0/token';
  const formPayload = {
    assertion: jwtAssertion,
    grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
    client_id: secret.lwClientId,
    client_secret: secret.lwClientSecret,
    scope: 'bot',
  };

  const response = UrlFetchApp.fetch(tokenUrl, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: formPayload,
    muteHttpExceptions: true,
  });

  const code = response.getResponseCode();
  const text = response.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error(`LINE WORKS OAuthトークン取得失敗 (HTTP ${code}): ${text}`);
  }

  const parsed = JSON.parse(text);
  if (!parsed.access_token) {
    throw new Error(`LINE WORKS OAuthレスポンスに access_token が含まれていません: ${text}`);
  }
  return parsed.access_token;
}

/**
 * LINE WORKS Bot 経由で指定のトークルーム（Channel ID）またはユーザーへプレーンテキストメッセージを送信する
 * @param {string} plainTextMessage
 * @return {{sent: boolean, reason?: string}}
 */
function sendLineWorksBotMessage(plainTextMessage) {
  const secret = getSecretConfig();
  const hasLwConfig =
    secret.lwClientId &&
    secret.lwClientSecret &&
    secret.lwServiceAccount &&
    secret.lwPrivateKey &&
    secret.lwBotId &&
    secret.lwReminderTargetId;

  if (!hasLwConfig) {
    Logger.log('[LINE WORKS通知スキップ] LW_* スクリプトプロパティが未設定です。');
    return { sent: false, reason: 'LW_PROPERTIES_NOT_CONFIGURED' };
  }

  const accessToken = getLineWorksAccessToken(secret);
  const botId = encodeURIComponent(secret.lwBotId);
  const targetId = encodeURIComponent(secret.lwReminderTargetId);

  const messagePayload = JSON.stringify({
    content: {
      type: 'text',
      text: plainTextMessage,
    },
  });

  const requestOptions = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
    payload: messagePayload,
    muteHttpExceptions: true,
  };

  // トークルーム（channels）宛てを優先試行し、404等の場合はユーザー個人（users）宛てにフォールバック
  const isLikelyUserAccount = secret.lwReminderTargetId.includes('@');
  const primaryUrl = isLikelyUserAccount
    ? `https://www.worksapis.com/v1.0/bots/${botId}/users/${targetId}/messages`
    : `https://www.worksapis.com/v1.0/bots/${botId}/channels/${targetId}/messages`;
  const fallbackUrl = isLikelyUserAccount
    ? `https://www.worksapis.com/v1.0/bots/${botId}/channels/${targetId}/messages`
    : `https://www.worksapis.com/v1.0/bots/${botId}/users/${targetId}/messages`;

  let res = UrlFetchApp.fetch(primaryUrl, requestOptions);
  let status = res.getResponseCode();

  if (status < 200 || status >= 300) {
    const fallbackRes = UrlFetchApp.fetch(fallbackUrl, requestOptions);
    if (fallbackRes.getResponseCode() >= 200 && fallbackRes.getResponseCode() < 300) {
      return { sent: true };
    }
    const errText = res.getContentText();
    notifySlackError(`LINE WORKS Bot送信エラー (HTTP ${status})`, errText, {
      botId: secret.lwBotId,
      targetId: secret.lwReminderTargetId,
    });
    throw new Error(`LINE WORKS Bot送信エラー (HTTP ${status}): ${errText}`);
  }

  return { sent: true };
}

/**
 * 未返書リストを抽出し、LINE WORKS Bot および Slack へリマインド通知を送信するメイン関数
 * @return {{
 *   totalCount: number,
 *   over14DaysCount: number,
 *   over7DaysCount: number,
 *   lineWorksSent: boolean,
 *   slackSent: boolean,
 *   messageText: string
 * }}
 */
function executeUnrepliedReminderNotification() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nowDate = new Date();
  const reportData = extractUnrepliedReferrals(nowDate, ss);
  const messageText = buildUnrepliedReminderPlainText(reportData, nowDate);

  let lineWorksSent = false;
  try {
    const lwResult = sendLineWorksBotMessage(messageText);
    lineWorksSent = lwResult.sent;
  } catch (lwErr) {
    Logger.log(`LINE WORKS送信エラー: ${lwErr.message}`);
  }

  // Slack Webhook が設定されている場合は Slack にも同時通知
  const slackSent = sendSlackNotification(
    '紹介患者 返書未作成リマインド',
    messageText,
    reportData.over14DaysCount > 0 ? 'WARNING' : 'INFO',
    {
      未作成合計: `${reportData.totalCount}件`,
      '14日以上経過(🚨)': `${reportData.over14DaysCount}件`,
      '7〜13日経過(⚠️)': `${reportData.over7DaysCount}件`,
      対象医師数: `${reportData.groupedByDoctor.length}名`,
    }
  );

  return {
    totalCount: reportData.totalCount,
    over14DaysCount: reportData.over14DaysCount,
    over7DaysCount: reportData.over7DaysCount,
    lineWorksSent: lineWorksSent,
    slackSent: slackSent,
    messageText: messageText,
  };
}

/**
 * 定期実行トリガーから呼び出されるエントリポイント関数（毎週月・木 9:00）
 */
function runScheduledUnrepliedReminder() {
  try {
    executeUnrepliedReminderNotification();
  } catch (err) {
    notifySlackError('未返書定期リマインド実行エラー', err);
    throw err;
  }
}

/**
 * メニュー「📢 未返書リストを今すぐ通知（LINE WORKS / Slack）」から呼び出される手動実行関数
 */
function sendUnrepliedRemindersManual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = executeUnrepliedReminderNotification();

  const channels = [];
  if (result.lineWorksSent) channels.push('LINE WORKS');
  if (result.slackSent) channels.push('Slack');

  const channelSummary =
    channels.length > 0
      ? `${channels.join(' / ')} へ通知しました`
      : '通知先プロパティ(LW_* / SLACK_WEBHOOK_URL)が未設定のためログ出力のみ完了しました';

  ss.toast(
    `未返書合計: ${result.totalCount}件（14日以上: ${result.over14DaysCount}件）を抽出し、${channelSummary}。`,
    '📢 未返書リマインド完了',
    8
  );
  Logger.log(result.messageText);
}

/**
 * メニュー「⚙️ 初期設定・メンテナンス > ⏰ 未返書リマインド定期トリガー設定（毎週月・木 9:00）」用関数
 * 既存の同名トリガーを整理した上で、毎週月曜・木曜 午前9時台の定期実行トリガーを登録する
 */
function setupReminderScheduleTrigger() {
  const existingTriggers = ScriptApp.getProjectTriggers();
  existingTriggers.forEach((trigger) => {
    if (trigger.getHandlerFunction() === REMINDER_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 毎週月曜日 9:00〜10:00
  ScriptApp.newTrigger(REMINDER_TRIGGER_HANDLER)
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();

  // 毎週木曜日 9:00〜10:00
  ScriptApp.newTrigger(REMINDER_TRIGGER_HANDLER)
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.THURSDAY)
    .atHour(9)
    .create();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) {
    ss.toast(
      '毎週月曜・木曜 9:00 に未返書リストを自動通知する定期トリガーを設定しました。',
      '⏰ 定期トリガー登録完了',
      6
    );
  }
}
