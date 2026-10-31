/**
 * ============================================================================
 * SANTE-BRIDGE (sante-bridge-liaison-hub) 本番コピペ用 全結合バンドルファイル
 * 結合順序: Config.js -> LedgerMatcher.js -> DocGenerator.js -> KintoneClient.js -> ReminderService.js -> Code.js
 * ============================================================================
 */
// ============================================================================
// FILE: src/Config.js
// ============================================================================
/**
 * @fileoverview SANTE-BRIDGE 設定・定数・秘匿情報管理・Slack通知モジュール
 * サンテ溝上病院 地域連携・紹介返書管理＆第一報自動化システム
 */

const SANTE_CONFIG = Object.freeze({
  SYSTEM_NAME: 'SANTE-BRIDGE (地域連携・紹介返書管理＆第一報自動化システム)',
  VERSION: '2.0.0',

  // シート名定義
  SHEETS: {
    INPUT: '紹介・逆紹介入力',
    FACILITY_MASTER: '施設マスター',
    LEDGER_SUFFIX: '年度_連携台帳', // 例: 2026年度_連携台帳
  },

  // タブ①: 紹介・逆紹介入力 シート座標定義
  INPUT_LAYOUT: {
    PASTE_CELL: 'C2',
    PASTE_RANGE: 'C2:C3',
    LABEL_COL: 2, // B列
    DATA_START_COL: 3, // C列以降に1患者＝1列で横展開
    STATUS_ROW: 4, // C4以降: 列処理ステータス表示行
    ROW_START: 5,
    ROW_END: 19,
    ROWS: {
      VISIT_DATE: 5,          // 5行目: 受診日 (YYYY-MM-DD)
      PATIENT_ID: 6,          // 6行目: 患者ID
      PATIENT_NAME: 7,        // 7行目: 患者氏名
      PHASE: 8,               // 8行目: 連携フェーズ
      VISIT_TYPE: 9,          // 9行目: 受入・送出形態
      FACILITY_NAME: 10,      // 10行目: 連携先施設名
      FACILITY_DEPT: 11,      // 11行目: 連携先診療科
      FACILITY_DOCTOR: 12,    // 12行目: 連携先医師名
      OUR_DEPT: 13,           // 13行目: 当院担当診療科
      OUR_DOCTOR: 14,         // 14行目: 担当医師
      NEED_FIRST_REPORT: 15,  // 15行目: 📝 第一報作成 (チェックボックス)
      REPORT_DOC_URL: 16,     // 16行目: 🔗 報告書ドキュメント (URL)
      FAX_SENT: 17,           // 17行目: ✅ FAX送信完了 (チェックボックス)
      TARGET_RECORD_INFO: 18, // 18行目: 🎯 返書消込ターゲットID
      REMARKS: 19,            // 19行目: 備考・連絡事項
    },
    LABELS: [
      '受診日',
      '患者ID',
      '患者氏名',
      '連携フェーズ',
      '受入・送出形態',
      '連携先施設名',
      '連携先診療科',
      '連携先医師名',
      '当院担当診療科',
      '担当医師',
      '📝 第一報作成',
      '🔗 報告書ドキュメント',
      '✅ FAX送信完了',
      '🎯 返書消込ターゲットID',
      '備考・連絡事項',
    ],
  },

  // 連携フェーズ定数
  PHASES: {
    FORWARD: '前方（紹介・受入）',
    BACKWARD: '後方（逆紹介・送出）',
    REPLY: '返書（経過・結果報告）',
    OTHER: 'その他（介護連携・対診）',
  },

  // 受入・送出形態定数
  VISIT_TYPES: {
    OUTPATIENT: '外来',
    INPATIENT: '入院',
  },

  // 返書ステータス定数
  REPLY_STATUS: {
    UNCREATED: '未作成',
    COMPLETED: '作成・送付完了',
    NOT_APPLICABLE: '対象外',
  },

  // FAX送信状況定数
  FAX_STATUS: {
    UNSENT: '未送信',
    SENT: '送信完了',
    NOT_REQUIRED: '不要',
  },

  // 列処理ステータス定数（4行目表示・二重POST防止用）
  COLUMN_SYNC_STATUS: {
    PENDING: '⏳ 未連携',
    SYNCED: '✅ kintone反映済',
    DOC_READY: '📄 第一報作成済（未連携）',
  },

  // タブ②: 年度別連携台帳 カラム定義（A列〜R列: 全18列）
  LEDGER_COLUMNS: {
    HEADERS: [
      'kintoneレコードID', // A列 (1)
      '登録日時',          // B列 (2)
      '受診日',            // C列 (3)
      '患者ID',            // D列 (4)
      '患者氏名',          // E列 (5)
      '連携フェーズ',      // F列 (6)
      '受入・送出形態',    // G列 (7)
      '更新用キー',        // H列 (8)
      '連携先施設名',      // I列 (9)
      '連携先診療科',      // J列 (10)
      '連携先医師名',      // K列 (11)
      '当院担当診療科',    // L列 (12)
      '担当医師',          // M列 (13)
      '返書ステータス',    // N列 (14)
      '返書日',            // O列 (15)
      '第一報URL',         // P列 (16)
      'FAX送信状況',       // Q列 (17)
      '備考・連絡事項',    // R列 (18)
    ],
    // 0-based インデックス（配列操作用）
    INDEX: {
      KINTONE_RECORD_ID: 0,
      REGISTERED_AT: 1,
      VISIT_DATE: 2,
      PATIENT_ID: 3,
      PATIENT_NAME: 4,
      PHASE: 5,
      VISIT_TYPE: 6,
      LOOKUP_KEY: 7,
      FACILITY_NAME: 8,
      FACILITY_DEPT: 9,
      FACILITY_DOCTOR: 10,
      OUR_DEPT: 11,
      OUR_DOCTOR: 12,
      REPLY_STATUS: 13,
      REPLY_DATE: 14,
      FIRST_REPORT_URL: 15,
      FAX_STATUS: 16,
      REMARKS: 17,
    },
    // 1-based 列番号（Sheet.getRange 用）
    COL_NUM: {
      KINTONE_RECORD_ID: 1,
      REGISTERED_AT: 2,
      VISIT_DATE: 3,
      PATIENT_ID: 4,
      PATIENT_NAME: 5,
      PHASE: 6,
      VISIT_TYPE: 7,
      LOOKUP_KEY: 8,
      FACILITY_NAME: 9,
      FACILITY_DEPT: 10,
      FACILITY_DOCTOR: 11,
      OUR_DEPT: 12,
      OUR_DOCTOR: 13,
      REPLY_STATUS: 14,
      REPLY_DATE: 15,
      FIRST_REPORT_URL: 16,
      FAX_STATUS: 17,
      REMARKS: 18,
    },
  },

  // タブ③: 施設マスター カラム定義
  FACILITY_MASTER_COLUMNS: {
    HEADERS: [
      '更新用キー',
      '正式名称',
      '住所',
      'FAX番号',
      '略称・別名1',
      '略称・別名2',
      '略称・別名3',
    ],
    INDEX: {
      LOOKUP_KEY: 0,    // A列
      OFFICIAL_NAME: 1, // B列
      ADDRESS: 2,       // C列
      FAX_NUMBER: 3,    // D列
      ALIAS_START: 4,   // E列以降すべて略称・別名リストとして走査
    },
  },

  // kintone アプリB (App ID: 195) デフォルト設定＆フィールドコードマッピング
  KINTONE: {
    DEFAULT_APP_ID: '195',
    BATCH_LIMIT: 100, // kintone REST API 1リクエストあたりの最大レコード数
    MAX_AUTO_RECOVERY_RETRIES: 8, // フィールド未定義エラー時の自動除外リトライ上限
    FIELD_CODES: {
      VISIT_DATE: '受診日',
      PATIENT_ID: '患者ID',
      PATIENT_NAME: '患者氏名',
      PHASE: '連携フェーズ',
      DIRECTION: '連携方向', // 既存フィールド互換用（前方（紹介・受入） / 後方（逆紹介・送出））
      VISIT_TYPE: '受入・送出形態',
      LOOKUP_KEY: '更新用キー',
      FACILITY_NAME: '連携先施設名',
      FACILITY_DEPT: '連携先診療科',
      FACILITY_DOCTOR: '連携先医師名',
      OUR_DEPT: '当院担当診療科',
      OUR_DOCTOR: '担当医師',
      REPLY_STATUS: '返書ステータス',
      REPLY_DATE: '返書日',
      FIRST_REPORT_URL: '第一報URL',
      FAX_STATUS: 'FAX送信状況',
      REMARKS: '備考・連絡事項',
    },
  },

  // PropertiesService のキー名定義
  PROPERTY_KEYS: {
    KINTONE_DOMAIN: 'KINTONE_DOMAIN',                   // 例: sante-mizogami.cybozu.com
    KINTONE_APP_ID: 'KINTONE_APP_ID',                   // デフォルト: 195
    KINTONE_API_TOKEN: 'KINTONE_API_TOKEN',             // アプリ195のAPIトークン（閲覧・追加・編集権限）
    KINTONE_API_TOKEN_MASTER: 'KINTONE_API_TOKEN_MASTER', // ルックアップ参照元（医療機関・施設マスター）のAPIトークン
    REPORT_TEMPLATE_DOC_ID: 'REPORT_TEMPLATE_DOC_ID',   // 第一報GoogleドキュメントテンプレートID
    REPORT_OUTPUT_FOLDER_ID: 'REPORT_OUTPUT_FOLDER_ID', // 生成した第一報ドキュメントの保存先フォルダID
    SLACK_WEBHOOK_URL: 'SLACK_WEBHOOK_URL',             // エラー・マイルストーン通知用Slack Incoming Webhook URL
    // LINE WORKS Bot (API 2.0) 定期リマインド通知用
    LW_CLIENT_ID: 'LW_CLIENT_ID',
    LW_CLIENT_SECRET: 'LW_CLIENT_SECRET',
    LW_SERVICE_ACCOUNT: 'LW_SERVICE_ACCOUNT',
    LW_PRIVATE_KEY: 'LW_PRIVATE_KEY',
    LW_BOT_ID: 'LW_BOT_ID',
    LW_REMINDER_TARGET_ID: 'LW_REMINDER_TARGET_ID',     // 送信先トークルーム(Channel)ID または ユーザーID
  },
});

/**
 * PropertiesService から秘匿設定値を取得する
 * @return {{
 *   kintoneDomain: string,
 *   kintoneAppId: string,
 *   kintoneApiToken: string,
 *   kintoneApiTokenMaster: string,
 *   reportTemplateDocId: string,
 *   reportOutputFolderId: string,
 *   slackWebhookUrl: string,
 *   lwClientId: string,
 *   lwClientSecret: string,
 *   lwServiceAccount: string,
 *   lwPrivateKey: string,
 *   lwBotId: string,
 *   lwReminderTargetId: string
 * }}
 */
function getSecretConfig() {
  const props = PropertiesService.getScriptProperties();
  const keys = SANTE_CONFIG.PROPERTY_KEYS;

  return {
    kintoneDomain: (props.getProperty(keys.KINTONE_DOMAIN) || '').trim(),
    kintoneAppId: (props.getProperty(keys.KINTONE_APP_ID) || SANTE_CONFIG.KINTONE.DEFAULT_APP_ID).trim(),
    kintoneApiToken: (props.getProperty(keys.KINTONE_API_TOKEN) || '').trim(),
    kintoneApiTokenMaster: (props.getProperty(keys.KINTONE_API_TOKEN_MASTER) || '').trim(),
    reportTemplateDocId: (props.getProperty(keys.REPORT_TEMPLATE_DOC_ID) || '').trim(),
    reportOutputFolderId: (props.getProperty(keys.REPORT_OUTPUT_FOLDER_ID) || '').trim(),
    slackWebhookUrl: (props.getProperty(keys.SLACK_WEBHOOK_URL) || '').trim(),
    lwClientId: (props.getProperty(keys.LW_CLIENT_ID) || '').trim(),
    lwClientSecret: (props.getProperty(keys.LW_CLIENT_SECRET) || '').trim(),
    lwServiceAccount: (props.getProperty(keys.LW_SERVICE_ACCOUNT) || '').trim(),
    lwPrivateKey: (props.getProperty(keys.LW_PRIVATE_KEY) || '').trim(),
    lwBotId: (props.getProperty(keys.LW_BOT_ID) || '').trim(),
    lwReminderTargetId: (props.getProperty(keys.LW_REMINDER_TARGET_ID) || '').trim(),
  };
}

/**
 * 初回セットアップ用: スクリプトプロパティのプレースホルダーを初期化する
 * （GASエディタから1回実行、または設定ダイアログから更新可能）
 */
function initScriptPropertiesTemplate() {
  const props = PropertiesService.getScriptProperties();
  const current = props.getProperties();
  const defaults = {
    [SANTE_CONFIG.PROPERTY_KEYS.KINTONE_DOMAIN]: 'your-subdomain.cybozu.com',
    [SANTE_CONFIG.PROPERTY_KEYS.KINTONE_APP_ID]: SANTE_CONFIG.KINTONE.DEFAULT_APP_ID,
    [SANTE_CONFIG.PROPERTY_KEYS.KINTONE_API_TOKEN]: '',
    [SANTE_CONFIG.PROPERTY_KEYS.KINTONE_API_TOKEN_MASTER]: '',
    [SANTE_CONFIG.PROPERTY_KEYS.REPORT_TEMPLATE_DOC_ID]: '',
    [SANTE_CONFIG.PROPERTY_KEYS.REPORT_OUTPUT_FOLDER_ID]: '',
    [SANTE_CONFIG.PROPERTY_KEYS.SLACK_WEBHOOK_URL]: '',
    [SANTE_CONFIG.PROPERTY_KEYS.LW_CLIENT_ID]: '',
    [SANTE_CONFIG.PROPERTY_KEYS.LW_CLIENT_SECRET]: '',
    [SANTE_CONFIG.PROPERTY_KEYS.LW_SERVICE_ACCOUNT]: '',
    [SANTE_CONFIG.PROPERTY_KEYS.LW_PRIVATE_KEY]: '',
    [SANTE_CONFIG.PROPERTY_KEYS.LW_BOT_ID]: '',
    [SANTE_CONFIG.PROPERTY_KEYS.LW_REMINDER_TARGET_ID]: '',
  };

  Object.keys(defaults).forEach((key) => {
    if (!Object.prototype.hasOwnProperty.call(current, key)) {
      props.setProperty(key, defaults[key]);
    }
  });
  Logger.log('ScriptProperties のテンプレートキーを初期化しました。');
}

/**
 * 指定日付（未指定時は現在日時）から日本の会計年度（4月1日〜翌3月31日）を算出する
 * @param {Date|string=} dateInput
 * @return {number} 会計年度（例: 2026）
 */
function getFiscalYear(dateInput) {
  let d;
  if (!dateInput) {
    d = new Date();
  } else if (dateInput instanceof Date && !isNaN(dateInput.getTime())) {
    d = dateInput;
  } else {
    const parsed = new Date(String(dateInput).replace(/\//g, '-'));
    d = isNaN(parsed.getTime()) ? new Date() : parsed;
  }

  const year = d.getFullYear();
  const month = d.getMonth() + 1; // 1〜12
  return month >= 4 ? year : year - 1;
}

/**
 * 会計年度に対応する台帳シート名を返す（例: "2026年度_連携台帳"）
 * @param {number=} fiscalYear
 * @return {string}
 */
function getLedgerSheetName(fiscalYear) {
  const fy = fiscalYear || getFiscalYear();
  return `${fy}${SANTE_CONFIG.SHEETS.LEDGER_SUFFIX}`;
}

/**
 * 日付を YYYY-MM-DD 形式の文字列に正規化する
 * @param {Date|string|null|undefined} value
 * @return {string} YYYY-MM-DD 形式（変換不可の場合は空文字またはトリム済文字列）
 */
function formatDateYMD(value) {
  if (value === null || value === undefined || value === '') {
    return '';
  }
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  const str = String(value)
    .trim()
    .replace(/[年月]/g, '-')
    .replace(/日/g, '')
    .replace(/\//g, '-')
    .replace(/\./g, '-');

  const match = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (match) {
    const y = match[1];
    const m = match[2].padStart(2, '0');
    const d = match[3].padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  return str;
}

/**
 * 現在日時を YYYY-MM-DD HH:mm:ss 形式で取得する
 * @return {string}
 */
function formatDateTimeNow() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
}

/**
 * Slack Incoming Webhook へメッセージを送信する
 * @param {string} title 通知タイトル
 * @param {string} message 本文メッセージ
 * @param {'INFO'|'SUCCESS'|'WARNING'|'ERROR'=} level 通知レベル
 * @param {Object=} extraFields 追加フィールド（key-value）
 * @return {boolean} 送信成功可否
 */
function sendSlackNotification(title, message, level, extraFields) {
  const { slackWebhookUrl } = getSecretConfig();
  if (!slackWebhookUrl) {
    Logger.log(
      `[Slack通知スキップ(URL未設定)] [${level || 'INFO'}] ${title} (length: ${String(message || '').length})`
    );
    return false;
  }

  const levelMap = {
    INFO: { emoji: 'ℹ️', color: '#2eb886' },
    SUCCESS: { emoji: '✅', color: '#2eb886' },
    WARNING: { emoji: '⚠️', color: '#daa038' },
    ERROR: { emoji: '🚨', color: '#e01e5a' },
  };
  const badge = levelMap[level] || levelMap.INFO;

  const fields = [];
  if (extraFields && typeof extraFields === 'object') {
    Object.keys(extraFields).forEach((k) => {
      fields.push({
        title: k,
        value: String(extraFields[k]),
        short: true,
      });
    });
  }

  const payload = {
    text: `${badge.emoji} *${SANTE_CONFIG.SYSTEM_NAME}*\n*${title}*`,
    attachments: [
      {
        color: badge.color,
        text: message,
        fields: fields,
        footer: `SANTE-BRIDGE v${SANTE_CONFIG.VERSION} | ${formatDateTimeNow()}`,
      },
    ],
  };

  try {
    const response = UrlFetchApp.fetch(slackWebhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    const code = response.getResponseCode();
    if (code < 200 || code >= 300) {
      Logger.log(`Slack通知エラー (HTTP ${code}): ${response.getContentText()}`);
      return false;
    }
    return true;
  } catch (err) {
    Logger.log(`Slack通知例外: ${err && err.message ? err.message : err}`);
    return false;
  }
}

/**
 * 致命的エラー発生時にSlackへアラート通知を送信する
 * @param {string} context エラー発生箇所（例: "kintone POST通信"）
 * @param {Error|string} error エラーオブジェクトまたはメッセージ
 * @param {Object=} metadata 付随メタデータ
 * @param {{includeStack?: boolean}=} options オプション設定
 */
function notifySlackError(context, error, metadata, options) {
  const includeStack = !options || options.includeStack !== false;
  const errorMsg =
    error instanceof Error
      ? includeStack
        ? `${error.message}\n\`\`\`${error.stack || ''}\`\`\``
        : String(error.message || '')
      : String(error);
  sendSlackNotification(
    `致命的エラー発生: ${context}`,
    errorMsg,
    'ERROR',
    metadata
  );
}

/**
 * 開発マイルストーン・デプロイ完了・バッチ処理完了をSlackへ通知する
 * @param {string} milestoneTitle マイルストーン名
 * @param {string} summary 概要メッセージ
 * @param {Object=} stats 統計情報
 */
function notifySlackMilestone(milestoneTitle, summary, stats) {
  sendSlackNotification(milestoneTitle, summary, 'SUCCESS', stats);
}

// ============================================================================
// FILE: src/LedgerMatcher.js
// ============================================================================
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
 *   isMatched: boolean,
 *   matchType: 'exact'|'core'|'partial'|'none',
 *   candidates: Array<{officialName: string, lookupKey: string}>
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
      matchType: 'none',
      candidates: [],
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
        matchType: 'exact',
        candidates: [],
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
          matchType: 'core',
          candidates: [],
        };
      }
    }
  }

  // Step 3: 部分一致（2文字以上の場合のみ）
  if (coreInput.length >= 2) {
    const candidates = [];
    for (let i = 0; i < list.length; i++) {
      const item = list[i];
      const matched = item.coreCandidates.some(
        (cand) => cand.length >= 2 && (cand.includes(coreInput) || coreInput.includes(cand))
      );
      if (matched) {
        candidates.push({
          officialName: item.officialName,
          lookupKey: item.lookupKey,
        });
      }
    }
    if (candidates.length > 0) {
      return {
        lookupKey: '',
        officialName: trimmedInput,
        address: '',
        faxNumber: '',
        isMatched: false,
        matchType: 'partial',
        candidates: candidates,
      };
    }
  }

  // 未ヒット時は元の施設名を維持して返却
  return {
    lookupKey: '',
    officialName: trimmedInput,
    address: '',
    faxNumber: '',
    isMatched: false,
    matchType: 'none',
    candidates: [],
  };
}

/**
 * 年度別台帳シートにヘッダー装飾（#0F4C5C）、行列スクロール固定（1行目・5列目E列まで）、
 * およびN列（返書ステータス）の条件付き書式ルールを適用する
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function applyLedgerSheetFormatting(sheet) {
  if (!sheet) return;

  const headers = SANTE_CONFIG.LEDGER_COLUMNS.HEADERS;
  const headerRange = sheet.getRange(1, 1, 1, headers.length);

  // ヘッダーが未設定の場合は見出し文字列をセット
  const firstCellVal = String(sheet.getRange(1, 1).getValue() || '').trim();
  if (!firstCellVal) {
    headerRange.setValues([headers]);
  }

  // A1:R1 ヘッダー装飾（ディープティール #0F4C5C / 白太字 / 中央揃え）
  headerRange
    .setBackground('#0F4C5C')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // 1行目（ヘッダー）と A〜E列（5列目: 患者氏名まで）を常時スクロール固定
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(5);

  // N列（返書ステータス: 14列目）への条件付き書式設定
  const replyStatusCol = SANTE_CONFIG.LEDGER_COLUMNS.COL_NUM.REPLY_STATUS; // 14 (N列)
  const maxRows = Math.max(sheet.getMaxRows(), 500);
  const statusRange = sheet.getRange(2, replyStatusCol, maxRows - 1, 1);

  const uncreatedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(SANTE_CONFIG.REPLY_STATUS.UNCREATED)
    .setBackground('#F4CCCC')
    .setRanges([statusRange])
    .build();

  const completedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(SANTE_CONFIG.REPLY_STATUS.COMPLETED)
    .setBackground('#D9EAD3')
    .setRanges([statusRange])
    .build();

  // 既存の他列ルールは維持しつつ、N列の重複ルールを防いで設定
  const existingRules = sheet.getConditionalFormatRules() || [];
  const filteredRules = existingRules.filter((rule) => {
    const ranges = rule.getRanges();
    return !ranges.some((r) => r.getColumn() === replyStatusCol);
  });

  filteredRules.push(uncreatedRule, completedRule);
  sheet.setConditionalFormatRules(filteredRules);
}

/**
 * 指定された会計年度の台帳シートを取得する（存在しない場合は新規作成・ヘッダー初期化する）
 * @param {number=} fiscalYear 会計年度（例: 2026）
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet=} ss
 * @param {boolean=} forceFormat 既存シートにも強制的にUI固定・条件付き書式を再適用するか
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateLedgerSheet(fiscalYear, ss, forceFormat) {
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = getLedgerSheetName(fiscalYear);
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    applyLedgerSheetFormatting(sheet);
  } else if (forceFormat) {
    applyLedgerSheetFormatting(sheet);
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
  return coreInput === coreLedger;
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

// ============================================================================
// FILE: src/DocGenerator.js
// ============================================================================
/**
 * @fileoverview 第一報（受診・入院報告書）Googleドキュメント自動生成モジュール
 * サンテ溝上病院 地域連携・紹介返書管理＆第一報自動化システム (SANTE-BRIDGE)
 */

/**
 * Googleドキュメント（Body / Header / Footer）内のプレースホルダーを一括置換する
 * @param {GoogleAppsScript.Document.Document} doc
 * @param {Object<string, string>} replacements キー: プレースホルダー名（例: "{{患者氏名}}"）, 値: 置換文字列
 */
function replacePlaceholdersInDocument(doc, replacements) {
  const sections = [doc.getBody(), doc.getHeader(), doc.getFooter()].filter(Boolean);

  Object.keys(replacements).forEach((placeholder) => {
    // DocumentApp.replaceText は正規表現として解釈されるため `{` `}` をエスケープ
    const escapedPattern = placeholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const safeValue = String(replacements[placeholder] !== undefined && replacements[placeholder] !== null ? replacements[placeholder] : '');

    sections.forEach((section) => {
      section.replaceText(escapedPattern, safeValue);
    });
  });
}

/**
 * 単一患者データから第一報（受診・入院報告書）Googleドキュメントを複製・生成する
 *
 * 置換対象プレースホルダー（全12項目）:
 *   {{報告日}}, {{連携先施設名}}, {{連携先診療科}}, {{連携先医師名}}, {{連携先FAX番号}},
 *   {{患者氏名}}, {{患者ID}}, {{受診日}}, {{受入・送出形態}}, {{当院担当診療科}}, {{担当医師}}, {{備考}}
 *
 * @param {{
 *   visitDate: string,
 *   patientId: string,
 *   patientName: string,
 *   phase: string,
 *   visitType: string,
 *   facilityName: string,
 *   facilityDept: string,
 *   facilityDoctor: string,
 *   faxNumber?: string,
 *   ourDept: string,
 *   ourDoctor: string,
 *   remarks: string
 * }} recordData
 * @param {Array<Object>=} facilityMasterList
 * @return {{docId: string, docUrl: string, docName: string}}
 */
function generateFirstReportDoc(recordData, facilityMasterList) {
  const secret = getSecretConfig();
  let templateDocId = secret.reportTemplateDocId;

  if (!templateDocId) {
    throw new Error(
      'スクリプトプロパティ「REPORT_TEMPLATE_DOC_ID」が未設定です。メニュー「⚙️ 初期設定 > 📄 第一報テンプレート自動生成」を実行するか、テンプレートのドキュメントIDを設定してください。'
    );
  }

  // 施設マスターから正式名称・FAX番号を補完
  const resolvedFacility = resolveFacility(recordData.facilityName, facilityMasterList);
  const facilityName = resolvedFacility.officialName || recordData.facilityName || '';
  const isFacilityConfirmed =
    resolvedFacility.matchType === 'exact' || resolvedFacility.matchType === 'core';
  const faxNumber = isFacilityConfirmed
    ? recordData.faxNumber || resolvedFacility.faxNumber || ''
    : '';

  const reportDate = formatDateYMD(new Date());
  const visitDate = formatDateYMD(recordData.visitDate) || reportDate;
  const patientName = String(recordData.patientName || '氏名未入力').trim();
  const visitType = String(recordData.visitType || SANTE_CONFIG.VISIT_TYPES.OUTPATIENT).trim();

  // ファイル名生成: 【第一報】2026-10-29_山田太郎様_〇〇クリニック
  const docTitle = `【第一報】${visitDate}_${patientName}様_${facilityName || '連携先医療機関'}`;

  let templateFile;
  try {
    templateFile = DriveApp.getFileById(templateDocId);
  } catch (err) {
    notifySlackError(
      '第一報テンプレート取得失敗',
      err,
      {
        templateDocId: templateDocId,
        patientId: recordData.patientId || '',
        column: recordData.colNumber || '',
      },
      { includeStack: false }
    );
    throw new Error(`第一報テンプレートファイル(ID: ${templateDocId})の取得に失敗しました: ${err.message}`);
  }

  // 出力先フォルダの指定があればそのフォルダへ複製、なければテンプレートと同じ親フォルダへ複製
  let copiedFile;
  if (secret.reportOutputFolderId) {
    try {
      const outputFolder = DriveApp.getFolderById(secret.reportOutputFolderId);
      copiedFile = templateFile.makeCopy(docTitle, outputFolder);
    } catch (folderErr) {
      Logger.log(`出力先フォルダ(${secret.reportOutputFolderId})取得失敗のためデフォルト保存先を使用します: ${folderErr.message}`);
      copiedFile = templateFile.makeCopy(docTitle);
    }
  } else {
    copiedFile = templateFile.makeCopy(docTitle);
  }

  const copiedDocId = copiedFile.getId();
  const doc = DocumentApp.openById(copiedDocId);

  const replacements = {
    '{{報告日}}': reportDate,
    '{{連携先施設名}}': facilityName,
    '{{連携先診療科}}': String(recordData.facilityDept || ''),
    '{{連携先医師名}}': String(recordData.facilityDoctor || ''),
    '{{連携先FAX番号}}': isFacilityConfirmed ? faxNumber : '（要確認: 施設マスター未一致）',
    '{{患者氏名}}': patientName,
    '{{患者ID}}': String(recordData.patientId || ''),
    '{{受診日}}': visitDate,
    '{{受入・送出形態}}': visitType,
    '{{当院担当診療科}}': String(recordData.ourDept || ''),
    '{{担当医師}}': String(recordData.ourDoctor || ''),
    '{{備考}}': String(recordData.remarks || ''),
  };

  replacePlaceholdersInDocument(doc, replacements);
  doc.saveAndClose();

  return {
    docId: copiedDocId,
    docUrl: copiedFile.getUrl(),
    docName: docTitle,
  };
}

/**
 * 初回構築支援用: サンテ溝上病院「受診・入院報告書（第一報）」の標準Googleドキュメントテンプレートを自動生成し、
 * スクリプトプロパティ `REPORT_TEMPLATE_DOC_ID` に自動登録する
 * @return {{docId: string, docUrl: string}}
 */
function createDefaultReportTemplateDoc() {
  const doc = DocumentApp.create('【テンプレート】サンテ溝上病院_受診・入院報告書（第一報）');
  const body = doc.getBody();
  body.clear();

  const titlePara = body.appendParagraph('ご紹介患者様 受診・入院のご報告（第一報）');
  titlePara.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  titlePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  const datePara = body.appendParagraph('報告日: {{報告日}}');
  datePara.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  body.appendParagraph(
    '{{連携先施設名}}\n{{連携先診療科}}　{{連携先医師名}} 先生 御机下\n（送信先FAX番号: {{連携先FAX番号}}）'
  );

  const senderPara = body.appendParagraph(
    'サンテ溝上病院 地域医療連携室\n担当診療科: {{当院担当診療科}} / 担当医師: {{担当医師}}'
  );
  senderPara.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  body.appendParagraph(
    '\n平素より当院の地域医療連携に格別のご高配を賜り、厚く御礼申し上げます。\nご紹介いただきました下記患者様が当院を受診されましたので、取り急ぎ第一報をご報告申し上げます。\n'
  );

  const tableData = [
    ['患者氏名', '{{患者氏名}} 様', '当院患者ID', '{{患者ID}}'],
    ['受診・入院日', '{{受診日}}', '受入・送出形態', '{{受入・送出形態}}'],
    ['当院担当診療科', '{{当院担当診療科}}', '当院担当医師', '{{担当医師}}'],
    ['備考・連絡事項', '{{備考}}', '', ''],
  ];

  const table = body.appendTable(tableData);
  table.setBorderWidth(1);

  body.appendParagraph(
    '\n※ 詳細な診療経過・検査結果につきましては、後日改めて担当医より診療情報提供書（返書）にてご報告申し上げます。\nご不明な点がございましたら、地域医療連携室までお問い合わせください。'
  );

  doc.saveAndClose();

  const docId = doc.getId();
  const docUrl = doc.getUrl();
  PropertiesService.getScriptProperties().setProperty(
    SANTE_CONFIG.PROPERTY_KEYS.REPORT_TEMPLATE_DOC_ID,
    docId
  );

  Logger.log(`第一報テンプレートを作成し、REPORT_TEMPLATE_DOC_ID に設定しました: ${docUrl}`);
  return { docId, docUrl };
}

// ============================================================================
// FILE: src/KintoneClient.js
// ============================================================================
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

// ============================================================================
// FILE: src/ReminderService.js
// ============================================================================
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
    let errorCode = 'UNKNOWN';
    try {
      const parsedErr = JSON.parse(errText);
      if (parsedErr && (parsedErr.code || parsedErr.error)) {
        errorCode = String(parsedErr.code || parsedErr.error);
      }
    } catch (_) {
      // 応答本文はSlackに送らない
    }
    notifySlackError(
      `LINE WORKS Bot送信エラー (HTTP ${status})`,
      `HTTP ${status} / code: ${errorCode}`,
      {
        httpStatus: String(status),
        errorCode: errorCode,
      }
    );
    throw new Error(`LINE WORKS Bot送信エラー (HTTP ${status}, code: ${errorCode})`);
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

  // Slack Webhook が設定されている場合は件数のみ通知（氏名・ID・施設名・医師名は送らない）
  const slackSummaryText = `返書未作成 ${reportData.totalCount}件（うち期限超過 ${reportData.over14DaysCount}件）。詳細はLINE WORKSを確認`;
  const slackSent = sendSlackNotification(
    '紹介患者 返書未作成リマインド',
    slackSummaryText,
    reportData.over14DaysCount > 0 ? 'WARNING' : 'INFO',
    {
      未作成合計: `${reportData.totalCount}件`,
      '14日以上経過(🚨)': `${reportData.over14DaysCount}件`,
      '7〜13日経過(⚠️)': `${reportData.over7DaysCount}件`,
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
  Logger.log(
    `返書未作成 ${result.totalCount}件（うち期限超過 ${result.over14DaysCount}件 / 7〜13日経過 ${result.over7DaysCount}件）`
  );
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

// ============================================================================
// FILE: src/Code.js
// ============================================================================
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
