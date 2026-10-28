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
    Logger.log(`[Slack通知スキップ(URL未設定)] [${level || 'INFO'}] ${title}: ${message}`);
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
 */
function notifySlackError(context, error, metadata) {
  const errorMsg = error instanceof Error ? `${error.message}\n\`\`\`${error.stack || ''}\`\`\`` : String(error);
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
