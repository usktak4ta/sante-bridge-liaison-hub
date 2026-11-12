/**
 * @fileoverview SANTE-BRIDGE 全自動E2Eモック検証スクリプト
 * GAS環境（SpreadsheetApp, DriveApp, DocumentApp, UrlFetchApp, PropertiesService, Logger, LockService）を
 * ローカルのNode.js上でモック化し、タスク1〜3およびコンプライアンス基準を自動検証する。
 * ※ 外部API通信は一切発生せず、すべて架空のダミーデータ（「山田太郎」「○○中央病院」「△△中央病院」等）で検証する。
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const ROOT_DIR = path.resolve(__dirname, '..');
const SRC_DIR = path.join(ROOT_DIR, 'src');

// ============================================================================
// Ин-メモリ Spreadsheet / Sheet / Range モック実装
// ============================================================================
class MockRange {
  constructor(sheet, row, col, numRows = 1, numCols = 1) {
    this.sheet = sheet;
    this.row = row;
    this.col = col;
    this.numRows = numRows;
    this.numCols = numCols;
  }
  getSheet() {
    return this.sheet;
  }
  getRow() {
    return this.row;
  }
  getColumn() {
    return this.col;
  }
  getNumRows() {
    return this.numRows;
  }
  getNumColumns() {
    return this.numCols;
  }
  getValue() {
    return this.sheet._getCell(this.row, this.col).value;
  }
  setValue(val) {
    this.sheet._getCell(this.row, this.col).value = val;
    return this;
  }
  getValues() {
    const res = [];
    for (let r = 0; r < this.numRows; r++) {
      const rowArr = [];
      for (let c = 0; c < this.numCols; c++) {
        rowArr.push(this.sheet._getCell(this.row + r, this.col + c).value);
      }
      res.push(rowArr);
    }
    return res;
  }
  setValues(matrix) {
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) {
        this.sheet._getCell(this.row + r, this.col + c).value =
          matrix[r] && matrix[r][c] !== undefined ? matrix[r][c] : '';
      }
    }
    return this;
  }
  getNote() {
    return this.sheet._getCell(this.row, this.col).note;
  }
  setNote(noteStr) {
    this.sheet._getCell(this.row, this.col).note = String(noteStr || '');
    this.sheet._noteEvents.push({
      type: 'setNote',
      row: this.row,
      col: this.col,
      note: String(noteStr || ''),
    });
    return this;
  }
  getNotes() {
    const res = [];
    for (let r = 0; r < this.numRows; r++) {
      const rowArr = [];
      for (let c = 0; c < this.numCols; c++) {
        rowArr.push(this.sheet._getCell(this.row + r, this.col + c).note);
      }
      res.push(rowArr);
    }
    return res;
  }
  setNotes(notesMatrix) {
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) {
        const val = notesMatrix[r] && notesMatrix[r][c] !== undefined ? String(notesMatrix[r][c]) : '';
        this.sheet._getCell(this.row + r, this.col + c).note = val;
        this.sheet._noteEvents.push({
          type: 'setNotes',
          row: this.row + r,
          col: this.col + c,
          note: val,
        });
      }
    }
    return this;
  }
  clearNote() {
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) {
        this.sheet._getCell(this.row + r, this.col + c).note = '';
        this.sheet._noteEvents.push({
          type: 'clearNote',
          row: this.row + r,
          col: this.col + c,
          note: '',
        });
      }
    }
    return this;
  }
  clearContent() {
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) {
        this.sheet._getCell(this.row + r, this.col + c).value = '';
      }
    }
    return this;
  }
  clearDataValidations() {
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) {
        this.sheet._getCell(this.row + r, this.col + c).dataValidation = null;
      }
    }
    return this;
  }
  setNumberFormat(fmt) {
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) {
        this.sheet._getCell(this.row + r, this.col + c).numberFormat = String(fmt || '');
      }
    }
    return this;
  }
  getNumberFormat() {
    return this.sheet._getCell(this.row, this.col).numberFormat || '';
  }
  setDataValidation(rule) {
    this.sheet._lastDataValidation = rule;
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) {
        this.sheet._getCell(this.row + r, this.col + c).dataValidation = rule;
      }
    }
    return this;
  }
  getDataValidation() {
    return this.sheet._getCell(this.row, this.col).dataValidation || null;
  }
  insertCheckboxes() {
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) {
        this.sheet._getCell(this.row + r, this.col + c).hasCheckbox = true;
      }
    }
    return this;
  }
  hasCheckbox() {
    return Boolean(this.sheet._getCell(this.row, this.col).hasCheckbox);
  }
  setBorder() {
    return this;
  }
  setHorizontalAlignment() {
    return this;
  }
  setFontWeight() {
    return this;
  }
  setFontColor() {
    return this;
  }
  setFontSize() {
    return this;
  }
  setBackground() {
    return this;
  }
}

class MockSheet {
  constructor(name, parent) {
    this.name = name;
    this.parent = parent;
    this.cells = new Map();
    this._noteEvents = [];
    this.conditionalRules = [];
  }
  getName() {
    return this.name;
  }
  getParent() {
    return this.parent;
  }
  _key(r, c) {
    return `${r}:${c}`;
  }
  _getCell(r, c) {
    const k = this._key(r, c);
    if (!this.cells.has(k)) {
      this.cells.set(k, { value: '', note: '' });
    }
    return this.cells.get(k);
  }
  getLastRow() {
    let maxR = 0;
    for (const [k, cell] of this.cells.entries()) {
      if (cell.value !== '' && cell.value !== null && cell.value !== undefined) {
        const r = parseInt(k.split(':')[0], 10);
        if (r > maxR) maxR = r;
      }
    }
    return maxR;
  }
  getLastColumn() {
    let maxC = 0;
    for (const [k, cell] of this.cells.entries()) {
      if (cell.value !== '' && cell.value !== null && cell.value !== undefined) {
        const c = parseInt(k.split(':')[1], 10);
        if (c > maxC) maxC = c;
      }
    }
    return maxC;
  }
  getMaxRows() {
    return Math.max(this.getLastRow(), 500);
  }
  getRange(arg1, arg2, arg3, arg4) {
    if (typeof arg1 === 'string') {
      if (arg1 === 'C2:C3') return new MockRange(this, 2, 3, 2, 1);
      if (arg1 === 'C2') return new MockRange(this, 2, 3, 1, 1);
      if (arg1 === 'B2') return new MockRange(this, 2, 2, 1, 1);
      if (arg1 === 'B3') return new MockRange(this, 3, 2, 1, 1);
      return new MockRange(this, 1, 1, 1, 1);
    }
    return new MockRange(this, arg1, arg2, arg3 || 1, arg4 || 1);
  }
  setFrozenRows() {}
  setFrozenColumns() {}
  setColumnWidth() {}
  getConditionalFormatRules() {
    return this.conditionalRules;
  }
  setConditionalFormatRules(rules) {
    this.conditionalRules = rules;
  }
  clear() {
    this.cells.clear();
    this._noteEvents.length = 0;
  }
}

class MockSpreadsheet {
  constructor() {
    this.sheets = new Map();
    this.toasts = [];
  }
  getSheets() {
    return Array.from(this.sheets.values());
  }
  getSheetByName(name) {
    return this.sheets.get(name) || null;
  }
  insertSheet(name) {
    const s = new MockSheet(name, this);
    this.sheets.set(name, s);
    return s;
  }
  toast(msg, title, timeout) {
    this.toasts.push({ msg, title, timeout });
  }
}

// ============================================================================
// テストサンドボックス構築
// ============================================================================
const state = {
  logs: [],
  httpRequests: [],
  docReplacements: [],
  props: {
    KINTONE_DOMAIN: 'example.cybozu.com',
    KINTONE_APP_ID: '195',
    KINTONE_API_TOKEN: 'dummy-kintone-token',
    REPORT_TEMPLATE_DOC_ID: 'dummy-template-doc-id',
    SLACK_WEBHOOK_URL: 'https://hooks.slack.com/services/T000/B000/DUMMY',
    LW_CLIENT_ID: 'dummy-lw-client-id',
    LW_CLIENT_SECRET: 'dummy-lw-secret',
    LW_SERVICE_ACCOUNT: 'dummy-sa@example.com',
    LW_PRIVATE_KEY: 'dummy-private-key',
    LW_BOT_ID: '99999',
    LW_REMINDER_TARGET_ID: 'channel-dummy-001',
  },
  failTemplateFetch: false,
  lwMessageHttpStatus: 200,
  lwMessageResponseBody: '{"status":"ok"}',
  kintoneHttpStatus: 200,
  kintoneResponseBody: null,
  activeSpreadsheet: new MockSpreadsheet(),
};

const sandbox = {
  Object,
  Array,
  String,
  Number,
  Boolean,
  Date,
  Math,
  RegExp,
  Set,
  Map,
  Error,
  JSON,
  parseInt,
  isNaN,
  Logger: {
    log: (msg) => state.logs.push(String(msg)),
  },
  PropertiesService: {
    getScriptProperties: () => ({
      getProperty: (k) => (state.props[k] !== undefined ? state.props[k] : ''),
      getProperties: () => ({ ...state.props }),
      setProperty: (k, v) => {
        state.props[k] = v;
      },
    }),
  },
  LockService: {
    getScriptLock: () => ({
      tryLock: () => true,
      releaseLock: () => {},
    }),
  },
  Utilities: {
    formatDate: (d, tz, fmt) => {
      const dt = d instanceof Date && !isNaN(d.getTime()) ? d : new Date('2026-10-31T09:00:00+09:00');
      const yyyy = dt.getFullYear();
      const mm = String(dt.getMonth() + 1).padStart(2, '0');
      const dd = String(dt.getDate()).padStart(2, '0');
      if (fmt === 'yyyy/MM/dd') return `${yyyy}/${mm}/${dd}`;
      if (fmt === 'yyyy-MM-dd HH:mm:ss') return `${yyyy}-${mm}-${dd} 09:00:00`;
      return `${yyyy}-${mm}-${dd}`;
    },
    base64Encode: () => 'ZHVtbXk=',
    base64EncodeWebSafe: () => 'dummysig',
    computeRsaSha256Signature: () => [1, 2, 3, 4],
    Charset: { UTF_8: 'UTF-8' },
  },
  DriveApp: {
    getFileById: (id) => {
      if (state.failTemplateFetch) {
        const err = new Error(`Template ID ${id} could not be opened`);
        err.stack = `Error: Template ID ${id} could not be opened\n    at generateFirstReportDoc (DocGenerator.js:74:28)`;
        throw err;
      }
      return {
        makeCopy: (title) => ({
          getId: () => 'copied-doc-001',
          getUrl: () => 'https://docs.google.com/document/d/copied-doc-001/edit',
          getName: () => title,
        }),
      };
    },
    getFolderById: () => ({}),
  },
  DocumentApp: {
    openById: () => ({
      getBody: () => ({
        replaceText: (pat, val) => state.docReplacements.push({ pat, val }),
      }),
      getHeader: () => null,
      getFooter: () => null,
      saveAndClose: () => {},
    }),
  },
  UrlFetchApp: {
    fetch: (url, options) => {
      const parsedPayload =
        options && options.payload && typeof options.payload === 'string'
          ? (() => {
              try {
                return JSON.parse(options.payload);
              } catch (_) {
                return options.payload;
              }
            })()
          : options && options.payload;

      state.httpRequests.push({
        url,
        options,
        parsedPayload,
      });

      // Slack Webhook
      if (url.includes('hooks.slack.com')) {
        return {
          getResponseCode: () => 200,
          getContentText: () => 'ok',
        };
      }
      // LINE WORKS OAuth Token
      if (url.includes('auth.worksmobile.com')) {
        return {
          getResponseCode: () => 200,
          getContentText: () => JSON.stringify({ access_token: 'dummy-lw-access-token' }),
        };
      }
      // LINE WORKS Bot Message
      if (url.includes('worksapis.com')) {
        return {
          getResponseCode: () => state.lwMessageHttpStatus,
          getContentText: () => state.lwMessageResponseBody,
        };
      }
      // kintone REST API
      if (url.includes('/k/v1/records.json')) {
        state.kintoneCallCounter = (state.kintoneCallCounter || 0) + 1;
        if (
          state.failOnKintoneCallNumber &&
          state.kintoneCallCounter === state.failOnKintoneCallNumber
        ) {
          return {
            getResponseCode: () => 500,
            getContentText: () => '{"code":"GAIA_SV01","message":"chunk error"}',
          };
        }
        if (state.kintoneHttpStatus && state.kintoneHttpStatus !== 200) {
          return {
            getResponseCode: () => state.kintoneHttpStatus,
            getContentText: () => state.kintoneResponseBody || '{"code":"GAIA_SV01","message":"error"}',
          };
        }
        const count =
          parsedPayload && Array.isArray(parsedPayload.records) ? parsedPayload.records.length : 1;
        return {
          getResponseCode: () => 200,
          getContentText: () =>
            JSON.stringify({
              ids: Array.from({ length: count }, (_, i) => String(9001 + i)),
              revisions: Array.from({ length: count }, () => '1'),
            }),
        };
      }
      return {
        getResponseCode: () => 200,
        getContentText: () => '{}',
      };
    },
  },
  SpreadsheetApp: {
    BorderStyle: { SOLID: 'SOLID', SOLID_MEDIUM: 'SOLID_MEDIUM' },
    newDataValidation: () => {
      let listValues = [];
      return {
        requireValueInList: (vals) => {
          listValues = vals;
          return {
            setAllowInvalid: () => ({
              build: () => ({ listValues }),
            }),
          };
        },
      };
    },
    newConditionalFormatRule: () => {
      let textVal = '';
      let bgVal = '';
      let ruleRanges = [];
      return {
        whenTextEqualTo: (t) => {
          textVal = t;
          return {
            setBackground: (bg) => {
              bgVal = bg;
              return {
                setRanges: (r) => {
                  ruleRanges = r;
                  return {
                    build: () => ({
                      textVal,
                      bgVal,
                      getRanges: () => ruleRanges,
                    }),
                  };
                },
              };
            },
          };
        },
      };
    },
    getActiveSpreadsheet: () => state.activeSpreadsheet,
  },
};

vm.createContext(sandbox);

const MODULE_ORDER = [
  'Config.js',
  'LedgerMatcher.js',
  'DocGenerator.js',
  'KintoneClient.js',
  'ReminderService.js',
  'PastDataMigration.js',
  'Code.js',
];

MODULE_ORDER.forEach((file) => {
  const code = fs.readFileSync(path.join(SRC_DIR, file), 'utf8');
  vm.runInContext(code, sandbox, { filename: file });
});

const results = [];
function recordPass(category, description) {
  results.push({ category, description, status: 'PASS' });
  console.log(`✅ [PASS] [${category}] ${description}`);
}

// ============================================================================
// 1. タスク1の自動検証（テンプレート取得失敗時のPHI非出力）
// ============================================================================
{
  state.logs.length = 0;
  state.httpRequests.length = 0;
  state.failTemplateFetch = true;
  state.props.SLACK_WEBHOOK_URL = 'https://hooks.slack.com/services/T000/B000/DUMMY';

  assert.throws(() => {
    sandbox.generateFirstReportDoc(
      {
        visitDate: '2026-10-31',
        patientId: 'P00123',
        patientName: '山田太郎',
        facilityName: '○○中央病院',
        colNumber: 3,
      },
      []
    );
  });

  const slackReqs = state.httpRequests.filter((r) => r.url.includes('hooks.slack.com'));
  assert.strictEqual(slackReqs.length, 1, 'Slack通知が1回送信されること');
  const slackBodyStr = JSON.stringify(slackReqs[0].parsedPayload);
  const attachment = slackReqs[0].parsedPayload.attachments[0];

  assert(!slackBodyStr.includes('山田太郎'), 'Slack通知ペイロードに患者氏名「山田太郎」が1文字も含まれないこと');
  assert(
    !attachment.text.includes('at generateFirstReportDoc') && !attachment.text.includes('DocGenerator.js'),
    'Slack通知ペイロードにスタックトレースが含まれないこと'
  );
  recordPass(
    'タスク1',
    'generateFirstReportDoc テンプレート取得エラー時、Slack通知に患者氏名（「山田太郎」）およびスタックトレースが1文字も含まれないこと'
  );

  const fieldMap = {};
  attachment.fields.forEach((f) => {
    fieldMap[f.title] = f.value;
  });
  assert.deepStrictEqual(
    Object.keys(fieldMap).sort(),
    ['column', 'patientId', 'templateDocId'],
    'メタ情報が templateDocId, patientId, column のみであること'
  );
  assert.strictEqual(fieldMap.templateDocId, 'dummy-template-doc-id');
  assert.strictEqual(fieldMap.patientId, 'P00123');
  assert.strictEqual(fieldMap.column, '3');

  // Webhook未設定時の代替ログ検証
  state.logs.length = 0;
  state.props.SLACK_WEBHOOK_URL = '';
  assert.throws(() => {
    sandbox.generateFirstReportDoc(
      {
        visitDate: '2026-10-31',
        patientId: 'P00123',
        patientName: '山田太郎',
        facilityName: '○○中央病院',
        colNumber: 3,
      },
      []
    );
  });
  assert.strictEqual(state.logs.length, 1);
  assert(!state.logs[0].includes('山田太郎'), 'Webhook未設定時の代替ログに患者氏名が含まれないこと');
  assert(
    !state.logs[0].includes('DocGenerator.js'),
    'Webhook未設定時の代替ログにスタックトレースや本文が含まれないこと'
  );
  assert(
    state.logs[0].includes('[ERROR]') &&
      state.logs[0].includes('致命的エラー発生: 第一報テンプレート取得失敗') &&
      /length:\s*\d+/.test(state.logs[0]),
    '代替ログにはタイトル・レベル・文字数のみが出力されること'
  );
  recordPass(
    'タスク1',
    'メタ情報として templateDocId, patientId, column のみが出力され、Webhook未設定時の代替ログにはタイトル・レベル・文字数のみが出力されること'
  );

  state.failTemplateFetch = false;
  state.props.SLACK_WEBHOOK_URL = 'https://hooks.slack.com/services/T000/B000/DUMMY';
}

// ============================================================================
// 2. タスク2の自動検証（未返書リマインドのSlack・ログPHI非出力）
// ============================================================================
{
  state.activeSpreadsheet = new MockSpreadsheet();
  sandbox.setupAllSheets();

  // 当年度台帳に未返書の前方紹介ダミーデータを2件投入（14日以上経過1件、7〜13日経過1件）
  const ledgerSheetName = sandbox.getLedgerSheetName();
  const ledgerSheet = state.activeSpreadsheet.getSheetByName(ledgerSheetName);
  ledgerSheet.getRange(2, 1, 2, 18).setValues([
    [
      '101',
      '2026-10-10 09:00:00',
      '2026-10-10', // 21日経過 (🚨)
      'P00123',
      '山田太郎',
      '前方（紹介・受入）',
      '入院',
      'TEST-01',
      '○○中央病院',
      '内科',
      '佐藤一郎',
      '内科',
      '溝上太郎',
      '未作成',
      '',
      '',
      '送信完了',
      '',
    ],
    [
      '102',
      '2026-10-22 09:00:00',
      '2026-10-22', // 9日経過 (⚠️)
      'P00456',
      '鈴木花子',
      '前方（紹介・受入）',
      '外来',
      'TEST-02',
      '△△中央病院',
      '外科',
      '高橋次郎',
      '外科',
      '溝上太郎',
      '未作成',
      '',
      '',
      '不要',
      '',
    ],
  ]);

  state.logs.length = 0;
  state.httpRequests.length = 0;
  state.lwMessageHttpStatus = 200;
  state.lwMessageResponseBody = '{"status":"ok"}';

  sandbox.sendUnrepliedRemindersManual();

  // LINE WORKS送信ペイロードには詳細な一覧が含まれること
  const lwReqs = state.httpRequests.filter((r) => r.url.includes('worksapis.com'));
  assert.strictEqual(lwReqs.length, 1, 'LINE WORKSへメッセージが1回送信されること');
  const lwText = lwReqs[0].parsedPayload.content.text;
  assert(lwText.includes('山田太郎') && lwText.includes('P00123') && lwText.includes('○○中央病院'), 'LINE WORKSには詳細一覧が送信されること');

  // Slack通知には件数サマリーのみが出力され、PHIが一切含まれないこと
  const slackReqs = state.httpRequests.filter((r) => r.url.includes('hooks.slack.com'));
  assert.strictEqual(slackReqs.length, 1, 'Slackへリマインド通知が1回送信されること');
  const slackPayloadStr = JSON.stringify(slackReqs[0].parsedPayload);
  const slackAttachment = slackReqs[0].parsedPayload.attachments[0];

  assert(
    /返書未作成\s*\d+件（うち期限超過\s*\d+件）。詳細はLINE WORKSを確認/.test(slackAttachment.text),
    'Slack通知本文が「返書未作成 N件（うち期限超過 M件）。詳細はLINE WORKSを確認」形式であること'
  );
  ['山田太郎', '鈴木花子', 'P00123', 'P00456', '○○中央病院', '△△中央病院', '佐藤一郎', '高橋次郎', '溝上太郎'].forEach(
    (phiToken) => {
      assert(!slackPayloadStr.includes(phiToken), `Slack通知に「${phiToken}」が含まれないこと`);
      state.logs.forEach((logLine) => {
        assert(!logLine.includes(phiToken), `Logger.log に「${phiToken}」が含まれないこと`);
      });
    }
  );
  recordPass(
    'タスク2',
    '未返書リマインド実行時、LINE WORKSには詳細一覧が送られ、SlackおよびLogger.logには件数サマリーのみが出力されPHI（氏名・ID・施設名・医師名）が一切含まれないこと'
  );

  // LINE WORKS APIエラー（HTTP 400 / 500）時にレスポンス本文がSlackに出ないこと
  state.httpRequests.length = 0;
  state.lwMessageHttpStatus = 400;
  state.lwMessageResponseBody = JSON.stringify({
    code: 'INVALID_PARAMETER',
    description: '山田太郎の秘密レスポンス本文エラー詳細',
  });

  sandbox.executeUnrepliedReminderNotification();
  const errSlackReqs = state.httpRequests.filter(
    (r) =>
      r.url.includes('hooks.slack.com') &&
      JSON.stringify(r.parsedPayload).includes('LINE WORKS Bot送信エラー')
  );
  assert.strictEqual(errSlackReqs.length, 1, 'LINE WORKSエラー時にSlackエラー通知が送られること');
  const errSlackStr = JSON.stringify(errSlackReqs[0].parsedPayload);
  assert(
    !errSlackStr.includes('山田太郎の秘密レスポンス本文エラー詳細'),
    'Slack通知にLINE WORKSのエラーレスポンス本文が含まれないこと'
  );
  assert(
    errSlackStr.includes('400') && errSlackStr.includes('INVALID_PARAMETER'),
    'Slack通知にHTTPステータスコードとエラーコードのみが含まれること'
  );
  recordPass(
    'タスク2',
    'LINE WORKS APIエラー時、Slack通知にレスポンス本文が含まれず、HTTPステータスとエラーコードのみが通知されること'
  );

  state.lwMessageHttpStatus = 200;
  state.lwMessageResponseBody = '{"status":"ok"}';
}

// ============================================================================
// 3. タスク3の自動検証（施設名の部分一致による誤FAX防止とメモ自動消去）
// ============================================================================
{
  state.activeSpreadsheet = new MockSpreadsheet();
  sandbox.setupAllSheets();

  // 施設マスターシートにダミーデータ2件（「○○中央病院」「△△中央病院」）を登録
  const masterSheet = state.activeSpreadsheet.getSheetByName('施設マスター');
  masterSheet.getRange(2, 1, 2, 7).setValues([
    ['TEST-01', '○○中央病院', '東京都千代田区1-1', '03-0000-0001', '', '', ''],
    ['TEST-02', '△△中央病院', '東京都中央区2-2', '03-0000-0002', '', '', ''],
  ]);

  const masterList = sandbox.loadFacilityMasterList(state.activeSpreadsheet);

  // 3-1. resolveFacility('中央病院') の戻り値検証
  const partialRes = sandbox.resolveFacility('中央病院', masterList);
  assert.strictEqual(partialRes.isMatched, false, '部分一致時は isMatched: false であること');
  assert.strictEqual(partialRes.matchType, 'partial', '部分一致時は matchType: "partial" であること');
  assert.strictEqual(partialRes.lookupKey, '', '部分一致時は lookupKey が空文字であること');
  assert.strictEqual(partialRes.faxNumber, '', '部分一致時は faxNumber が空文字であること');
  assert.strictEqual(partialRes.officialName, '中央病院', '部分一致時は officialName が入力値のままであること');
  assert.strictEqual(partialRes.candidates.length, 2, 'candidates に2件格納されていること');
  assert.deepStrictEqual(
    Array.from(partialRes.candidates.map((c) => c.officialName)),
    ['○○中央病院', '△△中央病院']
  );
  recordPass(
    'タスク3-1',
    'resolveFacility("中央病院") が isMatched: false, matchType: "partial", lookupKey: "", faxNumber: "" を返し、candidates に「○○中央病院」「△△中央病院」の2件が入ること'
  );

  // 3-2. expandJsonAndMatchLedger / executeBatchProcess / reRunFiveLockMatchingForAllColumns の3か所検証
  const inputSheet = state.activeSpreadsheet.getSheetByName('紹介・逆紹介入力');
  const expectedPartialNote =
    '施設名が部分一致のみ: 候補 ○○中央病院／△△中央病院。マスター別名に登録するか正式名称を入力してください';

  // (a) expandJsonAndMatchLedger
  sandbox.expandJsonAndMatchLedger(
    JSON.stringify([
      {
        受診日: '2026-10-31',
        患者ID: 'P00999',
        患者氏名: '山田太郎',
        連携フェーズ: '前方（紹介・受入）',
        受入形態: '外来',
        連携先施設名: '中央病院',
        連携先医師名: '佐藤一郎',
        当院担当診療科: '内科',
        担当医師: '溝上太郎',
      },
    ]),
    state.activeSpreadsheet
  );

  const facilityCellC10 = inputSheet.getRange(10, 3); // C10 (10行目・C列)
  assert.strictEqual(
    facilityCellC10.getValue(),
    '中央病院',
    'expandJsonAndMatchLedger で10行目が勝手に上書きされないこと'
  );
  assert.strictEqual(
    facilityCellC10.getNote(),
    expectedPartialNote,
    'expandJsonAndMatchLedger で10行目に部分一致警告メモが付与されること'
  );

  // (b) reRunFiveLockMatchingForAllColumns
  facilityCellC10.clearNote();
  sandbox.reRunFiveLockMatchingForAllColumns();
  assert.strictEqual(
    facilityCellC10.getValue(),
    '中央病院',
    'reRunFiveLockMatchingForAllColumns で10行目が勝手に上書きされないこと'
  );
  assert.strictEqual(
    facilityCellC10.getNote(),
    expectedPartialNote,
    'reRunFiveLockMatchingForAllColumns で10行目に部分一致警告メモが付与されること'
  );

  // (c) executeBatchProcess
  facilityCellC10.clearNote();
  sandbox.executeBatchProcess();
  assert.strictEqual(
    facilityCellC10.getValue(),
    '中央病院',
    'executeBatchProcess で10行目が勝手に上書きされないこと'
  );
  assert.strictEqual(
    facilityCellC10.getNote(),
    expectedPartialNote,
    'executeBatchProcess で10行目に部分一致警告メモが付与されること'
  );
  recordPass(
    'タスク3-2',
    'expandJsonAndMatchLedger, executeBatchProcess, reRunFiveLockMatchingForAllColumns の3か所で10行目が上書きされず警告メモが付くこと'
  );

  // 3-3. スタッフが10行目を「○○中央病院」に修正して再実行した際、警告メモが自動消去（clearNote）されること
  facilityCellC10.setValue('○○中央病院');
  assert.strictEqual(facilityCellC10.getNote(), expectedPartialNote, '修正前は警告メモが残っていること');

  sandbox.reRunFiveLockMatchingForAllColumns();
  assert.strictEqual(
    facilityCellC10.getNote(),
    '',
    'reRunFiveLockMatchingForAllColumns 再実行後、exact一致により10行目の警告メモが自動消去(clearNote)されること'
  );

  // executeBatchProcess でも clearNote されることを確認
  facilityCellC10.setNote(expectedPartialNote);
  inputSheet.getRange(4, 3).setValue('⏳ 未連携'); // 再POST対象に戻す
  sandbox.executeBatchProcess();
  assert.strictEqual(
    facilityCellC10.getNote(),
    '',
    'executeBatchProcess 再実行後、exact一致により10行目の警告メモが自動消去(clearNote)されること'
  );
  recordPass(
    'タスク3-3',
    '入力値を「○○中央病院」に修正して再実行すると matchType が "exact" になり、10行目の警告メモが自動で消去（clearNote）されること'
  );

  // 3-4. generateFirstReportDoc での {{連携先FAX番号}} プレースホルダー検証
  state.docReplacements.length = 0;
  sandbox.generateFirstReportDoc(
    {
      visitDate: '2026-10-31',
      patientId: 'P00999',
      patientName: '山田太郎',
      facilityName: '中央病院', // partial
    },
    masterList
  );
  const partialFaxRep = state.docReplacements.find((r) => r.pat.includes('連携先FAX番号'));
  assert.strictEqual(
    partialFaxRep.val,
    '（要確認: 施設マスター未一致）',
    'matchType === "partial" のとき {{連携先FAX番号}} が「（要確認: 施設マスター未一致）」になること'
  );

  state.docReplacements.length = 0;
  sandbox.generateFirstReportDoc(
    {
      visitDate: '2026-10-31',
      patientId: 'P00999',
      patientName: '山田太郎',
      facilityName: '未登録クリニック', // none
    },
    masterList
  );
  const noneFaxRep = state.docReplacements.find((r) => r.pat.includes('連携先FAX番号'));
  assert.strictEqual(
    noneFaxRep.val,
    '（要確認: 施設マスター未一致）',
    'matchType === "none" のとき {{連携先FAX番号}} が「（要確認: 施設マスター未一致）」になること'
  );

  state.docReplacements.length = 0;
  sandbox.generateFirstReportDoc(
    {
      visitDate: '2026-10-31',
      patientId: 'P00999',
      patientName: '山田太郎',
      facilityName: '○○中央病院', // exact
    },
    masterList
  );
  const exactFaxRep = state.docReplacements.find((r) => r.pat.includes('連携先FAX番号'));
  assert.strictEqual(
    exactFaxRep.val,
    '03-0000-0001',
    'matchType === "exact" のときはマスターのFAX番号（03-0000-0001）が設定されること'
  );
  recordPass(
    'タスク3-4',
    'matchType が "partial" または "none" の状態で generateFirstReportDoc を実行した際、{{連携先FAX番号}} が「（要確認: 施設マスター未一致）」に置換され他施設のFAX番号が混入しないこと'
  );

  // 3-5. isLock2FacilityMatched で部分一致が一致と判定されないこと
  const lock2Partial = sandbox.isLock2FacilityMatched(
    { lookupKey: '', officialName: '中央病院' },
    'TEST-01',
    '○○中央病院',
    masterList
  );
  assert.strictEqual(lock2Partial, false, 'isLock2FacilityMatched で部分一致が false になること');

  const lock2Core = sandbox.isLock2FacilityMatched(
    { lookupKey: '', officialName: '医療法人 ○○中央病院' },
    'TEST-01',
    '○○中央病院',
    masterList
  );
  assert.strictEqual(lock2Core, true, 'isLock2FacilityMatched でコア名称の完全一致が true になること');
  recordPass(
    'タスク3-5',
    'isLock2FacilityMatched において単なる部分一致（includes）が一致（true）と判定されないこと'
  );
}

// ============================================================================
// 4. 返書ステータス拡張・ロック4・リマインドカットオフ・kintone送信・Slack High回帰検証
// ============================================================================
{
  // 4-1. ロック4 (isLock4UnrepliedForward) の論理修正検証
  const santeConfig = vm.runInContext('SANTE_CONFIG', sandbox);
  const forwardPhase = santeConfig.PHASES.FORWARD;
  assert.strictEqual(
    sandbox.isLock4UnrepliedForward(forwardPhase, '返書不要（医師判断）', ''),
    false,
    '返書日空 + ステータス「返書不要（医師判断）」→ 未返書でない(false)'
  );
  assert.strictEqual(
    sandbox.isLock4UnrepliedForward(forwardPhase, '作成・送付完了', ''),
    false,
    '返書日空 + ステータス「作成・送付完了」→ 未返書でない(false)'
  );
  assert.strictEqual(
    sandbox.isLock4UnrepliedForward(forwardPhase, '対象外', ''),
    false,
    '返書日空 + ステータス「対象外」→ 未返書でない(false)'
  );
  assert.strictEqual(
    sandbox.isLock4UnrepliedForward(forwardPhase, '未作成', ''),
    true,
    '返書日空 + ステータス「未作成」→ 未返書である(true)'
  );
  assert.strictEqual(
    sandbox.isLock4UnrepliedForward(forwardPhase, '', ''),
    true,
    '返書日空 + 空ステータス → 未返書である(true)'
  );
  assert.strictEqual(
    sandbox.isLock4UnrepliedForward(forwardPhase, '未作成', '2026-10-30'),
    false,
    '返書日あり + ステータス「未作成」→ 未返書でない(false)'
  );
  recordPass(
    '改修① ロック4',
    'isLock4UnrepliedForward で返書日空でも「返書不要（医師判断）」「作成・送付完了」「対象外」は false、「未作成」または空のみ true となること'
  );

  // 4-2. リマインド抽出 (extractUnrepliedReferrals) の共通化とカットオフ除外検証
  state.activeSpreadsheet = new MockSpreadsheet();
  sandbox.setupAllSheets();
  const ledgerSheetName = sandbox.getLedgerSheetName();
  const ledgerSheet = state.activeSpreadsheet.getSheetByName(ledgerSheetName);
  ledgerSheet.getRange(2, 1, 4, 18).setValues([
    [
      '201',
      '2026-10-01 09:00:00',
      '2026-10-01', // カットオフ(2026-10-15)より前
      'P00001',
      '山田太郎',
      '前方（紹介・受入）',
      '外来',
      'TEST-01',
      '○○中央病院',
      '内科',
      '佐藤一郎',
      '内科',
      '溝上太郎',
      '未作成',
      '',
      '',
      '不要',
      '',
    ],
    [
      '202',
      '2026-10-20 09:00:00',
      '2026-10-20', // カットオフ以降だが「返書不要（医師判断）」
      'P00002',
      '鈴木花子',
      '前方（紹介・受入）',
      '外来',
      'TEST-01',
      '○○中央病院',
      '内科',
      '佐藤一郎',
      '内科',
      '溝上太郎',
      '返書不要（医師判断）',
      '',
      '',
      '不要',
      '',
    ],
    [
      '203',
      '2026-10-21 09:00:00',
      '2026-10-21', // カットオフ以降だが返書日空の「作成・送付完了」
      'P00003',
      '佐藤健太',
      '前方（紹介・受入）',
      '外来',
      'TEST-01',
      '○○中央病院',
      '内科',
      '佐藤一郎',
      '内科',
      '溝上太郎',
      '作成・送付完了',
      '',
      '',
      '不要',
      '',
    ],
    [
      '204',
      '2026-10-22 09:00:00',
      '2026-10-22', // カットオフ以降 かつ「未作成」→ 抽出対象
      'P00004',
      '高橋美咲',
      '前方（紹介・受入）',
      '外来',
      'TEST-01',
      '○○中央病院',
      '内科',
      '佐藤一郎',
      '内科',
      '溝上太郎',
      '未作成',
      '',
      '',
      '不要',
      '',
    ],
  ]);

  // 全期間（デフォルト REPLY_TRACKING_START_DATE = ''）: 201と204の2件のみ抽出され、202(返書不要)と203(作成・送付完了)は除外
  const fullPeriodReport = sandbox.extractUnrepliedReferrals(
    new Date('2026-10-31T09:00:00+09:00'),
    state.activeSpreadsheet
  );
  assert.strictEqual(
    fullPeriodReport.totalCount,
    2,
    '全期間抽出で「返書不要（医師判断）」と「作成・送付完了（返書日空）」が除外され2件になること'
  );

  // カットオフ指定 ('2026-10-15'): 201(2026-10-01受診)も除外され、204の1件のみになること
  const cutoffReport = sandbox.extractUnrepliedReferrals(
    new Date('2026-10-31T09:00:00+09:00'),
    state.activeSpreadsheet,
    '2026-10-15'
  );
  assert.strictEqual(
    cutoffReport.totalCount,
    1,
    'カットオフ開始日(2026-10-15)指定時、それより前の受診日(2026-10-01)が除外され1件になること'
  );
  assert.strictEqual(cutoffReport.groupedByDoctor[0].items[0].kintoneRecordId, '204');
  recordPass(
    '改修② リマインド抽出・カットオフ',
    'extractUnrepliedReferrals で「返書不要（医師判断）」およびカットオフ日前のレコードが確実に除外されること'
  );

  // 4-3. kintone POST ペイロード整理（ルックアップ時の施設名除外 ＆ 紹介先診療科5択正規化）
  const payloadWithLookup = sandbox.buildKintonePostRecordPayload({
    visitDate: '2026-10-31',
    patientId: 'P00100',
    patientName: '山田太郎',
    phase: '前方（紹介・受入）',
    visitType: '外来',
    lookupKey: 'TEST-01',
    facilityName: '○○中央病院',
    facilityDept: '整形外科',
    facilityDoctor: '佐藤一郎',
    ourDept: '整形外科',
    ourDoctor: '溝上太郎',
    replyStatus: '未作成',
  });
  assert.strictEqual(payloadWithLookup['更新用キー'].value, 'TEST-01');
  assert.strictEqual(
    payloadWithLookup['連携先施設名'],
    undefined,
    'lookupKey があるときは連携先施設名をペイロードに含めないこと'
  );
  assert.strictEqual(
    payloadWithLookup['紹介先診療科'].value,
    '整形外科',
    '5択に合致する診療科は紹介先診療科フィールドにセットされること'
  );

  const payloadWithoutLookup = sandbox.buildKintonePostRecordPayload({
    visitDate: '2026-10-31',
    patientId: 'P00101',
    patientName: '山田太郎',
    phase: '前方（紹介・受入）',
    visitType: '外来',
    lookupKey: '',
    facilityName: '未登録クリニック',
    facilityDept: '皮膚科', // 5択外
    facilityDoctor: '佐藤一郎',
    ourDept: '内科',
    ourDoctor: '溝上太郎',
    replyStatus: '未作成',
  });
  assert.strictEqual(
    payloadWithoutLookup['更新用キー'],
    undefined,
    'lookupKey が空のときは更新用キーを含めないこと'
  );
  assert.strictEqual(
    payloadWithoutLookup['連携先施設名'].value,
    '未登録クリニック',
    'lookupKey が空のときは自由入力として連携先施設名をセットすること'
  );
  assert.strictEqual(
    payloadWithoutLookup['連携先診療科'].value,
    '皮膚科',
    '自由テキストの連携先診療科は常にセットされること'
  );
  assert.strictEqual(
    payloadWithoutLookup['紹介先診療科'],
    undefined,
    '5択外の診療科（皮膚科）は紹介先診療科にセットせず GAIA_FO01 を起こさないこと'
  );
  recordPass(
    '改修③ kintone POST整理',
    'lookupKeyあり時の施設名除外、lookupKey空時の施設名送信、および紹介先診療科5択の安全な正規化セットが機能すること'
  );

  // 4-4. 台帳N列のステータス定数・入力規則・条件付き書式検証
  assert.strictEqual(
    santeConfig.REPLY_STATUS.NOT_REQUIRED,
    '返書不要（医師判断）',
    'SANTE_CONFIG.REPLY_STATUS.NOT_REQUIRED が定義されていること'
  );
  const rules = ledgerSheet.getConditionalFormatRules();
  const bgByText = {};
  rules.forEach((r) => {
    bgByText[r.textVal] = r.bgVal;
  });
  assert.strictEqual(bgByText['未作成'], '#F4CCCC');
  assert.strictEqual(bgByText['作成・送付完了'], '#D9EAD3');
  assert.strictEqual(bgByText['返書不要（医師判断）'], '#E6E6E6');
  assert(
    ledgerSheet._lastDataValidation &&
      Array.from(ledgerSheet._lastDataValidation.listValues).includes('返書不要（医師判断）'),
    '台帳N列の入力規則プルダウンに「返書不要（医師判断）」が含まれること'
  );
  recordPass(
    '改修④ ステータス定数と台帳UI',
    'REPLY_STATUS.NOT_REQUIRED 定数、台帳N列プルダウン入力規則、および条件付き書式（#E6E6E6）が設定されること'
  );

  // 4-5. Slack High-1 & High-2: kintone APIエラー時および executeBatchProcess catch 時に本文・PHI・スタックがSlackに出ないこと
  state.httpRequests.length = 0;
  state.logs.length = 0;
  state.kintoneHttpStatus = 400;
  state.kintoneResponseBody = JSON.stringify({
    id: 'ERR-999',
    code: 'GAIA_SV01',
    message: '山田太郎のレコード登録でエラー詳細メッセージ本文が返却されました',
    errors: {
      'records[0].患者氏名.value': {
        messages: ['山田太郎の氏名フィールドエラー詳細'],
      },
    },
  });

  // 入力シートに未連携データを1列用意して executeBatchProcess を実行（kintone POST失敗 → catch）
  const inputSheet = state.activeSpreadsheet.getSheetByName('紹介・逆紹介入力');
  sandbox.expandJsonAndMatchLedger(
    JSON.stringify([
      {
        受診日: '2026-10-31',
        患者ID: 'P00777',
        患者氏名: '山田太郎',
        連携フェーズ: '前方（紹介・受入）',
        受入形態: '外来',
        連携先施設名: '○○中央病院',
        連携先医師名: '佐藤一郎',
        当院担当診療科: '内科',
        担当医師: '溝上太郎',
      },
    ]),
    state.activeSpreadsheet
  );

  let caughtBatchErr = null;
  try {
    sandbox.executeBatchProcess();
  } catch (err) {
    caughtBatchErr = err;
  }

  assert(caughtBatchErr !== null, 'kintoneエラー時に例外がスローされること');
  assert(
    !caughtBatchErr.message.includes('山田太郎'),
    'throw された Error.message にレスポンス本文や患者氏名が含まれないこと'
  );
  assert(
    caughtBatchErr.message.includes('HTTP 400') && caughtBatchErr.message.includes('GAIA_SV01'),
    'throw された Error.message にHTTPステータスとエラーコードが含まれること'
  );

  const kintoneErrSlackReqs = state.httpRequests.filter((r) => r.url.includes('hooks.slack.com'));
  assert(kintoneErrSlackReqs.length >= 2, 'KintoneClient と executeBatchProcess catch のSlack通知が送信されること');
  kintoneErrSlackReqs.forEach((req) => {
    const bodyStr = JSON.stringify(req.parsedPayload);
    assert(!bodyStr.includes('山田太郎'), 'Slack通知ペイロードに患者氏名「山田太郎」やレスポンス本文が含まれないこと');
    assert(!bodyStr.includes('ERR-999'), 'Slack通知ペイロードに生レスポンスJSONが含まれないこと');
    assert(!bodyStr.includes('at sendKintoneRequestWithAutoRecovery'), 'Slack通知ペイロードにスタックトレースが含まれないこと');
  });

  state.kintoneHttpStatus = 200;
  state.kintoneResponseBody = null;
  recordPass(
    'Slack High-1 / High-2',
    'kintone APIエラー時および executeBatchProcess 異常終了時、Slack通知と例外メッセージにレスポンス本文・患者氏名・スタックトレースが一切含まれないこと'
  );
}

// ============================================================================
// 5. 当院担当医11名化・返書不要初期自動判定・自動クリア・過去データ移行検証
// ============================================================================
{
  const santeConfig = vm.runInContext('SANTE_CONFIG', sandbox);

  // 5-1. タスク1: normalizeOurDeptAndDoctor (現役11名通過、退職医2名・選択肢外の備考退避)
  const expected11Doctors = [
    '溝上泰一朗',
    '上床武史',
    '尤郁偉',
    '田中淳',
    '奈田慎一',
    '村山賢一郎',
    '大江健介',
    '吉原智仁',
    '劉軒',
    '泉政寛',
    '佐々木威治',
  ];
  assert.deepStrictEqual(
    Array.from(santeConfig.KINTONE.OUR_DOCTOR_OPTIONS),
    expected11Doctors,
    'OUR_DOCTOR_OPTIONS が現役11名と完全一致し、退職医2名（野田明生・佐藤恭一）を含まないこと'
  );

  expected11Doctors.forEach((docName) => {
    const res = sandbox.normalizeOurDeptAndDoctor('', docName, '', '');
    assert.strictEqual(res.doctor, docName, `現役医師 ${docName} がそのまま doctor に残ること`);
    assert(
      santeConfig.KINTONE.OUR_DEPT_OPTIONS.includes(res.dept),
      `現役医師 ${docName} から有効な診療科 (${res.dept}) が補完されること`
    );
    assert.strictEqual(res.extraNote, '', `現役医師 ${docName} の extraNote が空であること`);
  });

  // 退職医1: 野田明生 → 診療科「整形外科」、医師空、extraNote に「旧/他担当医:野田明生」
  const nodaRes = sandbox.normalizeOurDeptAndDoctor('', '野田明生', '', '');
  assert.strictEqual(nodaRes.dept, '整形外科');
  assert.strictEqual(nodaRes.doctor, '');
  assert.strictEqual(nodaRes.extraNote, '旧/他担当医:野田明生');

  // 退職医2: 佐藤恭一 → 文脈により「内科（一般）」または「循環器内科」、医師空、extraNote に「旧/他担当医:佐藤恭一」
  const satoGeneralRes = sandbox.normalizeOurDeptAndDoctor('内科', '佐藤恭一', '', '');
  assert.strictEqual(satoGeneralRes.dept, '内科（一般）');
  assert.strictEqual(satoGeneralRes.doctor, '');
  assert.strictEqual(satoGeneralRes.extraNote, '旧/他担当医:佐藤恭一');

  const satoCardioRes = sandbox.normalizeOurDeptAndDoctor('循環器内科', '佐藤恭一', '', '');
  assert.strictEqual(satoCardioRes.dept, '循環器内科');
  assert.strictEqual(satoCardioRes.doctor, '');
  assert.strictEqual(satoCardioRes.extraNote, '旧/他担当医:佐藤恭一');

  // buildKintonePostRecordPayload で退職医がドロップダウンに送られず備考へ退避されること
  const retiredPayload = sandbox.buildKintonePostRecordPayload({
    visitDate: '2026-10-31',
    patientId: '12345',
    patientName: '山田太郎',
    phase: '前方（紹介・受入）',
    visitType: '外来',
    lookupKey: 'TEST-01',
    facilityName: '○○中央病院',
    facilityDept: '整形外科',
    facilityDoctor: '佐藤一郎',
    ourDept: '',
    ourDoctor: '野田明生',
    replyStatus: '未作成',
    remarks: '紹介状持参',
  });
  assert.strictEqual(retiredPayload['当院担当診療科'].value, '整形外科');
  assert.strictEqual(
    retiredPayload['担当医師'],
    undefined,
    '退職医の場合、担当医師フィールドがペイロードから省略されること'
  );
  assert(
    retiredPayload['備考・連絡事項'].value.includes('紹介状持参') &&
      retiredPayload['備考・連絡事項'].value.includes('旧/他担当医:野田明生'),
    '既存の備考を消さずに末尾へ旧/他担当医が追記されること'
  );
  recordPass(
    'タスク1 医師11名化・備考退避',
    'normalizeOurDeptAndDoctor で現役11名が通過し、野田明生・佐藤恭一等の退職医は診療科のみ補完して医師フィールド省略＋備考退避されること'
  );

  // 5-2. タスク2: 前方紹介時の「返書不要（医師判断）」初期自動判定
  assert.strictEqual(
    sandbox.decideInitialReplyStatus('前方（紹介・受入）', 'ご本人希望により返書不要'),
    '返書不要（医師判断）'
  );
  assert.strictEqual(
    sandbox.decideInitialReplyStatus('前方（紹介・受入）', '返書は書かない方針'),
    '返書不要（医師判断）'
  );
  assert.strictEqual(
    sandbox.decideInitialReplyStatus('前方（紹介・受入）', '特記事項なし'),
    '未作成'
  );
  assert.strictEqual(
    sandbox.decideInitialReplyStatus('後方（逆紹介・送出）', '返書不要'),
    '対象外'
  );
  recordPass(
    'タスク2 返書不要初期自動判定',
    'decideInitialReplyStatus で前方＋備考「返書不要」「書かない」等が NOT_REQUIRED、通常前方が UNCREATED、後方が NOT_APPLICABLE になること'
  );

  // 5-3. タスク3: 一括実行完了後の入力シート自動クリア
  state.activeSpreadsheet = new MockSpreadsheet();
  sandbox.setupAllSheets();
  const masterSheet = state.activeSpreadsheet.getSheetByName('施設マスター');
  masterSheet.getRange(2, 1, 2, 7).setValues([
    ['TEST-01', '○○中央病院', '東京都千代田区1-1', '03-0000-0001', '', '', ''],
    ['TEST-02', '△△中央病院', '東京都中央区2-2', '03-0000-0002', '', '', ''],
  ]);

  const inputSheet = state.activeSpreadsheet.getSheetByName('紹介・逆紹介入力');
  inputSheet.getRange('C2').setValue('[{"dummy":"json"}]');
  sandbox.expandJsonAndMatchLedger(
    JSON.stringify([
      {
        受診日: '2026-10-31',
        患者ID: '12345',
        患者氏名: '山田太郎',
        連携フェーズ: '前方（紹介・受入）',
        受入形態: '外来',
        連携先施設名: '○○中央病院', // 完全一致 → 同期後自動クリア対象
        連携先医師名: '佐藤一郎',
        当院担当診療科: '脳神経外科',
        担当医師: '溝上泰一朗',
        備考: '返書不要との記載あり',
      },
      {
        受診日: '2026-10-31',
        患者ID: '67890',
        患者氏名: '鈴木花子',
        連携フェーズ: '前方（紹介・受入）',
        受入形態: '外来',
        連携先施設名: '中央病院', // 部分一致保留 → 入力シートに残す
        連携先医師名: '高橋次郎',
        当院担当診療科: '整形外科',
        担当医師: '吉原智仁',
        備考: '',
      },
    ]),
    state.activeSpreadsheet
  );

  const batchRes = sandbox.executeBatchProcess();
  assert.strictEqual(batchRes.kintonePosted, 2, '2件ともkintone POSTされること');
  assert.strictEqual(batchRes.autoClearedCount, 1, '完全一致で同期完了したC列(col 3)の1列のみ自動クリアされること');
  assert.strictEqual(inputSheet.getRange('C2').getValue(), '', 'C2:C3 貼付セルがクリアされること');
  // 完了したC列がクリアされた後、D列にあった部分一致保留（鈴木花子・中央病院）がC列(col 3)へ左詰めされ、D列(col 4)が空になること
  assert.strictEqual(
    inputSheet.getRange(7, 3).getValue(),
    '鈴木花子',
    '部分一致保留だったD列の患者がC列(col 3)へ左詰めされること'
  );
  assert.strictEqual(
    inputSheet.getRange(10, 3).getValue(),
    '中央病院',
    '部分一致保留の施設名がC列(col 3)に保持されること'
  );
  assert.strictEqual(inputSheet.getRange(4, 4).getValue(), '', '左詰め後に空いたD列(col 4)の4行目ステータスがクリアされること');
  assert.strictEqual(inputSheet.getRange(7, 4).getValue(), '', '左詰め後に空いたD列(col 4)の患者氏名がクリアされること');
  assert.strictEqual(inputSheet.getRange(5, 4).getNote(), '', '左詰め後に空いたD列(col 4)のセルNoteがクリアされること');
  assert.strictEqual(inputSheet.getRange(15, 4).getValue(), false, '左詰め後に空いたD列(col 4)の15行目チェックボックスが false になること');
  assert.strictEqual(inputSheet.getRange(17, 4).getValue(), false, '左詰め後に空いたD列(col 4)の17行目チェックボックスが false になること');
  const lastToast = state.activeSpreadsheet.toasts[state.activeSpreadsheet.toasts.length - 1];
  assert(
    lastToast &&
      lastToast.msg.includes('生成した第一報URLや登録内容は年度別連携台帳タブに記録され、入力シートの完了列をクリアしました'),
    '完了トーストに台帳記録・入力シート完了列クリアの案内が含まれること'
  );

  // 台帳側に1件目（返書不要）が NOT_REQUIRED で記録されていることも確認
  const ledgerSheet = state.activeSpreadsheet.getSheetByName(sandbox.getLedgerSheetName());
  assert.strictEqual(
    ledgerSheet.getRange(2, 14).getValue(),
    '返書不要（医師判断）',
    '備考に「返書不要」がある前方紹介が台帳N列に「返書不要（医師判断）」で記録されること'
  );
  recordPass(
    'タスク3 一括実行後の自動クリア',
    'executeBatchProcess 完了後、同期成功列と C2:C3 が自動クリア（チェックボックス false 化・Note消去）され、部分一致保留列は左詰めされて残ること'
  );

  // 5-4. タスク4: 過去データ移行（4段階ステータス判定・シート合体・BOM付CSV・kintone一括POST＆台帳反映）
  // (a) 4段階ステータス純粋関数テスト
  const stBackward = sandbox.classifyPastRecordReplyStatus('', '', '逆紹介送出', '');
  assert.strictEqual(stBackward.phase, '後方（逆紹介・送出）');
  assert.strictEqual(stBackward.replyStatus, '対象外');

  const stCompletedByDate = sandbox.classifyPastRecordReplyStatus('2026-05-10', '', '', '');
  assert.strictEqual(stCompletedByDate.phase, '前方（紹介・受入）');
  assert.strictEqual(stCompletedByDate.replyStatus, '作成・送付完了');

  const stCompletedByDone = sandbox.classifyPastRecordReplyStatus('', '済', '', '');
  assert.strictEqual(stCompletedByDone.replyStatus, '作成・送付完了');

  const stNotRequired = sandbox.classifyPastRecordReplyStatus('', '未', '', '患者キャンセルのため不要');
  assert.strictEqual(stNotRequired.replyStatus, '返書不要（医師判断）');

  const stUncreated = sandbox.classifyPastRecordReplyStatus('', '未', '', '紹介状あり');
  assert.strictEqual(stUncreated.replyStatus, '未作成');

  // (b) 月別シート合体 (buildPastDataMigrationSheet) -> CSV (buildMigrationCsvContent) -> 一括POST (sendPastDataToKintoneAppB)
  const pastMonthSheet = state.activeSpreadsheet.insertSheet('2026年4月');
  const pastRows = [
    [
      '来院日',
      '患者ID',
      '患者氏名',
      '紹介元病院名',
      '紹介元診療科',
      '紹介元医師名',
      '受入・送出形態',
      '当院診療科',
      '担当医師',
      '返書日',
      '済/未',
      '予備',
      '問い合わせ',
      '備考',
    ],
    [
      '2026-04-10',
      '10001',
      '山田太郎',
      '○○中央病院', // マスタ完全一致 -> TEST-01
      '整形外科',     // 5択一致
      '佐藤 一郎',
      '外来',
      '',
      '野田明生',     // 退職医 -> 整形外科 + 備考退避
      '2026-04-15',
      '済',
      '',
      '',
      '',
    ],
    [
      '2026-04-12',
      '10002',
      '鈴木花子',
      '中央病院',     // 部分一致のみ -> lookupKey 空 + 備考退避
      '皮膚科',       // 5択外
      '高橋 次郎',
      '検査',         // EXAM
      '循環器内科',
      '佐藤恭一',     // 退職医 -> 循環器内科 + 備考退避
      '',
      '未',
      '',
      '',
      '返書なし',     // -> 返書不要（医師判断）
    ],
  ];
  pastMonthSheet.getRange(1, 1, pastRows.length, 14).setValues(pastRows);

  const migBuildRes = sandbox.buildPastDataMigrationSheet(state.activeSpreadsheet);
  assert.strictEqual(migBuildRes.totalCount, 2, '月別シートから2件が合体抽出されること');
  assert.strictEqual(migBuildRes.countsByStatus['作成・送付完了'], 1);
  assert.strictEqual(migBuildRes.countsByStatus['返書不要（医師判断）'], 1);

  const csvRes = sandbox.downloadMigrationCsv(state.activeSpreadsheet);
  assert(csvRes.csvContent.startsWith('\uFEFF'), '生成されるCSVがBOM（\\uFEFF）で始まること');
  assert.strictEqual(csvRes.rowCount, 2);

  state.httpRequests.length = 0;
  const migPostRes = sandbox.sendPastDataToKintoneAppB(state.activeSpreadsheet);
  assert.strictEqual(migPostRes.postedCount, 2);
  assert.strictEqual(migPostRes.ledgerAppendedCount, 2);

  const kintonePostReqs = state.httpRequests.filter((r) => r.url.includes('/k/v1/records.json'));
  assert.strictEqual(kintonePostReqs.length, 1);
  const postedRecords = kintonePostReqs[0].parsedPayload.records;
  assert.strictEqual(
    postedRecords[0]['連携先施設名_0'],
    undefined,
    'ルックアップ自動コピー先の連携先施設名_0がPOSTペイロードに含まれないこと'
  );
  assert.strictEqual(
    postedRecords[0]['連携先施設名'],
    undefined,
    '移行POSTペイロードに連携先施設名が含まれないこと'
  );
  assert.strictEqual(postedRecords[0]['更新用キー'].value, 'TEST-01');
  assert.strictEqual(postedRecords[0]['当院担当診療科'].value, '整形外科');
  assert.strictEqual(postedRecords[0]['担当医師'], undefined, '退職医は担当医師フィールドが省略されること');
  assert.strictEqual(postedRecords[1]['更新用キー'], undefined, '部分一致のみの施設は更新用キーが省略されること');
  assert.strictEqual(postedRecords[1]['返書日'], undefined, '空の返書日はPOSTペイロードから省略されること');

  recordPass(
    'タスク4 過去データ移行機能',
    '返書ステータス4段階自動判定・全月タブ合体・BOM付UTF-8 CSV出力・kintone一括POST＆年度別台帳同時反映が正常に動作すること'
  );
}

// ============================================================================
// 6. 連携フェーズ6分類・登録対象外フィルタ・左詰め・過去データ報告/照会/教示除外・再実行防止 (A〜F)
// ============================================================================
{
  const santeConfig = vm.runInContext('SANTE_CONFIG', sandbox);

  // [A] PHASES が6分類・8行目プルダウンに6つ設定されること
  assert.deepStrictEqual(
    Object.keys(santeConfig.PHASES),
    ['FORWARD', 'BACKWARD', 'REPLY', 'REPORT_LETTER', 'INQUIRY', 'OTHER'],
    'SANTE_CONFIG.PHASES が6キーを持つこと'
  );
  assert.strictEqual(santeConfig.PHASES.REPORT_LETTER, '報告書・お手紙');
  assert.strictEqual(santeConfig.PHASES.INQUIRY, '教示依頼（照会）');

  state.activeSpreadsheet = new MockSpreadsheet();
  sandbox.setupAllSheets();
  const inputSheet = state.activeSpreadsheet.getSheetByName('紹介・逆紹介入力');
  sandbox.applyInputColumnValidations(inputSheet, 3, 2);
  const phaseValidation = inputSheet.getRange(8, 3).getDataValidation();
  assert(phaseValidation && Array.isArray(phaseValidation.listValues), '8行目に入力規則が設定されること');
  assert.deepStrictEqual(
    Array.from(phaseValidation.listValues),
    [
      '前方（紹介・受入）',
      '後方（逆紹介・送出）',
      '返書（経過・結果報告）',
      '報告書・お手紙',
      '教示依頼（照会）',
      'その他（介護連携・対診）',
    ],
    '8行目プルダウンに6分類すべてが含まれること'
  );
  recordPass(
    'タスクA 連携フェーズ6分類・入力規則',
    'SANTE_CONFIG.PHASES が6分類定義され、紹介・逆紹介入力シート8行目のプルダウンに6つの選択肢が設定されること'
  );

  // [B] normalizePhaseValue エイリアス・返書との非衝突・空/未知が ''（未分類）になり前方に化けないこと
  ['報告書', 'お手紙', '受診報告', '退院報告', '死亡報告', '報告書・お手紙'].forEach((alias) => {
    assert.strictEqual(
      sandbox.normalizePhaseValue(alias),
      '報告書・お手紙',
      `「${alias}」が「報告書・お手紙」に正規化されること`
    );
  });
  ['教示', '照会', '教示依頼', '教示依頼（照会）'].forEach((alias) => {
    assert.strictEqual(
      sandbox.normalizePhaseValue(alias),
      '教示依頼（照会）',
      `「${alias}」が「教示依頼（照会）」に正規化されること`
    );
  });
  ['返書（経過・結果報告）', '返書', '経過報告', '結果報告'].forEach((alias) => {
    assert.strictEqual(
      sandbox.normalizePhaseValue(alias),
      '返書（経過・結果報告）',
      `「${alias}」が報告書に奪われず「返書（経過・結果報告）」になること`
    );
  });
  assert.strictEqual(sandbox.normalizePhaseValue(''), '', '空文字は空文字（未分類）を返すこと');
  assert.strictEqual(sandbox.normalizePhaseValue('   '), '', '空白のみは空文字（未分類）を返すこと');
  assert.strictEqual(sandbox.normalizePhaseValue('分類不能な文字列'), '', '未知値は空文字（未分類）を返し前方にしないこと');

  // 未分類列の展開時ステータス・第一報チェック・一括実行での保留確認
  state.activeSpreadsheet = new MockSpreadsheet();
  sandbox.setupAllSheets();
  const masterSheetB = state.activeSpreadsheet.getSheetByName('施設マスター');
  masterSheetB.getRange(2, 1, 1, 7).setValues([
    ['TEST-01', '○○中央病院', '東京都千代田区1-1', '03-0000-0001', '', '', ''],
  ]);
  const inputSheetB = state.activeSpreadsheet.getSheetByName('紹介・逆紹介入力');
  sandbox.expandJsonAndMatchLedger(
    JSON.stringify([
      {
        受診日: '2026-10-31',
        患者ID: '12345',
        患者氏名: '山田太郎',
        連携フェーズ: '', // 空欄（未分類）
        受入形態: '入院', // 入院でも未分類なら第一報チェックを TRUE にしない
        連携先施設名: '○○中央病院',
        連携先医師名: '佐藤一郎',
        当院担当診療科: '内科',
        担当医師: '大江健介',
      },
      {
        受診日: '2026-10-31',
        患者ID: '23456',
        患者氏名: '鈴木花子',
        連携フェーズ: '謎の文書区分', // 未知値（未分類）
        受入形態: '外来',
        連携先施設名: '○○中央病院',
        連携先医師名: '佐藤一郎',
        当院担当診療科: '内科',
        担当医師: '大江健介',
      },
    ]),
    state.activeSpreadsheet
  );

  assert.strictEqual(
    inputSheetB.getRange(4, 3).getValue(),
    '⚠ 連携フェーズ未確定のため保留',
    'JSON展開時に未分類列の4行目が保留文言になること'
  );
  assert.strictEqual(
    inputSheetB.getRange(8, 3).getValue(),
    '',
    '未分類列の8行目が勝手に前方にならず空欄であること'
  );
  assert.strictEqual(
    inputSheetB.getRange(15, 3).getValue(),
    false,
    '未分類＋入院のとき15行目（第一報作成）が TRUE にならないこと'
  );

  state.httpRequests.length = 0;
  const unresolvedBatchRes = sandbox.executeBatchProcess();
  assert.strictEqual(unresolvedBatchRes.kintonePosted, 0, '未分類列はkintone POSTされないこと');
  assert.strictEqual(unresolvedBatchRes.kintonePutUpdated, 0, '未分類列はkintone PUTされないこと');
  assert.strictEqual(unresolvedBatchRes.autoClearedCount, 0, '未分類列は自動クリアされないこと');
  assert.strictEqual(
    inputSheetB.getRange(4, 3).getValue(),
    '⚠ 連携フェーズ未確定のため保留',
    '一括実行後も未分類列の4行目が保留文言のままであること'
  );
  assert.strictEqual(
    inputSheetB.getRange(7, 3).getValue(),
    '山田太郎',
    '未分類列がクリアされずにシートに残ること'
  );
  recordPass(
    'タスクB normalizePhaseValueエイリアス・未分類保留',
    '6分類エイリアスが正しく判定され、空・未知のフェーズは前方に化けず「⚠ 連携フェーズ未確定のため保留」としてPOST・台帳追記・クリアから除外されること'
  );

  // [C & D] フェーズ別振り分け（前方・後方のみPOST、報告書・照会・その他は18行目IDがあってもPOST/PUTせずクリア）＋未処理列の左詰め
  state.activeSpreadsheet = new MockSpreadsheet();
  sandbox.setupAllSheets();
  const masterSheetC = state.activeSpreadsheet.getSheetByName('施設マスター');
  masterSheetC.getRange(2, 1, 2, 7).setValues([
    ['TEST-01', '○○中央病院', '東京都千代田区1-1', '03-0000-0001', '', '', ''],
    ['TEST-02', '△△中央病院', '東京都中央区2-2', '03-0000-0002', '', '', ''],
  ]);
  const ledgerSheetC = state.activeSpreadsheet.getSheetByName(sandbox.getLedgerSheetName());
  // 後方消込ターゲット用の未返書前方レコードを台帳2行目に用意
  ledgerSheetC.getRange(2, 1, 1, 18).setValues([
    [
      '501',
      '2026-10-01 09:00:00',
      '2026-10-01',
      '11112',
      '佐藤次郎',
      '前方（紹介・受入）',
      '外来',
      'TEST-01',
      '○○中央病院',
      '内科',
      '佐藤一郎',
      '脳神経外科',
      '溝上泰一朗',
      '未作成',
      '',
      '',
      '不要',
      '',
    ],
  ]);

  const inputSheetC = state.activeSpreadsheet.getSheetByName('紹介・逆紹介入力');
  // 7列展開:
  // Col 3 (C): 前方（完全一致） -> POST + 台帳追加 + クリア
  // Col 4 (D): 報告書・お手紙（部分一致施設「中央病院」＋18行目にダミーIDあり） -> POST/PUTせず完了クリア
  // Col 5 (E): 前方（部分一致施設「中央病院」＋17行目FAXチェックTRUE） -> POSTされるが部分一致のため残存 -> 左詰めでC列(col 3)へ移動
  // Col 6 (F): 教示依頼（照会）（18行目にダミーIDあり） -> POST/PUTせず完了クリア
  // Col 7 (G): 未分類（連携フェーズ空欄） -> POST/PUTせず保留残存 -> 左詰めでD列(col 4)へ移動
  // Col 8 (H): その他（介護連携・対診）（18行目にダミーIDあり） -> POST/PUTせず完了クリア
  // Col 9 (I): 後方（完全一致＋5重ロック一致IDあり） -> POST + 台帳追加 + PUT消込 + クリア
  sandbox.expandJsonAndMatchLedger(
    JSON.stringify([
      {
        受診日: '2026-10-31',
        患者ID: '11111',
        患者氏名: '山田太郎',
        連携フェーズ: '前方（紹介・受入）',
        受入形態: '外来',
        連携先施設名: '○○中央病院',
        連携先医師名: '佐藤一郎',
        当院担当診療科: '脳神経外科',
        担当医師: '溝上泰一朗',
      },
      {
        受診日: '2026-10-31',
        患者ID: '11113',
        患者氏名: '高橋三郎',
        連携フェーズ: '退院報告', // -> 報告書・お手紙
        受入形態: '外来',
        連携先施設名: '中央病院', // 部分一致でも登録対象外なのでクリアされる
        連携先医師名: '佐藤一郎',
      },
      {
        受診日: '2026-10-31',
        患者ID: '11114',
        患者氏名: '田中四郎',
        連携フェーズ: '前方（紹介・受入）',
        受入形態: '入院', // -> 15行目チェックボックス TRUE
        連携先施設名: '中央病院', // 部分一致のためシートに残る
        連携先医師名: '佐藤一郎',
        当院担当診療科: '脳神経内科',
        担当医師: '田中淳',
      },
      {
        受診日: '2026-10-31',
        患者ID: '11115',
        患者氏名: '伊藤五郎',
        連携フェーズ: '照会', // -> 教示依頼（照会）
        受入形態: '外来',
        連携先施設名: '○○中央病院',
        連携先医師名: '佐藤一郎',
      },
      {
        受診日: '2026-10-31',
        患者ID: '11116',
        患者氏名: '渡辺六郎',
        連携フェーズ: '', // -> 未分類保留
        受入形態: '外来',
        連携先施設名: '○○中央病院',
        連携先医師名: '佐藤一郎',
      },
      {
        受診日: '2026-10-31',
        患者ID: '11117',
        患者氏名: '中村七郎',
        連携フェーズ: '介護連携', // -> その他（介護連携・対診）
        受入形態: '外来',
        連携先施設名: '○○中央病院',
        連携先医師名: '佐藤一郎',
      },
      {
        受診日: '2026-10-31',
        患者ID: '11112',
        患者氏名: '佐藤次郎',
        連携フェーズ: '後方（逆紹介・送出）',
        受入形態: '外来',
        連携先施設名: '○○中央病院',
        連携先医師名: '佐藤一郎',
        当院担当診療科: '脳神経外科',
        担当医師: '溝上泰一朗',
      },
    ]),
    state.activeSpreadsheet
  );

  // 登録対象外3列（Col 4, Col 6, Col 8）の18行目にあえてダミーの返書消込ターゲットIDを入れる
  inputSheetC.getRange(18, 4).setValue('kintone#501 / 台帳#2 (2026-10-01受診分)');
  inputSheetC.getRange(18, 6).setValue('kintone#501 / 台帳#2 (2026-10-01受診分)');
  inputSheetC.getRange(18, 8).setValue('kintone#501 / 台帳#2 (2026-10-01受診分)');
  // Col 5（田中四郎）の17行目FAX送信完了チェックボックスも true にして左詰め後の保持を確認
  inputSheetC.getRange(17, 5).setValue(true);

  state.httpRequests.length = 0;
  inputSheetC._noteEvents.length = 0;

  const batchResCD = sandbox.executeBatchProcess();
  // POSTされるのは Col 3(山田太郎・前方), Col 5(田中四郎・前方), Col 9(佐藤次郎・後方) の3件のみ
  assert.strictEqual(batchResCD.kintonePosted, 3, '前方2件＋後方1件の計3件のみがkintone POSTされること');
  // PUTされるのは Col 9(佐藤次郎・後方) の1件のみ（登録対象外3列の18行目ダミーIDはPUTされない）
  assert.strictEqual(batchResCD.kintonePutUpdated, 1, '後方の1件のみが返書消込PUTされ、登録対象外列の18行目IDはPUTされないこと');
  // クリアされるのは Col 3(前方完全一致), Col 4(報告書), Col 6(照会), Col 8(その他), Col 9(後方完全一致) の5列
  assert.strictEqual(batchResCD.autoClearedCount, 5, '完了2列＋登録対象外3列の計5列が自動クリアされること');

  // 登録対象外列の18行目に「登録対象外のため返書消込は行いません」のNoteが設定された履歴があること
  const excludedTargetNoteEvents = inputSheetC._noteEvents.filter(
    (e) => e.row === 18 && e.note.includes('登録対象外のため返書消込は行いません')
  );
  assert.strictEqual(
    excludedTargetNoteEvents.length,
    3,
    '登録対象外3列の18行目に「登録対象外のため返書消込は行いません」が記録されること'
  );

  // kintone POST/PUT ペイロードに「報告書・お手紙」「教示依頼（照会）」「その他（介護連携・対診）」が一切含まれないこと
  const kintoneReqsCD = state.httpRequests.filter((r) => r.url.includes('/k/v1/records.json'));
  kintoneReqsCD.forEach((req) => {
    const rawStr = JSON.stringify(req.parsedPayload);
    ['報告書・お手紙', '教示依頼（照会）', 'その他（介護連携・対診）', '高橋三郎', '伊藤五郎', '中村七郎'].forEach(
      (forbidden) => {
        assert(
          !rawStr.includes(forbidden),
          `kintone通信ペイロードに「${forbidden}」が1文字も含まれないこと`
        );
      }
    );
  });

  // 台帳にも登録対象外3件や未分類1件が追加されていないこと（初期1行＋新規POST 3行＝計4行、つまりデータ3行＋ヘッダ1行＝lastRow 5）
  assert.strictEqual(ledgerSheetC.getLastRow(), 5, '台帳には初期1件＋前方2件＋後方1件の計4件のみが存在すること');

  // [D] 左詰め結果の検証:
  // 残った2列（旧Col 5: 田中四郎[部分一致保留]、旧Col 7: 渡辺六郎[フェーズ未確定]）が C列(col 3) と D列(col 4) に左詰めされていること
  assert.strictEqual(inputSheetC.getRange(7, 3).getValue(), '田中四郎', '1つ目の残存列（部分一致保留）がC列(col 3)へ左詰めされること');
  assert.strictEqual(inputSheetC.getRange(6, 3).getValue(), '11114', 'C列の患者IDが保持されること');
  assert.strictEqual(inputSheetC.getRange(5, 3).getNumberFormat(), '@', '左詰め後のC列5行目（受診日）の書式 @ が保たれること');
  assert.strictEqual(inputSheetC.getRange(6, 3).getNumberFormat(), '@', '左詰め後のC列6行目（患者ID）の書式 @ が保たれること');
  assert(
    inputSheetC.getRange(10, 3).getNote().includes('施設名が部分一致のみ'),
    '左詰め後のC列10行目に部分一致警告Noteが保たれること'
  );
  assert.strictEqual(inputSheetC.getRange(15, 3).hasCheckbox(), true, '左詰め後のC列15行目にチェックボックスが保たれること');
  assert.strictEqual(inputSheetC.getRange(15, 3).getValue(), true, '左詰め後のC列15行目のチェック状態(true)が保たれること');
  assert.strictEqual(inputSheetC.getRange(17, 3).hasCheckbox(), true, '左詰め後のC列17行目にチェックボックスが保たれること');
  assert.strictEqual(inputSheetC.getRange(17, 3).getValue(), true, '左詰め後のC列17行目のチェック状態(true)が保たれること');
  const col3PhaseVal = inputSheetC.getRange(8, 3).getDataValidation();
  assert.strictEqual(col3PhaseVal.listValues.length, 6, '左詰め後のC列8行目に6分類プルダウン入力規則が保たれること');

  assert.strictEqual(inputSheetC.getRange(7, 4).getValue(), '渡辺六郎', '2つ目の残存列（未分類保留）がD列(col 4)へ左詰めされること');
  assert.strictEqual(
    inputSheetC.getRange(4, 4).getValue(),
    '⚠ 連携フェーズ未確定のため保留',
    '左詰め後のD列4行目に未確定保留ステータスが保たれること'
  );
  assert.strictEqual(inputSheetC.getRange(15, 4).getValue(), false, '左詰め後のD列15行目チェックボックスが false で保たれること');

  // E列(col 5)〜I列(col 9)は空にクリアされ、チェックボックスも false にリセットされていること
  for (let col = 5; col <= 9; col++) {
    assert.strictEqual(inputSheetC.getRange(4, col).getValue(), '', `左詰め後に空いた列(col ${col})の4行目が空であること`);
    assert.strictEqual(inputSheetC.getRange(7, col).getValue(), '', `左詰め後に空いた列(col ${col})の7行目が空であること`);
    assert.strictEqual(inputSheetC.getRange(15, col).getValue(), false, `左詰め後に空いた列(col ${col})の15行目が false であること`);
    assert.strictEqual(inputSheetC.getRange(17, col).getValue(), false, `左詰め後に空いた列(col ${col})の17行目が false であること`);
  }

  recordPass(
    'タスクC フェーズ別振り分け・対象外除外',
    '前方・後方のみがkintone POST＆台帳追記され、報告書・お手紙／教示依頼（照会）／その他は18行目にIDがあってもPOST/PUTされず完了クリアされること'
  );
  recordPass(
    'タスクD 完了列クリア後の左詰め整理',
    '完了列クリア後、残った未処理列（部分一致保留・フェーズ未確定）がC列開始へ左詰めされ、値・Note・6分類プルダウン・数値書式・チェックボックスが保持されること'
  );

  // [E] 過去データ移行での「報告／照会／教示」除外と「逆紹介」維持
  state.activeSpreadsheet = new MockSpreadsheet();
  sandbox.setupAllSheets();
  const masterSheetE = state.activeSpreadsheet.getSheetByName('施設マスター');
  masterSheetE.getRange(2, 1, 1, 7).setValues([
    ['TEST-01', '○○中央病院', '東京都千代田区1-1', '03-0000-0001', '', '', ''],
  ]);
  const pastSheetE = state.activeSpreadsheet.insertSheet('2026年5月');
  const pastRowsE = [
    [
      '来院日',
      '患者ID',
      '患者氏名',
      '紹介元病院名',
      '紹介元診療科',
      '紹介元医師名',
      '受入・送出形態',
      '当院診療科',
      '担当医師',
      '返書日',
      '済/未',
      '予備',
      '問い合わせ', // M列 (row[12])
      '備考',       // N列 (row[13])
    ],
    ['2026-05-01', '20001', '山田太郎', '○○中央病院', '内科', '佐藤一郎', '外来', '内科', '大江健介', '', '未', '', '受診報告のみ', ''], // スキップ（報告）
    ['2026-05-02', '20002', '鈴木花子', '○○中央病院', '内科', '佐藤一郎', '外来', '内科', '大江健介', '', '未', '', '経過照会あり', ''], // スキップ（照会）
    ['2026-05-03', '20003', '高橋三郎', '○○中央病院', '内科', '佐藤一郎', '外来', '内科', '大江健介', '', '未', '', '治療方針の教示依頼', ''], // スキップ（教示）
    ['2026-05-04', '20004', '田中四郎', '○○中央病院', '内科', '佐藤一郎', '外来', '内科', '大江健介', '', '未', '', '報告・照会・教示すべて含む', ''], // 先に当たった「報告」1理由のみカウント
    ['2026-05-05', '20005', '伊藤五郎', '○○中央病院', '内科', '佐藤一郎', '外来', '内科', '大江健介', '', '未', '', '逆紹介', ''], // 除外されず後方／対象外
    ['2026-05-06', '20006', '渡辺六郎', '○○中央病院', '内科', '佐藤一郎', '外来', '内科', '大江健介', '2026-05-10', '済', '', '', ''], // 通常前方／作成・送付完了
  ];
  pastSheetE.getRange(1, 1, pastRowsE.length, 14).setValues(pastRowsE);

  const migResE = sandbox.buildPastDataMigrationSheet(state.activeSpreadsheet);
  assert.strictEqual(migResE.totalCount, 2, '報告・照会・教示の4行が除外され、逆紹介と通常前方の2行のみ残ること');
  assert.deepStrictEqual(
    { ...migResE.skippedByReason },
    { 報告: 2, 照会: 1, 教示: 1 },
    '理由別のスキップ件数（報告:2, 照会:1, 教示:1）が正確に集計され二重カウントしないこと'
  );
  assert.strictEqual(migResE.skippedCount, 4);
  assert.strictEqual(migResE.countsByStatus['対象外'], 1, '逆紹介の行は除外されず「対象外」として1件カウントされること');
  assert.strictEqual(migResE.countsByStatus['作成・送付完了'], 1);

  const migSheetE = state.activeSpreadsheet.getSheetByName('【kintone一括登録用】紹介・逆紹介');
  const migSheetRowsE = migSheetE.getRange(2, 1, 2, 17).getValues();
  assert.strictEqual(migSheetRowsE[0][1], '20005', '1行目が逆紹介の患者（20005）であること');
  assert.strictEqual(migSheetRowsE[0][4], '後方（逆紹介・送出）', '逆紹介行のフェーズが後方（逆紹介・送出）であること');
  assert.strictEqual(migSheetRowsE[0][14], '対象外', '逆紹介行の返書ステータスが対象外であること');

  const csvResE = sandbox.downloadMigrationCsv(state.activeSpreadsheet);
  assert(!csvResE.csvContent.includes('20001') && !csvResE.csvContent.includes('20002') && !csvResE.csvContent.includes('20003'), '除外行がCSVにも一切含まれないこと');
  assert(csvResE.csvContent.includes('移行済レコードID'), 'CSVヘッダに移行済レコードID列が含まれること');
  recordPass(
    'タスクE 過去データ移行の報告/照会/教示除外',
    'M列（問い合わせ）に報告・照会・教示を含む行が除外されて理由別件数が集計され、逆紹介のみの行は除外されず後方／対象外になること'
  );

  // [F] sendPastDataToKintoneAppB の「移行済レコードID」による再実行防止と途中失敗時のチャンク保存
  state.httpRequests.length = 0;
  // 1回目のPOST（2件送信）
  const firstPostF = sandbox.sendPastDataToKintoneAppB(state.activeSpreadsheet);
  assert.strictEqual(firstPostF.postedCount, 2, '1回目は2件ともPOSTされること');
  assert.strictEqual(firstPostF.skippedCount, 0, '1回目のスキップ件数は0件であること');
  assert.strictEqual(migSheetE.getRange(2, 17).getValue(), '9001', '1行目の17列目（移行済レコードID）に発番ID 9001 が書き込まれること');
  assert.strictEqual(migSheetE.getRange(3, 17).getValue(), '9002', '2行目の17列目（移行済レコードID）に発番ID 9002 が書き込まれること');

  // 2回目の再実行（全件に移行済レコードIDがあるため送信0件・スキップ2件）
  state.httpRequests.length = 0;
  const secondPostF = sandbox.sendPastDataToKintoneAppB(state.activeSpreadsheet);
  assert.strictEqual(secondPostF.postedCount, 0, '2回目の再実行ではPOST件数が0件になること');
  assert.strictEqual(secondPostF.ledgerAppendedCount, 0, '2回目の再実行では台帳追記件数も0件になること');
  assert.strictEqual(secondPostF.skippedCount, 2, '2回目の再実行ではスキップ件数が2件になること');
  assert.strictEqual(
    state.httpRequests.filter((r) => r.url.includes('/k/v1/records.json')).length,
    0,
    '全件移行済みの場合はkintone APIリクエストが1回も発生しないこと'
  );

  // 100件超（120件）で2チャンク目が失敗した場合の途中保存＆再実行テスト
  state.activeSpreadsheet = new MockSpreadsheet();
  sandbox.setupAllSheets();
  const migSheet120 = state.activeSpreadsheet.insertSheet('【kintone一括登録用】紹介・逆紹介');
  const headers16 = [
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
  ]; // あえて旧16列ヘッダで作成し、17列目ヘッダ自動追加も検証
  migSheet120.getRange(1, 1, 1, 16).setValues([headers16]);
  const rows120 = Array.from({ length: 120 }, (_, i) => [
    '2026-05-10',
    String(30001 + i),
    '山田太郎',
    '前方（紹介・受入）',
    '前方（紹介・受入）',
    '外来',
    'TEST-01',
    '○○中央病院',
    '内科',
    '内科',
    '佐藤一郎',
    '内科（一般）',
    '大江健介',
    '2026-05-15',
    '作成・送付完了',
    '',
  ]);
  migSheet120.getRange(2, 1, 120, 16).setValues(rows120);

  // 1回目の実行: 1チャンク目(100件)は成功し、2チャンク目(20件)でHTTP 500エラーを起こす
  state.kintoneCallCounter = 0;
  state.failOnKintoneCallNumber = 2;
  assert.throws(() => {
    sandbox.sendPastDataToKintoneAppB(state.activeSpreadsheet);
  }, /HTTP 500/);

  assert.strictEqual(
    migSheet120.getRange(1, 17).getValue(),
    '移行済レコードID',
    '旧16列シートの17列目ヘッダに「移行済レコードID」が自動追加されること'
  );
  assert.strictEqual(
    migSheet120.getRange(2, 17).getValue(),
    '9001',
    '途中失敗時でも1チャンク目先頭行（2行目）に移行済レコードIDが保存されていること'
  );
  assert.strictEqual(
    migSheet120.getRange(101, 17).getValue(),
    '9100',
    '途中失敗時でも1チャンク目最終行（101行目）に移行済レコードIDが保存されていること'
  );
  assert.strictEqual(
    migSheet120.getRange(102, 17).getValue(),
    '',
    '失敗した2チャンク目先頭行（102行目）は未記入のままであること'
  );

  // エラー解消後の再実行: 成功済みの100件はスキップされ、残りの20件だけが送信されること
  state.failOnKintoneCallNumber = 0;
  state.httpRequests.length = 0;
  const retryRes120 = sandbox.sendPastDataToKintoneAppB(state.activeSpreadsheet);
  assert.strictEqual(retryRes120.skippedCount, 100, '再実行時に成功済み100件がスキップされること');
  assert.strictEqual(retryRes120.postedCount, 20, '再実行時に未記入の残り20件のみがPOSTされること');
  assert.strictEqual(retryRes120.ledgerAppendedCount, 20, '再実行時に残り20件のみが台帳追記されること');
  assert.strictEqual(
    migSheet120.getRange(121, 17).getValue(),
    '9020',
    '再実行後に最終行（121行目）にも移行済レコードIDが記録されること'
  );

  recordPass(
    'タスクF 移行済レコードIDによる再実行防止・途中失敗保護',
    'sendPastDataToKintoneAppB で移行済レコードID列への記録・再実行スキップ・100件チャンク途中失敗時の成功分ID保持が正常に動作すること'
  );
}

// ============================================================================
// 5. コンプライアンス静的検証 ＆ バンドル同期確認
// ============================================================================
{
  const filesToScan = [
    ...MODULE_ORDER.map((f) => path.join(SRC_DIR, f)),
    path.join(ROOT_DIR, 'dist', 'SANTE_BRIDGE_BUNDLE.gs'),
    path.join(ROOT_DIR, 'README.md'),
    path.join(ROOT_DIR, 'docs', 'NOTION_SPECIFICATION.md'),
    path.join(ROOT_DIR, 'appsscript.json'),
  ];

  // バンドルファイルの構文チェックと src/ との完全一致確認
  const bundleContent = fs.readFileSync(path.join(ROOT_DIR, 'dist', 'SANTE_BRIDGE_BUNDLE.gs'), 'utf8');
  new vm.Script(bundleContent, { filename: 'dist/SANTE_BRIDGE_BUNDLE.gs' });

  const expectedBundleHeader = [
    '/**',
    ' * ============================================================================',
    ' * SANTE-BRIDGE (sante-bridge-liaison-hub) 本番コピペ用 全結合バンドルファイル',
    ' * 結合順序: Config.js -> LedgerMatcher.js -> DocGenerator.js -> KintoneClient.js -> ReminderService.js -> PastDataMigration.js -> Code.js',
    ' * ============================================================================',
    ' */',
  ].join('\n');
  const expectedBundleBody = MODULE_ORDER.map((file) => {
    const srcText = fs.readFileSync(path.join(SRC_DIR, file), 'utf8').trim();
    return [
      '// ============================================================================',
      `// FILE: src/${file}`,
      '// ============================================================================',
      srcText,
    ].join('\n');
  }).join('\n\n');
  const expectedBundle = `${expectedBundleHeader}\n${expectedBundleBody}\n`;
  assert.strictEqual(
    bundleContent.trim(),
    expectedBundle.trim(),
    'dist/SANTE_BRIDGE_BUNDLE.gs の中身が src/ 各ファイルの結合結果と完全に一致すること'
  );

  filesToScan.forEach((filePath) => {
    const content = fs.readFileSync(filePath, 'utf8');
    assert(
      !content.includes('generativelanguage.googleapis.com'),
      `${path.basename(filePath)} に generativelanguage.googleapis.com が含まれないこと`
    );
    assert(
      !/AIza[0-9A-Za-z_-]{20,}/.test(content),
      `${path.basename(filePath)} に AIza 形式のAPIキーが含まれないこと`
    );
  });

  recordPass(
    'コンプライアンス静的検証',
    '全ソースコード・設定ファイルに Google AI Studio エンドポイント（generativelanguage.googleapis.com）および AIza 形式のAPIキーが含まれていないこと'
  );
}

console.log('\n============================================================');
console.log(`🎉 全自動検証 完了: 全 ${results.length} 項目 PASS`);
console.log('============================================================');
