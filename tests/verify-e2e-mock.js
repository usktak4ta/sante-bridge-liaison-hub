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
    return this;
  }
  setNumberFormat() {
    return this;
  }
  setDataValidation(rule) {
    this.sheet._lastDataValidation = rule;
    return this;
  }
  insertCheckboxes() {
    return this;
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
}

class MockSpreadsheet {
  constructor() {
    this.sheets = new Map();
    this.toasts = [];
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
    ' * 結合順序: Config.js -> LedgerMatcher.js -> DocGenerator.js -> KintoneClient.js -> ReminderService.js -> Code.js',
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
