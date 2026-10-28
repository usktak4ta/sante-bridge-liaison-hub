# SANTE-BRIDGE (`sante-bridge-liaison-hub`)
**サンテ溝上病院 地域連携・紹介返書管理＆第一報自動化システム**

Google Workspace (Gemini / Google Sheets / Google Docs / Google Apps Script) と kintone (App ID: 195)・Slack を連携し、医療情報ガイドライン（3省2ガイドライン・HIPAA BAA）準拠の Human-in-the-Loop 体制で地域医療連携業務（前方紹介受入・第一報作成・後方逆紹介・返書消込）を自動化するシステムです。

---

## 📁 ディレクトリ構成

```text
sante-bridge-liaison-hub/
├── .clasp.json.example         # clasp設定テンプレート（実ファイル .clasp.json は gitignore 対象）
├── .claspignore                # clasp push 除外設定
├── .gitignore                  # 秘匿情報 (.clasp.json, .env, トークン類) の除外設定
├── README.md                   # セットアップ・デプロイ手順書
├── appsscript.json             # GASマニフェスト (V8ランタイム, Asia/Tokyo, OAuthスコープ)
├── docs/
│   └── NOTION_SPECIFICATION.md # Notion同期用：プロジェクト仕様書・マッピング表・5重ロック設計書・運用マニュアル
└── src/
    ├── Config.js               # 定数・シート定義・PropertiesService秘匿情報管理・Slack通知モジュール
    ├── Code.js                 # エントリポイント(onOpen, onEdit)・JSON横展開・一括実行・列クリア整理
    ├── LedgerMatcher.js        # 施設マスター名寄せ・年度別台帳管理・5重ロック自動照合エンジン
    ├── KintoneClient.js        # kintone REST API (App ID: 195) 一括POST/PUT通信・チャンク分割・エラー通知
    └── DocGenerator.js         # 第一報（受診・入院報告書）Googleドキュメント複製＆プレースホルダー置換生成
```

---

## 🚀 セットアップ ＆ デプロイ手順 (clasp / GitHub)

### 1. clasp 設定と GAS への Push
```bash
# 1. テンプレートから .clasp.json を作成し、対象スプレッドシートに紐づく scriptId を記入
cp .clasp.json.example .clasp.json

# 2. Google アカウントへログイン（未ログインの場合）
clasp login

# 3. GAS プロジェクトへコードを反映
clasp push
```

### 2. GAS `PropertiesService`（スクリプトプロパティ）の設定
スプレッドシートを開き、メニュー **「🏥 SANTE-BRIDGE 連携メニュー」＞「⚙️ 初期設定・メンテナンス」＞「🔑 スクリプトプロパティ初期化」** を実行後、GASエディタの「プロジェクトの設定 ＞ スクリプト プロパティ」に以下の値を設定してください。

| プロパティ名 | 設定内容 | 必須 |
| :--- | :--- | :---: |
| `KINTONE_DOMAIN` | kintoneサブドメイン（例: `sante-mizogami.cybozu.com`） | ✅ |
| `KINTONE_APP_ID` | アプリBのApp ID（デフォルト: `195`） | ✅ |
| `KINTONE_API_TOKEN` | アプリB（ID: 195）のAPIトークン（レコード閲覧・追加・編集権限） | ✅ |
| `REPORT_TEMPLATE_DOC_ID` | 第一報GoogleドキュメントのテンプレートID（※メニューから自動生成可能） | ✅ |
| `REPORT_OUTPUT_FOLDER_ID` | 生成した第一報ドキュメントの保存先GoogleドライブフォルダID | 任意 |
| `SLACK_WEBHOOK_URL` | エラー・完了通知用 Slack Incoming Webhook URL | 任意 |

### 3. スプレッドシート3タブと第一報テンプレートの初期構築
1. メニュー **「⚙️ 初期設定・メンテナンス」＞「📋 3タブ初期構築・フォーマット整備」** をクリックすると、`紹介・逆紹介入力`・`2026年度_連携台帳`・`施設マスター` の3タブが自動生成・整形されます。
2. メニュー **「⚙️ 初期設定・メンテナンス」＞「📄 第一報Googleドキュメント標準テンプレート生成」** をクリックすると、12個のプレースホルダー（`{{報告日}}` 等）を備えた標準テンプレートDocが自動作成され、`REPORT_TEMPLATE_DOC_ID` に自動登録されます。
