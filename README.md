# SANTE-BRIDGE (`sante-bridge-liaison-hub`)
**サンテ溝上病院 地域連携・紹介返書管理＆第一報自動化システム (v2.0.0 - Phase 2)**

Google Workspace (Gemini / Google Sheets / Google Docs / Google Apps Script) と kintone (App ID: 195)・LINE WORKS・Slack を連携し、医療情報ガイドライン（3省2ガイドライン・HIPAA BAA）準拠の Human-in-the-Loop 体制で地域医療連携業務（前方紹介受入・第一報作成・後方逆紹介・返書消込・未返書定期リマインド）を自動化するシステムです。

---

## 🔒 コンプライアンス・契約・セキュリティの大前提

1. **Google AI Studioは「BAA対象外」であり、患者情報（PHI）の送信は厳禁**
   - Google AI Studio（開発者向けプラットフォーム）は、Google CloudのHIPAA BAAなどの「対象サービス（Covered Services）」には含まれません。そのため、Google AI Studioへの患者情報（PHI）の送信およびAPIキーの利用は厳禁です。
2. **BAAが契約上保証される対象サービス（Workspace Gem / Vertex AI）のみを使用**
   - 厳格なコンプライアンス（BAAの適用、データ処理追加条項の適用、入力データのモデル学習への利用禁止など）が契約上保証されるのは、**Google Workspace（有償版）標準のGemini画面（Gem）**、または**Google Cloudプラットフォーム上で提供される「Vertex AI」経由のGemini API**のみです。
   - 本システムでは、診療情報提供書PDFの読み取りはGoogle Workspace標準のWeb UI（Gem）で行い、出力されたJSONをスプレッドシートの `C2` セルへ貼り付ける方式を採用しています。
3. **決済・アカウント状態の前提**
   - 現在のメインのお支払い方法は**クレジットカード**に設定されており、現在のアカウント状態のままでは**請求書払い（Invoice）での利用はできない**前提です。
4. **BAA対象外ツール（Slack・GAS標準ログ）への患者情報（PHI）出力禁止と誤FAX防止設計**
   - BAA対象外のSlackおよびGAS標準ログ（`Logger.log`）には、患者氏名・紹介元施設名・医師名・レスポンス本文などの実データを一切出力せず、件数や安全なメタ情報のみを出力する設計です。
   - 施設マスター照合では、部分一致での自動選択を行わず（候補をセルメモに提示し、第一報FAX番号欄には `（要確認: 施設マスター未一致）` を表示）、別施設への誤FAX送信を防止する設計です。

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
    ├── LedgerMatcher.js        # 施設マスター名寄せ・年度別台帳管理(UI固定/条件付き書式)・5重ロック自動照合
    ├── KintoneClient.js        # kintone REST API (App ID: 195) 複数トークン連結・Auto-Recovery付POST/PUT通信
    ├── DocGenerator.js         # 第一報（受診・入院報告書）Googleドキュメント複製＆プレースホルダー置換生成
    └── ReminderService.js      # 未返書リスト抽出・担当医別グルーピング・LINE WORKS Bot / Slack 定期リマインド通知
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
| `KINTONE_API_TOKEN_MASTER` | ルックアップ参照元（医療機関・施設マスター）のAPIトークン（カンマ連結送信） | 推奨 |
| `REPORT_TEMPLATE_DOC_ID` | 第一報GoogleドキュメントのテンプレートID（※メニューから自動生成可能） | ✅ |
| `REPORT_OUTPUT_FOLDER_ID` | 生成した第一報ドキュメントの保存先GoogleドライブフォルダID | 任意 |
| `SLACK_WEBHOOK_URL` | エラー・完了・リマインド通知用 Slack Incoming Webhook URL | 任意 |
| `LW_CLIENT_ID` | LINE WORKS Developer Console の Client ID | 任意 |
| `LW_CLIENT_SECRET` | LINE WORKS Developer Console の Client Secret | 任意 |
| `LW_SERVICE_ACCOUNT` | LINE WORKS サービスアカウントID (`xxxx.serviceaccount@...`) | 任意 |
| `LW_PRIVATE_KEY` | LINE WORKS 秘密鍵（PEM形式・改行は `\n` のまま保存可） | 任意 |
| `LW_BOT_ID` | 通知送信元 LINE WORKS Bot ID | 任意 |
| `LW_REMINDER_TARGET_ID` | 未返書リマインド送信先のトークルーム（Channel）ID または ユーザーID | 任意 |

### 3. スプレッドシート3タブ・テンプレート・定期トリガーの初期構築
1. **3タブ整備**: メニュー **「⚙️ 初期設定・メンテナンス ＞ 📋 3タブ初期構築・フォーマット整備」** をクリック（スクロール固定およびN列条件付き書式が自動適用されます）。
2. **第一報テンプレート生成**: メニュー **「⚙️ 初期設定・メンテナンス ＞ 📄 第一報Googleドキュメント標準テンプレート生成」** をクリック。
3. **未返書リマインド定期トリガー登録**: メニュー **「⚙️ 初期設定・メンテナンス ＞ ⏰ 未返書リマインド定期トリガー設定（毎週月・木 9:00）」** をクリック。
