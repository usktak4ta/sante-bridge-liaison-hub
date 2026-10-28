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
  const faxNumber = recordData.faxNumber || resolvedFacility.faxNumber || '';

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
    notifySlackError('第一報テンプレート取得失敗', err, {
      templateDocId: templateDocId,
      patientName: patientName,
    });
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
    '{{連携先FAX番号}}': faxNumber,
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
