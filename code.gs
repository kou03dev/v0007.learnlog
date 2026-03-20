const SHEET_TRAN = 'トラン';           // トラン シート
const SHEET_REF  = '参照用マスター';   // 参照用マスター シート

// code.gs 全文（mode 1 / mode 2 両方対応）

//    AppVer 07
//    DeployVer 34


// https://www.perplexity.ai/search/itumooshi-hua-ninarimasu-xue-x-q3xpd5n7Ttq_2.oEWVHG7w
// code.gs（code.js の整理＋機能追加15反映版・全文）


// ◆重要◆
// 「デプロイ」を実行する前に、
// 必ず code.gs の「prepareNewDeploy (関数)」をドロップダウンリストボックスから選択して、「実行」ボタンを押す。
// →「prepareNewDeploy (関数)」を実行することで、DeployVer が + 1 されます。

// ★ 手入力で管理するアプリのバージョン（AppVer）
const APP_VER = '07';
const APP_BASE_NAME = `学習記録WebApp(AppVer${APP_VER})`;

// ★ Deploy バージョン情報とログの保存先
const VERSION_LOG_SHEET_NAME = 'バージョンログ';  // なければ自動作成

// Script Properties のキー
const PROP_DEPLOY_NO       = 'APP_DEPLOY_NO';        // DeployVer の数値
const PROP_DEPLOY_DATETIME = 'APP_DEPLOY_DATETIME';  // Deploy 時刻 (yyyy/MM/dd HH:mm:ss)
const PROP_DEPLOY_LOG      = 'APP_DEPLOY_LOG';       // テキスト履歴

// Sheet のデータが入力済みの行 (A とする) を調べて、
// Sheet の max の行 (B とする) と比較する。
// B - A の差が n 以内なら、行を追加する。
// const ROW_DIFF = 1;
const ROW_DIFF = 5;

// 「データが記入済みかどうか」を判定する基準となる列（ここでは C 列）
const DATA_COLUMN_INDEX = 3;

// 「出力用の値03」に出力する共通メタ情報
const APP_META = {
  WebAppName_en_US: 'LearnLogApp',
  WebAppName_ja_JP: '学習記録アプリ',
  WebAppVer: APP_VER
};

/**
 * Web アプリのエントリポイント
 * ⇒ ここでは DeployVer を増やさず、「最後に登録された DeployVer」を表示に使う。
 *
 * 仕様: URL クエリの mode=1 / mode=2 を取得し、テンプレートに渡す。
 *       1,2 以外または未指定の場合は null として渡し、クライアント側で「2」を初期値とする。
 */
function doGet(e) {
  const deployInfo = getCurrentDeployInfo_();

  const fullTitle = `${APP_BASE_NAME} (DeployVer ${deployInfo.no})`;

  const queryMode = e && e.parameter && e.parameter.mode ? String(e.parameter.mode) : '';
  let initialMode = null;
  if (queryMode === '1' || queryMode === '2') {
    initialMode = queryMode;
  }

  const template = HtmlService.createTemplateFromFile('index');
  template.appBaseName    = APP_BASE_NAME;
  template.appVer         = APP_VER;
  template.deployNo       = deployInfo.no;
  template.deployDateTime = deployInfo.datetime;
  template.fullTitle      = fullTitle;
  template.initialMode    = initialMode;

  const output = template.evaluate()
    .setTitle(fullTitle)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');

  return output;
}

/**
 * ★デプロイの前に手動実行する関数
 */
function prepareNewDeploy() {
  const props = PropertiesService.getScriptProperties();

  const currentNoStr = props.getProperty(PROP_DEPLOY_NO);
  const currentNo = currentNoStr ? Number(currentNoStr) || 0 : 0;
  const newNo = currentNo + 1;

  const now = new Date();
  const datetime = Utilities.formatDate(
    now,
    Session.getScriptTimeZone(),
    'yyyy/MM/dd HH:mm:ss'
  );

  props.setProperty(PROP_DEPLOY_NO, String(newNo));
  props.setProperty(PROP_DEPLOY_DATETIME, datetime);

  const logStr = props.getProperty(PROP_DEPLOY_LOG) || '';
  const newLine = `${newNo}\t${datetime}`;
  const updatedLog = logStr ? logStr + '\n' + newLine : newLine;
  props.setProperty(PROP_DEPLOY_LOG, updatedLog);

  logDeployToSheet_(newNo, datetime);
}

/**
 * 現在の DeployVer 情報を取得
 */
function getCurrentDeployInfo_() {
  const props = PropertiesService.getScriptProperties();
  const noStr = props.getProperty(PROP_DEPLOY_NO);
  const dt = props.getProperty(PROP_DEPLOY_DATETIME);

  const no = noStr ? Number(noStr) || 0 : 0;
  const datetime = dt || '';

  return { no, datetime };
}

/**
 * 「バージョンログ」シートに DeployVer と日時を追記
 */
function logDeployToSheet_(no, datetime) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(VERSION_LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(VERSION_LOG_SHEET_NAME);
    sheet.getRange(1, 1).setValue('DeployVer');
    sheet.getRange(1, 2).setValue('DateTime');
  }
  const lastRow = sheet.getLastRow();
  const nextRow = lastRow + 1;

  sheet.getRange(nextRow, 1).setValue(no);
  sheet.getRange(nextRow, 2).setValue(datetime);
}

/**
 * 参照用マスターから「自由語句の AND 条件」で候補を検索（mode 1）
 */
function searchMasterByTerms(terms) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_REF);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const targetColIndexes = [0, 1]; // A,B列

  const normTerms = (terms || [])
    .map(t => (t || '').toString().trim())
    .filter(t => t.length > 0)
    .map(t => t.toLowerCase());

  if (normTerms.length === 0) {
    return [];
  }

  const results = [];

  values.forEach((row, idx) => {
    const output01 = row[0];
    const output02 = row[1];

    const targetText = targetColIndexes
      .map(i => (row[i] != null ? String(row[i]) : ''))
      .join(' ')
      .toLowerCase();

    const ok = normTerms.every(term => targetText.indexOf(term) !== -1);
    if (!ok) return;

    results.push({
      rowIndex: idx + 2,
      display: output01,
      output01: output01,
      output02: output02
    });
  });

  return results.slice(0, 100);
}

/**
 * 互換用：単一語句検索
 */
function searchMaster(query) {
  query = (query || '').toString().trim();
  if (!query) return [];
  return searchMasterByTerms([query]);
}

/**
 * mode 2 用: 参照用マスター全件を取得
 */
function getMasterAllRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_REF);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const result = values.map((row, idx) => {
    const output01 = row[0];
    const output02 = row[1];
    return {
      rowIndex: idx + 2,
      output01: output01,
      output02: output02
    };
  });
  return result;
}

/**
 * 「トラン」シートの末尾行に追記する前に、
 * 必要に応じてシート末尾に行を追加する内部関数。
 *
 * ROW_DIFF 行ぶんの「テンプレート行」をまとめて追加しておき、
 * 実際のデータ書き込みは appendTran 側で 1 行だけ行う想定。
 *
 * 機能追加11: 行追加とコピー処理に要した時間 (ms) を返却情報に含める。
 */
function ensureTranRows_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TRAN);

  const writableLastRow = sheet.getMaxRows();

  // C列ベースの末尾行取得
  const dataLastRow = sheet
    .getRange(1, DATA_COLUMN_INDEX)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();

  let message = '';
  let rowAppendElapsedMs = 0;

  if (dataLastRow > writableLastRow - ROW_DIFF) {
    const lastCol = sheet.getLastColumn();
    const templateRowIndex = writableLastRow;
    const templateRange = sheet.getRange(templateRowIndex, 1, 1, lastCol);

    const tStart = new Date().getTime();

    // ROW_DIFF 行追加
    sheet.insertRowsAfter(writableLastRow, ROW_DIFF);

    const startRow = writableLastRow + 1;
    const addedRange = sheet.getRange(startRow, 1, ROW_DIFF, lastCol);

    const templateFormulas = templateRange.getFormulas();
    if (templateFormulas && templateFormulas.length > 0) {
      const rowFormulas = templateFormulas[0];
      const formulasForAdded = [];
      for (let r = 0; r < ROW_DIFF; r++) {
        formulasForAdded.push(rowFormulas);
      }
      addedRange.setFormulas(formulasForAdded);
    }

    const templateNumberFormats = templateRange.getNumberFormats();
    if (templateNumberFormats && templateNumberFormats.length > 0) {
      const rowNumberFormats = templateNumberFormats[0];
      const formatsForAdded = [];
      for (let r = 0; r < ROW_DIFF; r++) {
        formatsForAdded.push(rowNumberFormats);
      }
      addedRange.setNumberFormats(formatsForAdded);
    }

    templateRange.copyFormatToRange(
      sheet,
      1,
      lastCol,
      startRow,
      startRow + ROW_DIFF - 1
    );

    const tEnd = new Date().getTime();
    rowAppendElapsedMs = tEnd - tStart;

    message =
      `『トラン』シートの「C列」にデータが記載されている末尾の行が (${dataLastRow})、` +
      `シートの記入可能な末尾の行の番号 (${writableLastRow}) の値と比較して「` + ROW_DIFF + `」以内であるため、` +
      `『トラン』シートへ「` + ROW_DIFF + `」行追加しました。` +
      ` (行追加所要時間 ${rowAppendElapsedMs} ms)`;
  } else {
    message =
      `『トラン』シートの「C列」にデータが記載されている末尾の行が (${dataLastRow})、` +
      `シートの記入可能な末尾の行の番号 (${writableLastRow}) の値と比較して「` + ROW_DIFF + `」以内ではないため、` +
      `『トラン』シートへ行の追加処理は不要です。` +
      ` (行追加所要時間 ${rowAppendElapsedMs} ms)`;
  }

  return { dataLastRow, writableLastRow, message, rowAppendElapsedMs };
}

/**
 * HTML テンプレート用 include
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 候補 1 件が確定したときに「トラン」シート末尾に追記
 *
 * 機能追加10: 書き込み所要時間を計測して返す。
 *
 * 機能追加15: 6列目「出力用の値03」に WebApp メタ情報を JSON 形式で出力する。
 *
 * ※ROW_DIFF が 1 より大きい場合でも、ここで書き込むのは 1 行だけ。
 *   不足分のテンプレート行は ensureTranRows_ でまとめて追加される。
 */
function appendTran(selected) {
  if (!selected || !selected.output01) {
    return { success: false, message: '不正なパラメータです。' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TRAN);

  const info = ensureTranRows_();

  // データの最終行も C列ベースで取得
  const lastRow = sheet
    .getRange(1, DATA_COLUMN_INDEX)
    .getNextDataCell(SpreadsheetApp.Direction.DOWN)
    .getRow();
  const nextRow = lastRow + 1;

  const now = new Date();
  const formatted = Utilities.formatDate(
    now,
    Session.getScriptTimeZone(),
    'yyyy/MM/dd HH:mm:ss'
  );

  const tStartWrite = new Date().getTime();

  sheet.getRange(nextRow, 1).setFormula(
    '=HYPERLINK("#gid=316040792&range=C" & MAX(FILTER(ROW(C:C),C:C<>"" )), "C列の最終データへ")'
  );
  sheet.getRange(nextRow, 2).setFormula(
    '=HYPERLINK("#gid=316040792&range=C" & MIN(FILTER(ROW(C:C),C:C<>"" )), "C列の先頭データへ")'
  );
  sheet.getRange(nextRow, 3).setValue(formatted);
  sheet.getRange(nextRow, 4).setValue(selected.output01);
  sheet.getRange(nextRow, 5).setValue(selected.output02);

  // ★追加機能15: 出力用の値03（6列目）に JSON を出力
  const deployInfo = getCurrentDeployInfo_();
  const metaForRow = {
    WebAppName_en_US: APP_META.WebAppName_en_US,
    WebAppName_ja_JP: APP_META.WebAppName_ja_JP,
    WebAppVer: APP_META.WebAppVer,
    DeployVer: String(deployInfo.no),
    DeployDateTime: deployInfo.datetime
  };
  sheet.getRange(nextRow, 6).setValue(JSON.stringify(metaForRow));

  const tEndWrite = new Date().getTime();
  const writeElapsedMs = tEndWrite - tStartWrite;

  return {
    success: true,
    insertedRow: nextRow,
    tranRowInfo: info,
    message: info.message,
    writeElapsedMs: writeElapsedMs
  };
}
