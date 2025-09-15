/**
 * データ入力フォーム機能
 * Google AppSheetスタイルのWebアプリケーション
 */

// ======= Webアプリケーション関連 =======

// 構造シート/フォーム生成用 定数
const STRUCTURE_SHEET_NAME = 'STRUCTURE';
const MAX_STRUCTURE_COLS = 9;

// === Staged pipeline constants (S1/S2/S3, F1/F2) ===
// 10行ヘッダー前提: 1行目=型, 2-10行目=L1..L9
const HEADER_ROWS = 10;
// DATAシート名（新パイプライン用）
const DATA_SHEET_NAME = 'DATA';
// type→入力UIの簡易対応
const TYPE_TO_WIDGET = {
  'TEXT': 'text',
  'TEXTAREA': 'textarea',
  'NUMBER': 'number',
  'DATE': 'date',
  'EXISTENCE': 'checkbox',
  'LIST': 'select'
};

// Webアプリのメインページを表示
function doGet(e) {
  // パラメータが存在しない場合のエラーハンドリング
  if (!e || !e.parameter) {
    console.log('doGet called without parameters - returning main page');
    const t = HtmlService.createTemplateFromFile('views/webapp');
    t.versionString = getJstVersionString_();
    t.showFormBuilder = isFeatureEnabled_('ENABLE_FORM_BUILDER', false);
    t.showDebugLink = isFeatureEnabled_('ENABLE_DEBUG_PAGE', false);
    return t.evaluate()
      .setTitle('📊 データ管理アプリ')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  // デバッグモードのチェック
  const page = e.parameter.page || 'main';
  const enableDebug = isFeatureEnabled_('ENABLE_DEBUG_PAGE', false);
  const enableFormBuilder = isFeatureEnabled_('ENABLE_FORM_BUILDER', false);
  const enableCsvTest = isFeatureEnabled_('ENABLE_TEST_CSV', false);
  if (page === 'form_builder') {
    if (!enableFormBuilder) {
      return HtmlService.createHtmlOutput('<h3>フォームビルダーは無効化されています</h3>');
    }
    return HtmlService.createTemplateFromFile('views/form_builder')
      .evaluate()
      .setTitle('🧩 CSV→JSON→HTML フォーム生成')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  if (page === 'debug') {
    if (!enableDebug) {
      return HtmlService.createHtmlOutput('<h3>デバッグページは無効化されています</h3>');
    }
    const t = HtmlService.createTemplateFromFile('views/debug');
    t.versionString = getJstVersionString_();
    return t.evaluate()
      .setTitle('🔧 デバッグページ - データ管理アプリ')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'reception') {
    var file = 'views/reception_form';
    const t = HtmlService.createTemplateFromFile(file);
    t.versionString = getJstVersionString_();
    return t.evaluate()
      .setTitle('📝 受付入力フォーム')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'structure_form') {
    // STRUCTUREシートからフォームHTMLを生成し、テンプレートへ挿入
    var refresh = e.parameter.refresh === '1' || e.parameter.refresh === 'true';
    var formHtml = generateFormFromStructureSheet({ refresh: refresh });
    const t = HtmlService.createTemplateFromFile('views/structure_form');
    t.formHtml = formHtml;
    t.versionString = getJstVersionString_();
    return t.evaluate()
      .setTitle('🧱 STRUCTURE フォーム')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  // ※ 旧: page=csv_converter は reception_form を返すだけの不整合な暫定ルートだったため撤去。
  //    CSV→JSON 変換の検証は ?page=test_csv で実施可能。
  if (page === 'test_csv') {
    if (!enableCsvTest) {
      return HtmlService.createHtmlOutput('<h3>CSVテストページは無効化されています</h3>');
    }
    return HtmlService.createTemplateFromFile('views/test_csv_converter')
      .evaluate()
      .setTitle('🧪 CSV to JSON 変換テスト')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // === New staged testing routes (S1/S2/S3, F1/F2) ===
  if (page === 'test_structure_to_csv') {
    const csv = structureToCsv_();
    return HtmlService.createHtmlOutput('<pre>' + escHtml_(csv) + '</pre>')
      .setTitle('🧪 S1: STRUCTURE → CSV');
  }
  if (page === 'test_csv_to_json') {
    const json = csvToJson_(structureToCsv_());
    return HtmlService.createHtmlOutput('<pre>' + escHtml_(JSON.stringify(json, null, 2)) + '</pre>')
      .setTitle('🧪 S2: CSV → JSON');
  }
  if (page === 'test_json_to_form') {
    const json = csvToJson_(structureToCsv_());
    const html = jsonToReceptionHtml_(json, {});
    return HtmlService.createHtmlOutput(html)
      .setTitle('🧪 S3: JSON → HTML');
  }
  if (page === 'reception_from_structure') {
    const json = csvToJson_(structureToCsv_());
    const html = jsonToReceptionHtml_(json, {});
    return HtmlService.createHtmlOutput(html)
      .setTitle('📝 受付（STRUCTUREベース）');
  }
  if (page === 'reception_edit') {
    const rowIndex = Number(e.parameter.row || '2');
    const initial = getReceptionJsonByRow_(rowIndex);
    const json = csvToJson_(structureToCsv_());
    const html = jsonToReceptionHtml_(json, initial);
    return HtmlService.createHtmlOutput(html)
      .setTitle('📝 受付編集');
  }
  
  const t = HtmlService.createTemplateFromFile('views/webapp');
  t.versionString = getJstVersionString_();
  t.showFormBuilder = enableFormBuilder;
  t.showDebugLink = enableDebug;
  return t.evaluate()
    .setTitle('📊 データ管理アプリ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}











// ======= スプレッドシート内編集機能（既存機能） =======

// メニュー作成
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📝 受付入力')
    .addItem('🌐 Webアプリを開く', 'openWebApp')
    .addSeparator()
    .addSubMenu(ui.createMenu('🔧 テスト・デバッグ')
      .addItem('🔍 簡単テスト画面', 'showSimpleTest'))
    .addToUi();
}

// Webアプリを開く
function openWebApp() {
  const scriptId = ScriptApp.getScriptId();
  const url = `https://script.google.com/macros/s/${scriptId}/exec`;
  
  const htmlOutput = HtmlService.createHtmlOutput(`
    <script>
      window.open('${url}', '_blank');
      google.script.host.close();
    </script>
  `);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Webアプリを開いています...');
}


// シンプルなテストページを表示
function showSimpleTest() {
  const htmlOutput = HtmlService.createTemplateFromFile('views/simple_test')
    .evaluate()
    .setWidth(600)
    .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🔧 簡単テスト');
}

// ======= ユーティリティ関数 =======

// JSTベースのバージョン文字列生成
function getJstVersionString_() {
  try {
    var s = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
    return 'ver.' + s;
  } catch (_e) {
    // 失敗時はUTCを使用
    var d = new Date();
    var yyyy = d.getUTCFullYear();
    var mm = ('0' + (d.getUTCMonth()+1)).slice(-2);
    var dd = ('0' + d.getUTCDate()).slice(-2);
    var HH = ('0' + d.getUTCHours()).slice(-2);
    var MM = ('0' + d.getUTCMinutes()).slice(-2);
    var SS = ('0' + d.getUTCSeconds()).slice(-2);
    return 'ver.' + yyyy + mm + dd + '_' + HH + MM + SS;
  }
}

// ======= 設定/フラグユーティリティ =======
/**
 * スクリプトプロパティのフラグを判定する
 * 許容値: '1','true','yes','on' (大文字小文字無視)
 * 未設定時は defaultVal を返す
 */
function isFeatureEnabled_(key, defaultVal) {
  try {
    var v = PropertiesService.getScriptProperties().getProperty(key);
    if (v == null || v === '') return !!defaultVal;
    v = String(v).toLowerCase();
    return v === '1' || v === 'true' || v === 'yes' || v === 'on';
  } catch (_e) {
    return !!defaultVal;
  }
}

// ======= 共通ユーティリティ（Sパイプライン/F I/O用） =======
const escHtml_ = s => String(s ?? '').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const joinPath_ = (a) => (a || []).filter(Boolean).join('.');
function flattenJson_(obj, prefix = '', out = {}) {
  for (const [k, v] of Object.entries(obj || {})) {
    const key = prefix ? prefix + '.' + k : k;
    if (v && typeof v === 'object' && !Array.isArray(v)) flattenJson_(v, key, out);
    else out[key] = v;
  }
  return out;
}
function setByPath_(obj, segs, val) {
  let cur = obj;
  for (let i = 0; i < segs.length; i++) {
    const k = segs[i];
    if (i === segs.length - 1) cur[k] = val;
    else { cur[k] = cur[k] || {}; cur = cur[k]; }
  }
}

// ======= S1: STRUCTURE → CSV =======
function structureToCsv_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(STRUCTURE_SHEET_NAME);
  if (!sh) throw new Error('STRUCTUREシートが見つかりません');
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(1, 1, HEADER_ROWS, lastCol).getValues(); // [10 x N]

  const rows = [];
  for (let col = 0; col < lastCol; col++) {
    const type = header[0][col] || '';
    const levels = [];
    for (let r = 1; r < HEADER_ROWS; r++) levels.push(header[r][col] || '');
    rows.push([type, ...levels]); // 10列
  }

  const lines = rows.map(r =>
    r.map(v => {
      const s = String(v ?? '');
      const needsQuote = (s.indexOf('"') !== -1) || (s.indexOf(',') !== -1) || (s.indexOf('\n') !== -1) || (s.indexOf('\r') !== -1);
      return needsQuote ? '"' + s.replace(/"/g, '""') + '"' : s;
    }).join(',')
  );
  return lines.join('\n');
}

// ======= S2: CSV → JSON（フォーム仕様） =======
// 出力例: [{ type:'TEXT', path:['受付','基本','氏名'], key:'受付.基本.氏名', widget:'text', label:'氏名' }, ...]
function csvToJson_(csvText) {
  const rows = Utilities.parseCsv(csvText);
  const out = [];
  for (const row of rows) {
    const type = (row[0] || '').toString().trim().toUpperCase();
    const path = row.slice(1, HEADER_ROWS).map(x => (x || '').toString().trim()).filter(Boolean);
    if (!path.length) continue;
    const key = joinPath_(path);
    const widget = TYPE_TO_WIDGET[type] || 'text';
    out.push({ type, path, key, widget, label: path[path.length - 1] });
  }
  return out;
}

// ======= S3: JSON → HTML（受付フォーム生成） =======
function jsonToReceptionHtml_(fields, initialData) {
  const getValueByKey = (obj, dottedKey) => {
    const segs = dottedKey.split('.');
    let cur = obj || {};
    for (const s of segs) { if (cur && Object.prototype.hasOwnProperty.call(cur, s)) cur = cur[s]; else return ''; }
    return (cur == null ? '' : cur);
  };
  const controls = fields.map(f => {
    const id = 'f_' + f.key.replace(/\W+/g, '_');
    const name = f.key;
    const label = escHtml_(f.label || f.key);
    const val = escHtml_(getValueByKey(initialData, f.key));
    switch (f.widget) {
      case 'textarea':
        return '<label class=\"block\"><div>' + label + '</div><textarea name=\"' + name + '\" id=\"' + id + '\">' + val + '</textarea></label>';
      case 'checkbox': {
        const checked = (String(val) === 'true' || String(val) === '1') ? ' checked' : '';
        return '<label class=\"block\"><input type=\"checkbox\" name=\"' + name + '\" id=\"' + id + '\" value=\"true\"' + checked + '/> ' + label + '</label>';
      }
      case 'number':
        return '<label class=\"block\"><div>' + label + '</div><input type=\"number\" name=\"' + name + '\" id=\"' + id + '\" value=\"' + val + '\"/></label>';
      case 'date':
        return '<label class=\"block\"><div>' + label + '</div><input type=\"date\" name=\"' + name + '\" id=\"' + id + '\" value=\"' + val + '\"/></label>';
      case 'select':
        return '<label class=\"block\"><div>' + label + '</div><select name=\"' + name + '\" id=\"' + id + '\"></select></label>';
      default:
        return '<label class=\"block\"><div>' + label + '</div><input type=\"text\" name=\"' + name + '\" id=\"' + id + '\" value=\"' + val + '\"/></label>';
    }
  }).join('\n');

  const html = `
<html><body>
  <h2>受付フォーム（Structure→CSV→JSON→HTML）</h2>
  <form id="reception-form" onsubmit="return submitForm(event)">
    ${controls}
    <div style="margin-top:12px;">
      <button type="submit">保存</button>
    </div>
  </form>
  <script>
    function formToJson_(form){
      const obj = {};
      for (const el of form.elements) {
        if (!el.name) continue;
        const dotted = el.name;
        const segs = dotted.split('.');
        let cur = obj;
        for (let i=0;i<segs.length;i++){
          const k = segs[i];
          if (i === segs.length - 1) {
            let v = (el.type === 'checkbox') ? el.checked : el.value;
            cur[k] = v;
          } else {
            cur[k] = cur[k] || {};
            cur = cur[k];
          }
        }
      }
      return obj;
    }
    function submitForm(e){
      e.preventDefault();
      const data = formToJson_(document.getElementById('reception-form'));
      google.script.run
        .withSuccessHandler(function(){ alert('保存しました'); })
        .withFailureHandler(function(err){ alert('保存失敗: ' + err); })
        .saveReceptionDataFromJson(data);
      return false;
    }
  </script>
</body></html>`;
  return html;
}

// ======= F1: JSON → DATAシート（保存） =======
function saveReceptionDataFromJson(json) {
  const sh = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET_NAME);
  if (!sh) throw new Error('DATAシートが見つかりません');
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(1, 1, HEADER_ROWS, lastCol).getValues(); // 10xN

  const keyToCol = new Map();
  for (let c = 0; c < lastCol; c++) {
    const path = [];
    for (let r = 1; r < HEADER_ROWS; r++) {
      const v = header[r][c];
      if (v) path.push(String(v));
    }
    if (path.length) keyToCol.set(joinPath_(path), c + 1); // 1-based
  }

  const flat = flattenJson_(json);
  const row = new Array(lastCol).fill('');
  for (const [k, v] of Object.entries(flat)) {
    const col = keyToCol.get(k);
    if (col) row[col - 1] = (v == null) ? '' : v;
  }
  // 現段階は新規追加（更新は後段TODO）
  sh.appendRow(row);
}

// ======= F2: DATAシート → JSON（指定行を復元） =======
function getReceptionJsonByRow_(rowIndex) {
  const sh = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET_NAME);
  if (!sh) throw new Error('DATAシートが見つかりません');
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(1, 1, HEADER_ROWS, lastCol).getValues();
  const values = sh.getRange(rowIndex, 1, 1, lastCol).getValues()[0];

  const out = {};
  for (let c = 0; c < lastCol; c++) {
    const segs = [];
    for (let r = 1; r < HEADER_ROWS; r++) {
      const v = header[r][c];
      if (v) segs.push(String(v));
    }
    if (!segs.length) continue;
    setByPath_(out, segs, values[c]);
  }
  return out;
}


// ========= 受付（多段ヘッダー） =========

// 受付用シートを取得/作成
function getReceptionSheet_() {
  const ssId = PropertiesService.getScriptProperties().getProperty('DATA_SPREADSHEET_ID');
  const ss = ssId ? SpreadsheetApp.openById(ssId) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('受付データ');
  if (!sheet) {
    sheet = ss.insertSheet('受付データ');
  }
  return sheet;
}

// オブジェクトをパス配列にフラット化
function flattenToPaths_(obj, prefix) {
  prefix = prefix || [];
  const out = [];
  const isObj = function(v){ return v && typeof v === 'object' && !Array.isArray(v); };
  
  // 深さ制限: 9層を超える場合は処理停止
  if (prefix.length >= 9) {
    return out;
  }
  
  if (Array.isArray(obj)) {
    // LIST内にDICTがある場合はエラー
    for (var i=0;i<obj.length;i++) {
      if (isObj(obj[i])) {
        throw new Error('LIST内にDICTが含まれています。仕様違反です: ' + JSON.stringify(obj[i]));
      }
    }
    // 配列全体をJSON文字列として単一パスで出力
    out.push([prefix, JSON.stringify(obj)]);
  } else if (isObj(obj)) {
    // DICT遭遇時: 親ノードの存在フラグ（値1）を追加
    out.push([prefix, 1]);
    // 9層目の場合は子要素を処理しない（深さ制限）
    if (prefix.length < 9) {
      // 子要素を処理
      for (var k in obj) {
        if (Object.prototype.hasOwnProperty.call(obj,k)) {
          out.push.apply(out, flattenToPaths_(obj[k], prefix.concat(k)));
        }
      }
    }
  } else {
    out.push([prefix, obj]);
  }
  return out;
}

// ヘッダーを作成（10行固定）: 1行目=型情報, 2-10行目=L1,L2,L3,L4,L5,L6,L7,L8,L9
function createNestedHeaders_(sheet, headerPaths, headerKinds) {
  if (!headerPaths || headerPaths.length === 0) return 0;
  // 新仕様: 10行固定 (1行目=型情報, 2-10行目=L1,L2,L3,L4,L5,L6,L7,L8,L9)
  var depth = 10;
  var cols = headerPaths.length;
  var values = [];
  
  // 1行目: 型情報（headerKinds）
  var typeRow = [];
  for (var c=0; c<cols; c++) {
    typeRow.push(headerKinds && headerKinds[c] ? headerKinds[c] : 'SCALAR');
  }
  values.push(typeRow);
  
  // 2-10行目: L1,L2,L3,L4,L5,L6,L7,L8,L9のパス情報
  for (var r=1; r<depth; r++) {
    var row = [];
    for (var c=0; c<cols; c++) {
      // パスが9段未満の場合は"NULL"で埋める
      row.push(headerPaths[c][r-1] || 'NULL');
    }
    values.push(row);
  }
  
  sheet.clear();
  sheet.getRange(1,1,depth,cols).setValues(values);
  sheet.setFrozenRows(depth);
  
  // スタイル
  var headerRange = sheet.getRange(1,1,depth,cols);
  headerRange.setBackground('#1f2937');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  for (var c2=1; c2<=cols; c2++) sheet.setColumnWidth(c2, 180);
  
  return depth;
}

// 既存ヘッダー（10行固定）を取得: 1行目=型情報, 2-10行目=パス
function readHeaderPaths_(sheet, headerRows) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0 || headerRows === 0) return { paths: [], kinds: [] };
  // 10行固定に対応：headerRowsが10以外でも10段として読み取る
  var actualRows = Math.max(headerRows, 10);
  var vals = sheet.getRange(1,1,actualRows,lastCol).getValues();
  var paths = [];
  var kinds = [];
  
  for (var c=0; c<lastCol; c++) {
    // 1行目: 型情報
    var kind = (vals[0] && vals[0][c]) ? vals[0][c].toString() : 'SCALAR';
    kinds.push(kind);
    
    // 2-10行目: パス情報 (L1,L2,L3,L4,L5,L6,L7,L8,L9)
    var path = [];
    for (var r=1; r<10; r++) { // 2-10行目を読み取り
      var cell = (vals[r] && vals[r][c]) ? vals[r][c].toString() : 'NULL';
      if (cell !== 'NULL' && cell !== '') {
        path.push(cell);
      }
    }
    paths.push(path);
  }
  return { paths: paths, kinds: kinds };
}

// パスをカノニカル文字列へ（キー）
function pathKey_(path){ return path.join('›'); }

// パス配列の辞書順ソート（安定）
function sortPaths_(paths){ return paths.slice().sort(function(a,b){ var A=pathKey_(a), B=pathKey_(b); return A<B?-1:A>B?1:0; }); }

// パスから型情報を生成
function generateHeaderKinds_(headerPaths, flatData) {
  var kinds = [];
  var pathValueMap = {};
  
  // flatDataから各パスの値の種類を収集
  if (flatData && flatData.length > 0) {
    flatData.forEach(function(rowFlat) {
      rowFlat.forEach(function(item) {
        var pathKey = pathKey_(item[0]);
        var value = item[1];
        if (!pathValueMap[pathKey]) pathValueMap[pathKey] = [];
        pathValueMap[pathKey].push(value);
      });
    });
  }
  
  // 各パスに対して型を決定
  headerPaths.forEach(function(path) {
    var pathKey = pathKey_(path);
    var values = pathValueMap[pathKey] || [];
    var kind = 'SCALAR'; // デフォルト
    
    // 値のパターンから型を推定
    for (var i = 0; i < values.length; i++) {
      var val = values[i];
      if (typeof val === 'string' && val.match(/^\[.*\]$/)) {
        // 配列リテラル文字列
        kind = 'LIST';
        break;
      } else if (val === 1) {
        // 値が1の場合は常にEXISTENSE（DICT存在フラグ）
        kind = 'EXISTENSE';
        break;
      }
    }
    kinds.push(kind);
  });
  
  return kinds;
}

// 受付データを保存（多段ヘッダー対応）
function saveReceptionData(payload) {
  try {
    var sheet = getReceptionSheet_();
    var flat = flattenToPaths_(payload, []); // [[path[], value]]
    // ヘッダー候補生成（値は列作成のため path のみ抽出）
    var headerPaths = sortPaths_(flat.map(function(x){return x[0];}));

    var headerRows = sheet.getFrozenRows();
    if (sheet.getLastRow() <= headerRows) {
      // シート新規 or 空 → ヘッダー作成
      var headerKinds = generateHeaderKinds_(headerPaths, [flat]);
      headerRows = createNestedHeaders_(sheet, headerPaths, headerKinds);
    }

    // 現在のヘッダーを取得
    var headerInfo = readHeaderPaths_(sheet, headerRows);
    var current = headerInfo.paths;
    var currentKinds = headerInfo.kinds;
    var keyToCol = {};
    for (var i=0;i<current.length;i++) keyToCol[pathKey_(current[i])] = i+1; // 1-based col

    // 新規パスがあれば列追加
    var newPaths = headerPaths.filter(function(p){ return !(pathKey_(p) in keyToCol); });
    if (newPaths.length>0){
      var startCol = sheet.getLastColumn()+1;
      var allPaths = current.concat(newPaths);
      // すべてを再並べ替えして書き直す（列増分を右端に追加）
      var sorted = sortPaths_(allPaths);
      var allFlat = [flat]; // 現在のデータを含める
      var allKinds = generateHeaderKinds_(sorted, allFlat);
      createNestedHeaders_(sheet, sorted, allKinds); // ヘッダー全体を書き直し
      // 再マップ
      keyToCol = {};
      for (var j=0;j<sorted.length;j++) keyToCol[pathKey_(sorted[j])] = j+1;
    }

    // ID列キー
    var idKey = pathKey_(['システム','ID']);
    var createdKey = pathKey_(['システム','作成日時']);
    var updatedKey = pathKey_(['システム','更新日時']);

    // 既存データのID->行マップを作成
    var idCol = keyToCol[idKey];
    var writeRow;
    var nowIso = new Date().toISOString();

    if (idCol) {
      var lastRowAll = sheet.getLastRow();
      var idVals = lastRowAll > headerRows ? sheet.getRange(headerRows+1, idCol, lastRowAll - headerRows, 1).getValues() : [];
      // 既存行の探索（payload.system.id がある場合）
      var payloadId = null;
      try { payloadId = (payload && payload.システム && payload.システム.ID != null) ? String(payload.システム.ID) : null; } catch(_e) { payloadId = null; }
      var existingRow = null;
      if (payloadId != null && payloadId !== '') {
        for (var r=0;r<idVals.length;r++) {
          if (String(idVals[r][0]) === String(payloadId)) { existingRow = headerRows + 1 + r; break; }
        }
      }

      if (existingRow) {
        // 既存更新
        writeRow = existingRow;
      } else {
        // 新規 → MAX+1 採番
        var maxId = 0;
        for (var r2=0;r2<idVals.length;r2++) {
          var v = parseInt(idVals[r2][0], 10);
          if (!isNaN(v) && v > maxId) maxId = v;
        }
        var newId = maxId + 1;
        // payloadにID/時刻を反映
        if (!payload.システム) payload.システム = {};
        payload.システム.ID = newId;
        payload.システム.作成日時 = nowIso;
        payload.システム.更新日時 = nowIso;
        flat = flattenToPaths_(payload, []);
        // 追加入力分の列が増えていないかを再確認
        headerPaths = sortPaths_(flat.map(function(x){return x[0];}));
        var newPaths2 = headerPaths.filter(function(p){ return !(pathKey_(p) in keyToCol); });
        if (newPaths2.length>0){
          var allPaths2 = current.concat(newPaths2);
          var sorted2 = sortPaths_(allPaths2);
          var allFlat2 = [flat]; // 更新されたデータを含める
          var allKinds2 = generateHeaderKinds_(sorted2, allFlat2);
          createNestedHeaders_(sheet, sorted2, allKinds2);
          keyToCol = {};
          for (var j2=0;j2<sorted2.length;j2++) keyToCol[pathKey_(sorted2[j2])] = j2+1;
        }
        writeRow = Math.max(sheet.getLastRow(), headerRows) + 1;
      }
    } else {
      // ID列がない → ヘッダー未整備のため新規として扱う
      writeRow = Math.max(sheet.getLastRow(), headerRows) + 1;
    }

    var lastCol = sheet.getLastColumn();
    var rowVals = new Array(lastCol).fill('');
    flat.forEach(function(item){
      var key = pathKey_(item[0]);
      var col = keyToCol[key];
      if (col) rowVals[col-1] = item[1];
    });
    // 更新時刻の更新（既存行）
    if (keyToCol[updatedKey]) {
      rowVals[keyToCol[updatedKey]-1] = nowIso;
    }
    sheet.getRange(writeRow, 1, 1, lastCol).setValues([rowVals]);
    var savedId = keyToCol[idKey] ? rowVals[keyToCol[idKey]-1] : (payload && payload.システム && payload.システム.ID);
    return { ok: true, row: writeRow, id: savedId, message: '✅ 受付データを保存しました（行 ' + writeRow + '）' };
  } catch (err) {
    console.error('saveReceptionData error:', err);
    throw new Error('受付データ保存に失敗: ' + err);
  }
}

// パス配列辞書からオブジェクトへ復元
function unflattenFromPaths_(entries) {
  var root = {};
  function setPath(obj, path, value) {
    var cur = obj;
    for (var i=0;i<path.length-1;i++) {
      var k = path[i];
      // 数字なら配列インデックスとして扱う（1-based -> 0-based）
      var idx = String(k).match(/^\d+$/) ? (parseInt(k,10)-1) : null;
      if (idx != null) {
        if (!Array.isArray(cur)) return; // 型崩れ回避
        if (!cur[idx]) cur[idx] = {};
        cur = cur[idx];
      } else {
        if (typeof cur[k] !== 'object' || cur[k] == null) cur[k] = {};
        cur = cur[k];
      }
    }
    var last = path[path.length-1];
    cur[last] = value;
  }
  entries.forEach(function(ent){ setPath(root, ent[0], ent[1]); });
  return root;
}

// IDで1件取得（多段ヘッダー）
function getReceptionById(id) {
  var sheet = getReceptionSheet_();
  var headerRows = sheet.getFrozenRows();
  var headerInfo = readHeaderPaths_(sheet, headerRows);
  var current = headerInfo.paths;
  var keyToCol = {};
  for (var i=0;i<current.length;i++) keyToCol[pathKey_(current[i])] = i+1;
  var idCol = keyToCol[pathKey_(['システム','ID'])];
  if (!idCol) throw new Error('ID列が見つかりません');
  var lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) throw new Error('データがありません');
  var vals = sheet.getRange(headerRows+1, idCol, lastRow - headerRows, 1).getValues();
  var targetRow = null;
  for (var r=0;r<vals.length;r++) { if (String(vals[r][0]) === String(id)) { targetRow = headerRows + 1 + r; break; } }
  if (!targetRow) throw new Error('指定IDが見つかりません: ' + id);
  var lastCol = sheet.getLastColumn();
  var row = sheet.getRange(targetRow, 1, 1, lastCol).getValues()[0];
  var entries = [];
  for (var c=1;c<=lastCol;c++) {
    var path = current[c-1];
    var val = row[c-1];
    if (path && path.length>0) entries.push([path, val]);
  }
  return unflattenFromPaths_(entries);
}

// 一覧向け（簡易フラット配列）
function listReceptionIndex() {
  var sheet = getReceptionSheet_();
  var headerRows = sheet.getFrozenRows();
  var headerInfo = readHeaderPaths_(sheet, headerRows);
  var current = headerInfo.paths;
  var keyToCol = {};
  for (var i=0;i<current.length;i++) keyToCol[pathKey_(current[i])] = i+1;
  var idCol = keyToCol[pathKey_(['システム','ID'])] || null;
  var dateCol = keyToCol[pathKey_(['メタ','受付日'])] || null;
  var assigneeCol = keyToCol[pathKey_(['メタ','担当者'])] || null;
  var wardCol = keyToCol[pathKey_(['メタ','病棟'])] || null;
  var contactCol = keyToCol[pathKey_(['メタ','連絡方法'])] || null;
  var consultTargetCol = keyToCol[pathKey_(['相談','相談対象'])] || null;
  var consultTypesCol = keyToCol[pathKey_(['相談','相談種別'])] || null;
  var responseTypesCol = keyToCol[pathKey_(['対応','対応種別'])] || null;
  var detailsCol = keyToCol[pathKey_(['相談','詳細'])] || null;
  var createdCol = keyToCol[pathKey_(['システム','作成日時'])] || null;
  var updatedCol = keyToCol[pathKey_(['システム','更新日時'])] || null;

  var lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) return [];
  var lastCol = sheet.getLastColumn();
  var data = sheet.getRange(headerRows+1, 1, lastRow - headerRows, lastCol).getValues();
  function getVal(row, col) { return col ? row[col-1] : ''; }

  // 相談種別（階層JSON保存）の列群を特定（後方互換なし）
  var consultPrefixCols = [];
  for (var ci=0; ci<current.length; ci++) {
    var p = current[ci];
    if (p && p.length >= 3 && p[0]==='相談' && p[1]==='相談種別') {
      consultPrefixCols.push({ col: ci+1, tail: p.slice(2) });
    }
  }

  function summarizeConsultationRow(rowVals){
    if (consultPrefixCols.length === 0) return '';
    var picks = consultPrefixCols.filter(function(meta){
      var v = getVal(rowVals, meta.col);
      return v !== '' && v !== null && v !== undefined && v !== 0 && v !== false;
    }).map(function(meta){ return meta.tail; });
    if (picks.length === 0) return '';

    var groups = {};
    for (var i=0;i<picks.length;i++){
      var t = picks[i];
      var top = t[0];
      groups[top] = groups[top] || [];
      groups[top].push(t.slice(1));
    }
    var parts = [];
    if (groups['一般論的相談']){
      var gs = groups['一般論的相談'].map(function(s){ return s[0]; }).filter(function(x){ return !!x && x !== '（選択）'; });
      if (gs.length) parts.push('一般論的相談: ' + gs.join('、'));
      else parts.push('一般論的相談');
    }
    if (groups['個別案件相談']){
      var env = [];
      var eco = [];
      var leaves = [];
      groups['個別案件相談'].forEach(function(arr){
        if (!arr || arr.length===0) return;
        var head = arr[0];
        if (head === '生活環境被害') {
          if (arr[1] && arr[1] !== '（選択）') env.push(arr[1]);
          else env.push('（選択）');
        } else if (head === '生態系かく乱') {
          if (arr[1] && arr[1] !== '（選択）') eco.push(arr[1]);
          else eco.push('（選択）');
        } else {
          leaves.push(head);
        }
      });
      var segs = [];
      // 表示は子があれば括弧付き、なければ名称のみ
      var envChildren = env.filter(function(x){ return x && x !== '（選択）'; });
      var ecoChildren = eco.filter(function(x){ return x && x !== '（選択）'; });
      if (env.length) segs.push(envChildren.length ? ('生活環境被害(' + envChildren.join('、') + ')') : '生活環境被害');
      if (eco.length) segs.push(ecoChildren.length ? ('生態系かく乱(' + ecoChildren.join('、') + ')') : '生態系かく乱');
      if (leaves.length) segs.push(leaves.join('、'));
      if (segs.length) parts.push('個別案件相談: ' + segs.join('、'));
    }
    if (groups['その他']){
      parts.push('その他');
    }
    return parts.join(' / ');
  }

  return data.map(function(row){
    return [
      getVal(row, idCol),
      getVal(row, dateCol),
      getVal(row, consultTargetCol),
      getVal(row, assigneeCol),
      getVal(row, wardCol),
      getVal(row, contactCol),
      summarizeConsultationRow(row),
      Array.isArray(getVal(row, responseTypesCol)) ? getVal(row, responseTypesCol).join(',') : getVal(row, responseTypesCol),
      getVal(row, detailsCol),
      getVal(row, createdCol),
      getVal(row, updatedCol)
    ];
  });
}

// 旧シートの削除ユーティリティは廃止しました（受付データのみを対象とするため）

// スプレッドシート情報をデバッグ用に取得
function debugSpreadsheetInfo() {
  try {
    console.log('=== SERVER: debugSpreadsheetInfo called ===');
    
    // スプレッドシートを取得
    const ssId = PropertiesService.getScriptProperties().getProperty('DATA_SPREADSHEET_ID');
    const ss = ssId ? SpreadsheetApp.openById(ssId) : SpreadsheetApp.getActiveSpreadsheet();
    
    // 基本情報を取得
    const spreadsheetInfo = {
      id: ss.getId(),
      name: ss.getName(),
      url: ss.getUrl(),
      sheets: ss.getSheets().map(sheet => sheet.getName())
    };
    
    console.log('=== SERVER: spreadsheetInfo collected ===', JSON.stringify(spreadsheetInfo));
    
    return {
      spreadsheetInfo: spreadsheetInfo,
      timestamp: new Date().toISOString(),
      success: true
    };
  } catch (error) {
    console.error('=== SERVER: debugSpreadsheetInfo error ===', error);
    return {
      error: error.toString(),
      timestamp: new Date().toISOString(),
      success: false
    };
  }
}

// ======= 汎用 JSON ⇔ シート 変換（5行ヘッダー：1行目=型, 2-5行目=4層構造） =======

// 指定シートを取得/作成
function getOrCreateSheet_(name) {
  const ssId = PropertiesService.getScriptProperties().getProperty('DATA_SPREADSHEET_ID');
  const ss = ssId ? SpreadsheetApp.openById(ssId) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

// JSON(配列 or オブジェクト) → シート
// data: Object | Object[]
// 仕様: 10行ヘッダー（1行目=型情報, 2-10行目=9層構造）にフラット化して列を作成し、各行へ値を書き込む
function jsonToSheet(sheetName, data) {
  try {
    var sheet = getOrCreateSheet_(sheetName);
    var rows = Array.isArray(data) ? data.slice() : [data];
    if (rows.length === 0) return { ok: true, rows: 0, message: '入力データが空です' };

    // 全行のパスを収集
    var allPathsSet = {};
    var flattenedRows = rows.map(function(obj){
      var flat = flattenToPaths_(obj, []); // [[path[], value]]
      flat.forEach(function(pv){ allPathsSet[pathKey_(pv[0])] = pv[0]; });
      return flat;
    });
    var allPaths = Object.keys(allPathsSet).map(function(k){ return allPathsSet[k]; });
    allPaths = sortPaths_(allPaths);

    // ヘッダー作成（10行固定：1行目=型情報, 2-10行目=L1,L2,L3,L4,L5,L6,L7,L8,L9）
    var headerKinds = generateHeaderKinds_(allPaths, flattenedRows);
    var headerRows = createNestedHeaders_(sheet, allPaths, headerKinds);

    // 列マップ
    var keyToCol = {};
    for (var i=0;i<allPaths.length;i++) keyToCol[pathKey_(allPaths[i])] = i+1;

    // 既存データをクリア（ヘッダー以外）
    var lastRow = sheet.getLastRow();
    if (lastRow > headerRows) sheet.getRange(headerRows+1, 1, lastRow - headerRows, sheet.getMaxColumns()).clearContent();

    // 書き込み
    var lastCol = sheet.getLastColumn();
    var outValues = [];
    flattenedRows.forEach(function(flat){
      var rowVals = new Array(lastCol).fill('');
      flat.forEach(function(item){
        var key = pathKey_(item[0]);
        var col = keyToCol[key];
        if (col) rowVals[col-1] = item[1];
      });
      outValues.push(rowVals);
    });
    if (outValues.length) sheet.getRange(headerRows+1, 1, outValues.length, lastCol).setValues(outValues);

    return { ok: true, rows: outValues.length, cols: lastCol, headerRows: headerRows };
  } catch (e) {
    console.error('jsonToSheet error:', e);
    throw new Error('jsonToSheet 失敗: ' + e);
  }
}

// シート → JSON(配列)
function sheetToJson(sheetName) {
  try {
    var sheet = getOrCreateSheet_(sheetName);
    var headerRows = Math.max(10, sheet.getFrozenRows() || 0);
    var headerInfo = readHeaderPaths_(sheet, headerRows);
    var paths = headerInfo.paths;
    var kinds = headerInfo.kinds;
    if (!paths || paths.length === 0) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow <= headerRows) return [];
    var lastCol = sheet.getLastColumn();
    var data = sheet.getRange(headerRows+1, 1, lastRow - headerRows, lastCol).getValues();
    var out = [];
    for (var r=0;r<data.length;r++){
      var row = data[r];
      var entries = [];
      for (var c=0;c<paths.length;c++){
        var p = paths[c];
        if (!p || p.length===0) continue;
        var value = row[c];
        var kind = kinds[c] || 'SCALAR';
        
        // LIST型の場合はJSON.parse()で配列復元
        if (kind === 'LIST' && typeof value === 'string' && value.match(/^\[.*\]$/)) {
          try {
            value = JSON.parse(value);
          } catch (parseErr) {
            console.warn('LIST型パース失敗:', value, parseErr);
            // パース失敗時は元の文字列のまま
          }
        }
        
        entries.push([p, value]);
      }
      out.push(unflattenFromPaths_(entries));
    }
    return out;
  } catch (e) {
    console.error('sheetToJson error:', e);
    throw new Error('sheetToJson 失敗: ' + e);
  }
}

// ======= 簡易テスト用エンドポイント =======

function simpleTest() {
  return 'ok';
}

function pingTest() {
  return { pong: true, at: new Date().toISOString() };
}

// データの安全取得（存在しない場合はnull）
function getAllDataSafe() {
  try {
    return sheetToJson('受付データ');
  } catch (e) {
    return null;
  }
}

function debugGetAllDataStep() {
  try {
    var data = sheetToJson('受付データ');
    return 'rows=' + data.length;
  } catch (e) {
    return 'error: ' + e;
  }
}

// 直接テスト: サンプルJSONをシートへ書き込み→読み戻し件数を返す
function runDirectTest() {
  // サンプルデータ（LIST内にはプリミティブのみ - DICT不可）
  var sample = [
    {
      メタ: { 連絡方法: 'メール' },
      相談: {
        相談対象: '外来生物',
        相談種別: {
          一般論的相談: { 入門: true },
          個別案件相談: { 生活環境被害: { 騒音: true } }
        },
        詳細: 'テスト1'
      },
      tags: ['外来生物', '相談', 'メール']
    },
    {
      メタ: { 連絡方法: '電話' },
      相談: { 相談対象: '在来生物', 詳細: 'テスト2' },
      tags: ['在来生物', '電話']
    }
  ];

  var sheetName = 'JSON変換テスト';
  var sheet = getOrCreateSheet_(sheetName);
  var res = jsonToSheet(sheetName, sample);
  var back = sheetToJson(sheetName);
  return {
    ok: true,
    write: res,
    readRows: back.length,
    sheet: { name: sheet.getName(), url: sheet.getParent().getUrl() }
  };
}

// ======= Debug helpers for JSON⇔Sheet =======

function getJsonSampleSimple_() {
  return {
    id: 1,
    name: 'Alice',
    active: true,
    score: 98.5,
    contact: { email: 'alice@example.com', phone: '090-0000-0000' },
    tags: ['alpha', 'beta']
  };
}

function getJsonSampleVarious_() {
  return {
    システム: { バージョン: '1.0', 生成日時: new Date().toISOString() },
    ユーザー: {
      ID: 123,
      氏名: { 姓: '山田', 名: '太郎' },
      連絡先: { メール: 'taro@example.com', 電話: null }
    },
    注文: [
      'A-001:1200円',
      'A-002:800円'
    ],
    設定: { 通知: { メール: true, SMS: false }, 言語: 'ja' },
    備考: ''
  };
}

function debugJsonSamples() {
  return { simple: getJsonSampleSimple_(), various: getJsonSampleVarious_() };
}

function debugWriteSample(sampleType, sheetName) {
  var sample = sampleType === 'various' ? getJsonSampleVarious_() : getJsonSampleSimple_();
  var name = sheetName && String(sheetName).trim() ? String(sheetName).trim() : ('DEBUG_' + (sampleType === 'various' ? 'VARIANTS' : 'SIMPLE'));
  var res = jsonToSheet(name, Array.isArray(sample) ? sample : [sample]);
  var sheet = getOrCreateSheet_(name);
  return { ok: true, write: res, sheet: { name: sheet.getName(), url: sheet.getParent().getUrl() } };
}

function debugReadSheet(sheetName) {
  if (!sheetName) throw new Error('sheetName is required');
  return sheetToJson(String(sheetName));
}

function debugWriteRawJson(sheetName, jsonText) {
  if (!sheetName) throw new Error('sheetName is required');
  var parsed;
  try { parsed = JSON.parse(jsonText); } catch(e) { throw new Error('JSON parse error: ' + e); }
  return jsonToSheet(String(sheetName), parsed);
}

// ======= CSV to JSON Converter Test Functions =======

/**
 * CSV変換機能の総合テスト
 * @return {Object} テスト結果
 */
function testCSVConverter() {
  var results = {
    success: true,
    totalTests: 0,
    passedTests: 0,
    failedTests: 0,
    testResults: [],
    summary: ''
  };
  
  function addTestResult(testName, passed, expected, actual, error) {
    results.totalTests++;
    if (passed) {
      results.passedTests++;
    } else {
      results.failedTests++;
      results.success = false;
    }
    
    results.testResults.push({
      name: testName,
      passed: passed,
      expected: expected,
      actual: actual,
      error: error || null
    });
  }
  
  try {
    // Test 1: parseCSVCell - 基本分割
    var cell1 = parseCSVCell('氏名/re:^.{1,40}$/姓と名の間は空白1つ');
    var expected1 = { title: '氏名', type: 're:^.{1,40}$', hint: '姓と名の間は空白1つ' };
    addTestResult('parseCSVCell - 基本分割', 
      cell1.title === expected1.title && cell1.type === expected1.type && cell1.hint === expected1.hint,
      expected1, cell1);
    
    // Test 2: parseCSVCell - 部分省略
    var cell2 = parseCSVCell('X/display:満年齢');
    var expected2 = { title: 'X', type: 'display:満年齢', hint: null };
    addTestResult('parseCSVCell - 部分省略', 
      cell2.title === expected2.title && cell2.type === expected2.type && cell2.hint === expected2.hint,
      expected2, cell2);
    
    // Test 3: parseCSVCell - タイトルのみ
    var cell3 = parseCSVCell('T2');
    var expected3 = { title: 'T2', type: null, hint: null };
    addTestResult('parseCSVCell - タイトルのみ', 
      cell3.title === expected3.title && cell3.type === expected3.type && cell3.hint === expected3.hint,
      expected3, cell3);
    
    // Test 4: escapeSlashes - エスケープ処理
    var escaped = escapeSlashes('test\\/path\\/to\\/file');
    var expectedEscaped = 'test/path/to/file';
    addTestResult('escapeSlashes - エスケープ処理', 
      escaped === expectedEscaped, expectedEscaped, escaped);
    
    // Test 5: validateFormStructure - 正常なツリー
    var validTree = {
      title: 'ROOT',
      children: [
        { title: 'S1', type: 'selector:RADIO' },
        { title: 'S2', type: 'selector:RADIO' }
      ]
    };
    var validation1 = validateFormStructure(validTree);
    addTestResult('validateFormStructure - 正常なツリー', 
      validation1.valid === true, true, validation1.valid);
    
    // Test 6: validateFormStructure - 混在エラー
    var invalidTree = {
      title: 'ROOT',
      children: [
        { title: 'S1', type: 'selector:RADIO' },
        { title: 'S2', type: 're:^[0-9]+$' }
      ]
    };
    var validation2 = validateFormStructure(invalidTree);
    addTestResult('validateFormStructure - 混在エラー検出', 
      validation2.valid === false, false, validation2.valid);
    
    // Test 7: parseCSVFormDefinition - 仕様書正常例（サーバ仕様に準拠）
    var csvNormal = 'L1,L2,L3,L4,L5,L6,L7,L8,L9\n' +
                   'フォーム,,,,,,,,\n' +
                   ',基本情報/display:入力必須,,,,,,,,\n' +
                   ',,性別/selector:RADIO,,,,,,,\n' +
                   ',,,男性/selector:RADIO,,,,,,\n' +
                   ',,,女性/selector:RADIO,,,,,,\n' +
                   ',年齢/re:^\\d{1,3}$,,,,,,,\n' +
                   ',備考/re:^.{0,200}$,,,,,,,';
    
    var result1 = parseCSVFormDefinition(csvNormal);
    addTestResult('parseCSVFormDefinition - 仕様書正常例', 
      result1.success === true, true, result1.success, result1.error);
    
    // Test 8: parseCSVFormDefinition - エラー例（混在）
    var csvError = 'L1,L2,L3,L4,L5,L6,L7,L8,L9\n' +
                   'Block_X,,,,,,,,\n' +
                   'Q1,,,,,,,,\n' +
                   ',S1/selector:RADIO,,,,,,,\n' +
                   ',S2/re:^[0-9]+$,,,,,,,';
    
    var result2 = parseCSVFormDefinition(csvError);
    addTestResult('parseCSVFormDefinition - エラー例（混在）', 
      result2.success === false, false, result2.success);
    
    // Test 9: parseCSVFormDefinition - 無効な正規表現
    var csvInvalidRegex = 'L1,L2,L3,L4,L5,L6,L7,L8,L9\n' +
                         'Block_X,,,,,,,,\n' +
                         'field/re:[invalid,,,,,,,';
    
    var result3 = parseCSVFormDefinition(csvInvalidRegex);
    addTestResult('parseCSVFormDefinition - 無効な正規表現', 
      result3.success === false, false, result3.success);
    
    // Test 10: parseCSVFormDefinition - 空のCSV
    var result4 = parseCSVFormDefinition('');
    addTestResult('parseCSVFormDefinition - 空のCSV', 
      result4.success === false, false, result4.success);
    
  } catch (error) {
    addTestResult('テスト実行中エラー', false, 'エラーなし', error.toString(), error.toString());
  }
  
  results.summary = results.passedTests + '/' + results.totalTests + ' テスト通過 (' + 
                   Math.round((results.passedTests / results.totalTests) * 100) + '%)';
  
  return results;
}

/**
 * 特定のCSVデータでのテスト実行
 * @param {string} csvData - テスト用CSVデータ
 * @return {Object} テスト結果
 */
function testSpecificCSV(csvData) {
  try {
    var result = parseCSVFormDefinition(csvData);
    return {
      success: true,
      conversionSuccess: result.success,
      data: result.data,
      error: result.error,
      message: result.message
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
      message: 'テスト実行でエラーが発生しました'
    };
  }
}

// 互換: デバッグページのシート作成テスト用
function testCreateDataSheet() {
  var sheet = getOrCreateSheet_('受付データ');
  // 最低限のヘッダー
  var headers = [ ['システム','ID'], ['システム','作成日時'], ['システム','更新日時'] ];
  var headerKinds = ['SCALAR', 'SCALAR', 'SCALAR']; // 基本的な型情報
  createNestedHeaders_(sheet, headers, headerKinds);
  return '受付データ シートを初期化しました';
}

// ======= CSV to JSON Form Definition Converter =======

/**
 * エスケープされたスラッシュを復元する
 * @param {string} value - 処理対象の文字列
 * @return {string} エスケープ解除された文字列
 */
function escapeSlashes(value) {
  if (!value || typeof value !== 'string') return value;
  return value.replace(/\\\//g, '/');
}

/**
 * CSVセル値をtitle/type/hintに分割
 * @param {string} cellValue - CSVセルの値
 * @return {Object} {title, type, hint}
 */
function parseCSVCell(cellValue) {
  if (!cellValue || typeof cellValue !== 'string') {
    return { title: null, type: null, hint: null };
  }
  
  // エスケープ処理後にスラッシュで分割（最大2回）
  var unescaped = escapeSlashes(cellValue);
  var parts = unescaped.split('/');
  
  return {
    title: parts[0] || null,
    type: parts[1] || null,
    hint: parts[2] || null
  };
}

/**
 * 解析済み行データからノードツリーを構築
 * @param {Array} parsedRows - 解析済み行データの配列
 * @return {Object} ルートノード
 */
function buildNodeTree(parsedRows) {
  if (!parsedRows || parsedRows.length === 0) {
    throw new Error('パース済みデータが空です');
  }
  
  var nodes = new Map(); // ノードキー → ノードオブジェクト
  var nodesByParent = new Map(); // 親ノードキー → 子ノード配列
  
  // 各行を処理
  for (var rowIdx = 0; rowIdx < parsedRows.length; rowIdx++) {
    var row = parsedRows[rowIdx];
    
    // 各列（L1-L9）を処理
    for (var colIdx = 0; colIdx < 9; colIdx++) {
      var cell = row[colIdx];
      if (!cell || !cell.title) continue;
      
      // 親ノードを探索
      var parentNode = null;
      var parentKey = null;
      
      if (colIdx > 0) {
        // 同行の左列をチェック
        for (var leftCol = colIdx - 1; leftCol >= 0; leftCol--) {
          if (row[leftCol] && row[leftCol].title) {
            parentKey = getNodeKey(rowIdx, leftCol, row[leftCol].title, row[leftCol].type);
            parentNode = nodes.get(parentKey);
            break;
          }
        }
        
        // 左列に親がない場合、上の行を遡る
        if (!parentNode && colIdx > 0) {
          for (var upRow = rowIdx - 1; upRow >= 0; upRow--) {
            var upRowData = parsedRows[upRow];
            if (upRowData[colIdx - 1] && upRowData[colIdx - 1].title) {
              parentKey = getNodeKey(upRow, colIdx - 1, upRowData[colIdx - 1].title, upRowData[colIdx - 1].type);
              parentNode = nodes.get(parentKey);
              break;
            }
          }
        }
      }
      
      // 現在のノードキーを生成
      var nodeKey = getNodeKey(rowIdx, colIdx, cell.title, cell.type);
      
      // 既存ノードがあるかチェック（マージ対象）
      var existingNode = nodes.get(nodeKey);
      if (!existingNode) {
        // 新規ノード作成
        var newNode = {
          title: cell.title,
          type: cell.type,
          hint: cell.hint,
          children: []
        };
        
        // type, hintがnullの場合は省略
        if (!newNode.type) delete newNode.type;
        if (!newNode.hint) delete newNode.hint;
        
        nodes.set(nodeKey, newNode);
        
        // 親子関係を登録
        if (parentNode) {
          if (!nodesByParent.has(parentKey)) {
            nodesByParent.set(parentKey, []);
          }
          nodesByParent.get(parentKey).push(newNode);
          parentNode.children.push(newNode);
        }
      }
    }
  }
  
  // ルートノードを探す（親がないノード）
  var rootNodes = [];
  for (var [key, node] of nodes) {
    var hasParent = false;
    for (var [parentKey, children] of nodesByParent) {
      if (children.includes(node)) {
        hasParent = true;
        break;
      }
    }
    if (!hasParent) {
      rootNodes.push(node);
    }
  }
  
  if (rootNodes.length === 0) {
    throw new Error('ルートノードが見つかりません');
  }
  
  if (rootNodes.length === 1) {
    return rootNodes[0];
  }
  
  // 複数のルートがある場合は仮想ルートを作成
  return {
    title: 'ROOT',
    children: rootNodes
  };
}

/**
 * ノードの一意キーを生成
 * @param {number} row - 行番号
 * @param {number} col - 列番号  
 * @param {string} title - タイトル
 * @param {string} type - タイプ
 * @return {string} ノードキー
 */
function getNodeKey(row, col, title, type) {
  return col + ':' + (title || '') + ':' + (type || '');
}

/**
 * フォーム構造をバリデーション
 * @param {Object} nodeTree - ノードツリー
 * @return {Object} バリデーション結果
 */
function validateFormStructure(nodeTree) {
  var errors = [];
  
  function validateNode(node, path) {
    if (!node) return;
    
    var currentPath = path ? path + ' > ' + node.title : node.title;
    
    // typeの語彙チェック
    if (node.type) {
      var validTypes = /^(selector:(RADIO|CHECKBOX|DROPDOWN)|re:.+|display:.+)$/;
      if (!validTypes.test(node.type)) {
        errors.push('無効なtype: ' + node.type + ' at ' + currentPath);
      }
      
      // 正規表現の妥当性チェック
      if (node.type.startsWith('re:')) {
        try {
          var pattern = node.type.substring(3);
          new RegExp(pattern);
        } catch (e) {
          errors.push('無効な正規表現: ' + node.type + ' at ' + currentPath + ' - ' + e.message);
        }
      }
    }
    
    // 子ノードの選択肢グループチェック
    if (node.children && node.children.length > 0) {
      var selectorChildren = node.children.filter(function(child) {
        return child.type && child.type.startsWith('selector:');
      });
      
      if (selectorChildren.length > 0) {
        // 選択肢グループの混在チェック
        var selectorTypes = selectorChildren.map(function(child) {
          return child.type;
        });
        var uniqueTypes = selectorTypes.filter(function(type, index) {
          return selectorTypes.indexOf(type) === index;
        });
        
        if (uniqueTypes.length > 1) {
          errors.push('選択肢グループ内でselector種別が混在: ' + uniqueTypes.join(', ') + ' at ' + currentPath);
        }
        
        // 非selector要素の混入チェック
        var nonSelectorChildren = node.children.filter(function(child) {
          return !child.type || !child.type.startsWith('selector:');
        });
        
        if (nonSelectorChildren.length > 0) {
          errors.push('選択肢グループに非selector要素が混入 at ' + currentPath);
        }
      }
    }
    
    // 子ノードを再帰的にバリデーション
    if (node.children) {
      node.children.forEach(function(child) {
        validateNode(child, currentPath);
      });
    }
  }
  
  validateNode(nodeTree, '');
  
  return {
    valid: errors.length === 0,
    errors: errors
  };
}

/**
 * CSV形式のフォーム定義をJSONに変換
 * @param {string} csvContent - CSV形式の文字列
 * @return {Object} 変換結果
 */
function parseCSVFormDefinition(csvContent) {
  try {
    if (!csvContent || typeof csvContent !== 'string') {
      throw new Error('CSVコンテンツが無効です');
    }
    
    // CSVをパース（簡易実装 - Google Apps ScriptのUtilities.parseCsv()を使用予定）
    var lines = csvContent.split('\n').filter(function(line) {
      return line.trim().length > 0;
    });
    
    if (lines.length === 0) {
      throw new Error('CSVデータが空です');
    }
    
    // ヘッダー行をチェック
    var headerLine = lines[0];
    var expectedHeader = 'L1,L2,L3,L4,L5,L6,L7,L8,L9';
    if (!headerLine.includes('L1') || !headerLine.includes('L9')) {
      throw new Error('CSVヘッダーが正しくありません。L1,L2,...,L9の形式である必要があります');
    }
    
    // データ行をパース
    var parsedRows = [];
    for (var i = 1; i < lines.length; i++) {
      var line = lines[i].trim();
      if (!line) continue;
      
      // CSV行を9列に分割（簡易実装）
      var cells = line.split(',');
      while (cells.length < 9) {
        cells.push('');
      }
      
      var parsedRow = [];
      for (var j = 0; j < 9; j++) {
        var cellValue = cells[j] ? cells[j].trim() : '';
        // ダブルクォートを削除（CSV標準）
        if (cellValue.startsWith('"') && cellValue.endsWith('"')) {
          cellValue = cellValue.slice(1, -1).replace(/""/g, '"');
        }
        parsedRow.push(parseCSVCell(cellValue));
      }
      parsedRows.push(parsedRow);
    }
    
    if (parsedRows.length === 0) {
      throw new Error('有効なデータ行がありません');
    }
    
    // ノードツリーを構築
    var nodeTree = buildNodeTree(parsedRows);
    
    // バリデーション
    var validation = validateFormStructure(nodeTree);
    if (!validation.valid) {
      throw new Error('バリデーションエラー: ' + validation.errors.join(', '));
    }
    
  return {
      success: true,
      data: nodeTree,
      message: 'CSV to JSON変換が完了しました'
    };
    
  } catch (error) {
    console.error('parseCSVFormDefinition error:', error);
    return {
      success: false,
      error: error.message || error.toString(),
      message: 'CSV to JSON変換に失敗しました'
    };
  }
}

// ======= CSV -> HTML (GAS側) =======

/**
 * ノードツリーから縦並び/インデント/動的展開付きHTMLを生成
 * @param {Object} model - { title, type?, hint?, children[] }
 * @return {{success:boolean, html?:string, error?:string}}
 */
function generateFormHtmlFromNodeTree_(model) {
  try {
    model = model || { title: 'ROOT', children: [] };
    function esc(s){ return String(s == null ? '' : s).replace(/[&<>]/g, function(ch){ return ({'&':'&amp;','<':'&lt;','>':'&gt;'}[ch]); }); }
    var css = [
      ':root{--bg:#f6f8fb;--card:#fff;--text:#111827;--muted:#6b7280;--primary:#2563eb;--border:#e5e7eb;}',
      '*{box-sizing:border-box;}',
      'body{margin:0;background:linear-gradient(120deg,#eef2ff,#fdf2f8);font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Hiragino Kaku Gothic ProN,Meiryo,Arial,sans-serif;color:var(--text);}',
      '.container{max-width:1200px;margin:0 auto;padding:20px;}',
      '.header{display:flex;align-items:center;justify-content:space-between;gap:12px;background:var(--card);border:1px solid var(--border);border-radius:14px;padding:14px 16px;box-shadow:0 8px 24px rgba(0,0,0,.06);}',
      '.title{font-size:18px;font-weight:700;}',
      '.blocks{display:flex;flex-direction:column;gap:14px;margin-top:14px;}',
      '.block{background:var(--card);border:1px solid var(--border);border-radius:14px;padding:14px;box-shadow:0 8px 24px rgba(0,0,0,.06);}',
      '.block h2{margin:0 0 8px;font-size:16px;color:#2563eb;}',
      '.step{margin:10px 0;}',
      '.label{font-weight:600;margin-bottom:6px;display:flex;align-items:center;gap:8px;}',
      '.hint-badge{font-size:12px;color:#2563eb;border:1px solid #c7d2fe;background:#eef2ff;border-radius:999px;padding:2px 8px;}',
      '.options{display:flex;flex-direction:column;gap:8px;align-items:flex-start;}',
      '.chip{display:block;border:1px solid var(--border);background:#fff;border-radius:10px;padding:8px 10px;}',
      'input,textarea,select{width:100%;max-width:600px;border:1px solid var(--border);border-radius:10px;padding:8px 10px;background:#fafafa;}'
    ].join('\n');

    var html = [];
    html.push('<!DOCTYPE html>');
    html.push('<html lang="ja">');
    html.push('<head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">');
    html.push('<title>フォーム生成 (CSVより)</title>');
    html.push('<style>'+css+'</style></head>');
    html.push('<body><div class="container">');
    html.push('<div class="header"><div class="title">フォーム生成 (CSVより)</div><div class="actions"><button class="btn" id="btn-reset">リセット</button><button class="btn" id="btn-export">入力値をJSONで取得</button></div></div>');
    html.push('<div class="blocks" id="blocks">');

    var idSeq = 0;
    function renderNode(node, depth){
      depth = depth || 0;
      var hasChildren = Array.isArray(node.children) && node.children.length>0;
      var indent = Math.max(0, depth) * 16;

      if (depth === 0){
        html.push('<div class="block">');
        if (node.title){ html.push('<h2>'+esc(node.title)+'</h2>'); }
        if (node.hint){ html.push('<div class="label">'+esc(node.hint)+'</div>'); }
        (node.children||[]).forEach(function(ch){ renderNode(ch, depth+1); });
        html.push('</div>');
        return;
      }

      if (hasChildren){
        var selectorChildren = node.children.filter(function(c){ return c.type && c.type.startsWith('selector:'); });
        var nonSelectorChildren = node.children.filter(function(c){ return !c.type || !c.type.startsWith('selector:'); });
        var uniq = selectorChildren.map(function(c){ return c.type; }).filter(function(t,i,a){ return a.indexOf(t)===i; });
        if (selectorChildren.length>=2 && nonSelectorChildren.length===0 && uniq.length===1){
          var kind = uniq[0];
          var groupId = 'g_'+(idSeq++);
          html.push('<div class="step" style="margin-left:'+indent+'px">');
          html.push('<div class="label">'+esc(node.title||'')+(node.hint?' <span class="hint-badge" title="'+esc(node.hint)+'">hint</span>':'')+'</div>');
          if (kind==='selector:RADIO'){
            var name='r_'+groupId;
            html.push('<div class="options" data-group="'+groupId+'">');
            node.children.forEach(function(opt){
              var optId='opt_'+(idSeq++);
              var contId='cont_'+optId;
              html.push('<label class="chip"><input type="radio" name="'+name+'" value="'+esc(opt.title||'')+'" data-role="opt-radio" data-group="'+groupId+'" data-target="'+contId+'">'+esc(opt.title||'')+'</label>');
              html.push('<div id="'+contId+'" class="opt-children" data-group="'+groupId+'" style="display:none;">');
              renderNode(opt, depth+1);
              html.push('</div>');
            });
            html.push('</div>');
          } else if (kind==='selector:CHECKBOX'){
            var cbGroup=groupId;
            html.push('<div class="options" data-group="'+cbGroup+'">');
            node.children.forEach(function(opt){
              var optId='opt_'+(idSeq++);
              var contId='cont_'+optId;
              html.push('<label class="chip"><input type="checkbox" value="'+esc(opt.title||'')+'" data-role="opt-checkbox" data-group="'+cbGroup+'" data-target="'+contId+'">'+esc(opt.title||'')+'</label>');
              html.push('<div id="'+contId+'" class="opt-children" data-group="'+cbGroup+'" style="display:none;">');
              renderNode(opt, depth+1);
              html.push('</div>');
            });
            html.push('</div>');
          } else if (kind==='selector:DROPDOWN'){
            var selId='sel_'+groupId;
            html.push('<select id="'+selId+'" data-role="opt-select" data-group="'+groupId+'">');
            node.children.forEach(function(opt){
              var optId='opt_'+(idSeq++);
              var contId='cont_'+optId;
              html.push('<option value="'+contId+'">'+esc(opt.title||'')+'</option>');
            });
            html.push('</select>');
            node.children.forEach(function(opt){
              var optId='opt_'+(idSeq++);
              var contId='cont_'+optId;
              html.push('<div id="'+contId+'" class="opt-children" data-group="'+groupId+'" style="display:none;">');
              renderNode(opt, depth+1);
              html.push('</div>');
            });
          }
          html.push('</div>');
          return;
        }
      }

      if (node.title && (!node.type || node.type.startsWith('display:'))){
        html.push('<div class="step" style="margin-left:'+indent+'px">');
        html.push('<div class="label">'+esc(node.title)+(node.hint?' <span class="hint-badge" title="'+esc(node.hint)+'">hint</span>':'')+'</div>');
        html.push('</div>');
        (node.children||[]).forEach(function(ch){ renderNode(ch, depth+1); });
        return;
      }

      if (node.type && node.type.startsWith('re:')){
        html.push('<div class="step" style="margin-left:'+indent+'px">');
        html.push('<div class="label">'+esc(node.title||'')+(node.hint?' <span class="hint-badge" title="'+esc(node.hint)+'">hint</span>':'')+'</div>');
        html.push('<input type="text">');
        html.push('</div>');
        (node.children||[]).forEach(function(ch){ renderNode(ch, depth+1); });
        return;
      }

      (node.children||[]).forEach(function(ch){ renderNode(ch, depth+1); });
    }

    if (model.title==='ROOT' && Array.isArray(model.children)){
      model.children.forEach(function(ch){ renderNode(ch, 0); });
    } else {
      renderNode(model, 0);
    }

    html.push('</div>');
    html.push('<script>(function(){\n' +
      'function getLabelText(label){ if(!label) return ""; var c=label.cloneNode(true); var badges=c.querySelectorAll(\'.hint-badge\'); for(var i=0;i<badges.length;i++){badges[i].remove();} return (c.textContent||\'\').trim(); }\n' +
      'function resetAll(){ var root=document.getElementById(\'blocks\'); if(!root) return; var inputs=root.querySelectorAll(\'input, textarea, select\'); for(var i=0;i<inputs.length;i++){ var el=inputs[i]; var t=(el.type||\'\').toLowerCase(); if(t===\'radio\'||t===\'checkbox\'){ el.checked=false; } else if(el.tagName===\'SELECT\'){ el.selectedIndex=0; } else { el.value=\'\'; } } }\n' +
      'function hideGroup(group){ var nodes=document.querySelectorAll(\'.opt-children[data-group="\'+group+\'"]\'); for(var i=0;i<nodes.length;i++){ nodes[i].style.display=\'none\'; } }\n' +
      'function bindToggles(){\n' +
      '  var radios=document.querySelectorAll(\'input[data-role=opt-radio]\');\n' +
      '  for(var i=0;i<radios.length;i++){ radios[i].addEventListener(\'change\', function(){ var g=this.getAttribute(\'data-group\'); hideGroup(g); var t=this.getAttribute(\'data-target\'); var el=document.getElementById(t); if(el) el.style.display=this.checked?\'block\':\'none\'; }); }\n' +
      '  var cbs=document.querySelectorAll(\'input[data-role=opt-checkbox]\');\n' +
      '  for(var j=0;j<cbs.length;j++){ cbs[j].addEventListener(\'change\', function(){ var t=this.getAttribute(\'data-target\'); var el=document.getElementById(t); if(el) el.style.display=this.checked?\'block\':\'none\'; }); }\n' +
      '  var sels=document.querySelectorAll(\'select[data-role=opt-select]\');\n' +
      '  for(var k=0;k<sels.length;k++){ sels[k].addEventListener(\'change\', function(){ var g=this.getAttribute(\'data-group\'); hideGroup(g); var contId=this.value; var el=document.getElementById(contId); if(el) el.style.display=\'block\'; }); }\n' +
      '}\n' +
      'function toJson(){ var root=document.getElementById(\'blocks\'); var steps=root?root.querySelectorAll(\'.step\'):[]; var out={}; function add(key,val){ if(!key) return; if(out.hasOwnProperty(key)){ if(!Array.isArray(out[key])) out[key]=[out[key]]; out[key].push(val); } else { out[key]=val; } } for(var i=0;i<steps.length;i++){ var step=steps[i]; var label=step.querySelector(\'.label\'); var key=getLabelText(label); if(!key) continue; var radios=step.querySelectorAll(\'input[type=radio]\'); if(radios.length){ var chosen=null; for(var r=0;r<radios.length;r++){ if(radios[r].checked){ chosen=radios[r].value; break; } } add(key, chosen); continue; } var checks=step.querySelectorAll(\'input[type=checkbox]\'); if(checks.length){ var vals=[]; for(var c=0;c<checks.length;c++){ if(checks[c].checked) vals.push(checks[c].value); } add(key, vals); continue; } var sel=step.querySelector(\'select\'); if(sel){ var opt=sel.options[sel.selectedIndex]; var val=opt?opt.textContent:sel.value; add(key, val); continue; } var ta=step.querySelector(\'textarea\'); if(ta){ add(key, ta.value); continue; } var txt=step.querySelector(\'input[type=text],input[type=date],input[type=time],input[type=number],input[type=email],input[type=tel],input[type=hidden]\'); if(txt){ add(key, txt.value); continue; } } return out; }\n' +
      'function downloadJson(data){ try{ var blob=new Blob([JSON.stringify(data,null,2)],{type:\'application/json\'}); var url=URL.createObjectURL(blob); var a=document.createElement(\'a\'); a.href=url; a.download=\'form-inputs.json\'; document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url); } catch(e){ alert(\'JSON出力に失敗: \'+e); } }\n' +
      'var btnReset=document.getElementById(\'btn-reset\'); if(btnReset) btnReset.addEventListener(\'click\', function(){ resetAll(); });\n' +
      'var btnExport=document.getElementById(\'btn-export\'); if(btnExport) btnExport.addEventListener(\'click\', function(){ var data=toJson(); downloadJson(data); });\n' +
      'bindToggles();\n' +
    '})();</script>');
    html.push('</div></body></html>');
    return { success:true, html: html.join('') };
  } catch(e) { return { success:false, error: e.message || String(e) }; }
}

/**
 * CSV→MODEL→HTML（GAS公開用）
 * form_builder.html から google.script.run.generateFormFromCsv(csv) で呼ばれる
 */
function generateFormFromCsv(csvText){
  var parsed = parseCSVFormDefinition(csvText || '');
  if (!parsed || !parsed.success){ return { success:false, error: parsed && parsed.error || 'parse failed' }; }
  var html = generateFormHtmlFromNodeTree_(parsed.data);
  if (!html || !html.success){ return { success:false, error: html && html.error || 'html failed' }; }
  return { success:true, model: parsed.data, html: html.html };
}

// ======= STRUCTUREシート → ツリー → HTML =======

/**
 * STRUCTUREシートを取得（なければ作成）
 */
function getStructureSheet_() {
  return getOrCreateSheet_(STRUCTURE_SHEET_NAME);
}

/**
 * STRUCTUREシートの使用範囲を2次元配列で取得（最大 MAX_STRUCTURE_COLS 列）
 * 文字列はtrimしたものを返す
 */
function readStructureGrid_() {
  var sheet = getStructureSheet_();
  var lastRow = sheet.getLastRow();
  var lastCol = Math.min(sheet.getLastColumn(), MAX_STRUCTURE_COLS);
  if (lastRow === 0 || lastCol === 0) return [];
  var vals = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var out = [];
  for (var r = 0; r < vals.length; r++) {
    var row = [];
    for (var c = 0; c < lastCol; c++) {
      var v = vals[r][c];
      row.push(v == null ? '' : String(v).trim());
    }
    out.push(row);
  }
  return out;
}

/**
 * 構造グリッドのMD5署名を計算
 */
function getStructureSignature_() {
  var grid = readStructureGrid_();
  var payload = JSON.stringify(grid);
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, payload);
  var hex = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = bytes[i];
    if (b < 0) b += 256; // convert signed byte to unsigned
    var h = b.toString(16);
    if (h.length < 2) h = '0' + h;
    hex += h;
  }
  return hex;
}

/**
 * セル文字列を title/type/hint/(4つ目以降) に分割
 * デリミタは常に '/'
 */
function parseStructureCell_(cellText) {
  if (!cellText || typeof cellText !== 'string') {
    return { title: null, type: null, hint: null, extras: [] };
  }
  var unescaped = escapeSlashes(cellText);
  var parts = unescaped.split('/');
  return {
    title: parts[0] || null,
    type: parts.length > 1 ? (parts[1] || null) : null,
    hint: parts.length > 2 ? (parts[2] || null) : null,
    extras: parts.length > 3 ? parts.slice(3) : []
  };
}

/**
 * STRUCTUREシートからノードツリーを構築
 * 左→右 親子。親が同行左列にいなければ、左列の上方直近を親とする
 */
function parseStructureSheet() {
  var grid = readStructureGrid_();
  if (!grid || grid.length === 0) {
    // 空なら仮想ルートのみ返す
    return { title: 'ROOT', children: [] };
  }
  var parsedRows = [];
  for (var r = 0; r < grid.length; r++) {
    var row = [];
    for (var c = 0; c < MAX_STRUCTURE_COLS; c++) {
      var cellVal = c < grid[r].length ? grid[r][c] : '';
      row.push(parseStructureCell_(cellVal));
    }
    parsedRows.push(row);
  }
  // 既存のビルダーを再利用
  var nodeTree = buildNodeTree(parsedRows);
  return nodeTree;
}

/**
 * タイトル中の '_' を '__' へエスケープ
 */
function escapeUnderscore_(s) {
  return String(s == null ? '' : s).replace(/_/g, '__');
}

/**
 * ツリー全体のIDを祖先タイトル連結で割り当て、重複チェック
 * 仕様: 祖先→自分のタイトルを '_' 連結（タイトル内 '_' は '__' にエスケープ）
 * 例: A1-B1-C2 → C2.id = 'A1_B1_C2'
 */
function computeAndValidateNodeIds_(root) {
  var used = Object.create(null);
  function visit(node, ancestors) {
    if (!node || !node.title) return;
    var titleEsc = escapeUnderscore_(node.title);
    var id = ancestors.length ? (ancestors.join('_') + '_' + titleEsc) : titleEsc;
    node.id = id;
    if (used[id]) {
      throw new Error('ID重複: ' + id + '（タイトル経路の重複）');
    }
    used[id] = true;
    var nextAnc = ancestors.concat(titleEsc);
    if (Array.isArray(node.children)) {
      for (var i = 0; i < node.children.length; i++) {
        visit(node.children[i], nextAnc);
      }
    }
  }
  // ルート（仮想ROOT）の場合は子から開始
  if (root && root.title === 'ROOT' && Array.isArray(root.children)) {
    for (var j = 0; j < root.children.length; j++) {
      visit(root.children[j], []);
    }
  } else {
    visit(root, []);
  }
  return root;
}

/**
 * STRUCTUREツリーからシンプルなフォームHTMLを生成
 * 対応type: text, textarea, number, date, email, checkbox（他はtext扱い）
 * 子を持つノードは group として fieldset/legend でレンダリング
 */
function generateFormHtmlFromStructureTree_(root) {
  function esc(s){ return String(s == null ? '' : s).replace(/[&<>]/g, function(ch){ return ({'&':'&amp;','<':'&lt;','>':'&gt;'}[ch]); }); }
  var html = [];
  html.push('<div class="structure-form" style="display:flex;flex-direction:column;gap:12px;">');

  function renderNode(node, depth){
    var hasChildren = Array.isArray(node.children) && node.children.length > 0;
    var type = (node.type || '').toLowerCase();
    if (hasChildren) {
      // group container
      if (node.title && node.title !== 'ROOT') {
        html.push('<fieldset data-type="group" style="border:1px solid #e5e7eb;border-radius:10px;padding:10px 12px;">');
        html.push('<legend style="padding:0 6px;color:#2563eb;">' + esc(node.title) + '</legend>');
      }
      for (var i=0;i<node.children.length;i++) renderNode(node.children[i], depth+1);
      if (node.title && node.title !== 'ROOT') html.push('</fieldset>');
      return;
    }

    // leaf input
    var id = node.id || '';
    var name = id;
    var hint = node.hint || '';
    var label = node.title || '';
    var commonAttr = ' id="' + esc(id) + '" name="' + esc(name) + '" data-title="' + esc(node.title||'') + '" data-type="' + esc(node.type||'') + '" data-hint="' + esc(node.hint||'') + '"';
    html.push('<div class="form-field" style="display:flex;flex-direction:column;gap:6px;">');
    if (type === 'checkbox') {
      html.push('<label style="display:flex;align-items:center;gap:8px;">');
      html.push('<input type="checkbox"' + commonAttr + ' />');
      html.push('<span>' + esc(label) + '</span>');
      html.push('</label>');
    } else if (type === 'textarea') {
      html.push('<label for="' + esc(id) + '">' + esc(label) + '</label>');
      html.push('<textarea' + commonAttr + (hint ? (' placeholder="' + esc(hint) + '"') : '') + '></textarea>');
    } else if (type === 'number' || type === 'date' || type === 'email' || type === 'text') {
      html.push('<label for="' + esc(id) + '">' + esc(label) + '</label>');
      html.push('<input type="' + esc(type || 'text') + '"' + commonAttr + (hint ? (' placeholder="' + esc(hint) + '"') : '') + ' />');
    } else {
      // unknown → text
      html.push('<label for="' + esc(id) + '">' + esc(label) + '</label>');
      html.push('<input type="text"' + commonAttr + (hint ? (' placeholder="' + esc(hint) + '"') : '') + ' />');
    }
    html.push('</div>');
  }

  if (root && root.title === 'ROOT') {
    for (var k=0;k<root.children.length;k++) renderNode(root.children[k], 0);
  } else if (root) {
    renderNode(root, 0);
  }
  html.push('</div>');
  return html.join('');
}

/**
 * STRUCTUREシートからフォームHTMLを生成（シグネチャ比較＋キャッシュ）
 * @param {{refresh?:boolean}} opts
 * @return {string} HTML
 */
function generateFormFromStructureSheet(opts) {
  opts = opts || {};
  var refresh = !!opts.refresh;
  var sig = getStructureSignature_();
  var cacheKey = 'form:' + sig;
  var cache = CacheService.getScriptCache();
  if (!refresh) {
    try {
      var cached = cache.get(cacheKey);
      if (cached) return cached;
    } catch (_e) {}
  }

  // 生成処理
  var tree = parseStructureSheet();
  tree = computeAndValidateNodeIds_(tree);
  var html = generateFormHtmlFromStructureTree_(tree);

  // 保存
  try { cache.put(cacheKey, html, 21600); } catch (_e2) {}
  try { PropertiesService.getScriptProperties().setProperty('STRUCTURE_SIGNATURE', sig); } catch(_e3) {}
  return html;
}
