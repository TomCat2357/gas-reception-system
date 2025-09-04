/**
 * データ入力フォーム機能
 * Google AppSheetスタイルのWebアプリケーション
 */

// ======= Webアプリケーション関連 =======

// Webアプリのメインページを表示
function doGet(e) {
  // パラメータが存在しない場合のエラーハンドリング
  if (!e || !e.parameter) {
    console.log('doGet called without parameters - returning main page');
    return HtmlService.createTemplateFromFile('webapp')
      .evaluate()
      .setTitle('📊 データ管理アプリ')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  // デバッグモードのチェック
  const page = e.parameter.page || 'main';
  
  if (page === 'debug') {
    return HtmlService.createTemplateFromFile('debug')
      .evaluate()
      .setTitle('🔧 デバッグページ - データ管理アプリ')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'reception') {
    var file = 'reception_form';
    return HtmlService.createTemplateFromFile(file)
      .evaluate()
      .setTitle('📝 受付入力フォーム')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'csv_converter') {
    return HtmlService.createTemplateFromFile('reception_form')
      .evaluate()
      .setTitle('📄 CSV to JSON フォーム定義変換')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'test_csv') {
    return HtmlService.createTemplateFromFile('test_csv_converter')
      .evaluate()
      .setTitle('🧪 CSV to JSON 変換テスト')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  return HtmlService.createTemplateFromFile('webapp')
    .evaluate()
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
  const htmlOutput = HtmlService.createTemplateFromFile('simple_test')
    .evaluate()
    .setWidth(600)
    .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🔧 簡単テスト');
}

// ======= ユーティリティ関数 =======


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
    
    // Test 7: parseCSVFormDefinition - 仕様書正常例
    var csvNormal = 'L1,L2,L3,L4,L5,L6,L7,L8,L9\n' +
                   'Block_B1,,,,,,,,\n' +
                   'X/display:満年齢,,,,,,,,\n' +
                   ',,S1/selector:RADIO,,,,,,,\n' +
                   ',,S2/selector:RADIO,,,,,,,\n' +
                   ',,S3/selector:RADIO,T1/selector:RADIO,,,,,,\n' +
                   ',,S3/selector:RADIO,T2/selector:RADIO,text/re:^.{0,20}$,,,,';
    
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
