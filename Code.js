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
  if (Array.isArray(obj)) {
    for (var i=0;i<obj.length;i++) {
      out.push.apply(out, flattenToPaths_(obj[i], prefix.concat(String(i+1))));
    }
  } else if (isObj(obj)) {
    for (var k in obj) {
      if (Object.prototype.hasOwnProperty.call(obj,k)) {
        out.push.apply(out, flattenToPaths_(obj[k], prefix.concat(k)));
      }
    }
  } else {
    out.push([prefix, obj]);
  }
  return out;
}

// ヘッダーを作成（4段固定）: 各段すべてのセルにラベルを書き込む（マージ代替）
function createNestedHeaders_(sheet, headerPaths) {
  if (!headerPaths || headerPaths.length === 0) return 0;
  // 仕様書に従い4段固定 (L1, L2, L3, L4)
  var depth = 4;
  var cols = headerPaths.length;
  var values = [];
  for (var r=0;r<depth;r++) {
    var row = [];
    for (var c=0;c<cols;c++) {
      // パスが4段未満の場合は空文字で埋める
      row.push(headerPaths[c][r] || '');
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
  for (var c2=1;c2<=cols;c2++) sheet.setColumnWidth(c2, 180);
  return depth;
}

// 既存ヘッダー（4段固定）を取得
function readHeaderPaths_(sheet, headerRows) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0 || headerRows === 0) return [];
  // 4段固定に対応：headerRowsが4以外でも4段として読み取る
  var actualRows = Math.max(headerRows, 4);
  var vals = sheet.getRange(1,1,actualRows,lastCol).getValues();
  var paths = [];
  for (var c=0;c<lastCol;c++) {
    var path = [];
    for (var r=0;r<4;r++) { // 4段固定で読み取り
      var cell = (vals[r] && vals[r][c]) ? vals[r][c].toString() : '';
      path.push(cell);
    }
    // 末尾空串を削除
    while (path.length>0 && path[path.length-1]==='') path.pop();
    paths.push(path);
  }
  return paths;
}

// パスをカノニカル文字列へ（キー）
function pathKey_(path){ return path.join('›'); }

// パス配列の辞書順ソート（安定）
function sortPaths_(paths){ return paths.slice().sort(function(a,b){ var A=pathKey_(a), B=pathKey_(b); return A<B?-1:A>B?1:0; }); }

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
      headerRows = createNestedHeaders_(sheet, headerPaths);
    }

    // 現在のヘッダーを取得
    var current = readHeaderPaths_(sheet, headerRows);
    var keyToCol = {};
    for (var i=0;i<current.length;i++) keyToCol[pathKey_(current[i])] = i+1; // 1-based col

    // 新規パスがあれば列追加
    var newPaths = headerPaths.filter(function(p){ return !(pathKey_(p) in keyToCol); });
    if (newPaths.length>0){
      var startCol = sheet.getLastColumn()+1;
      var allPaths = current.concat(newPaths);
      // すべてを再並べ替えして書き直す（列増分を右端に追加）
      var sorted = sortPaths_(allPaths);
      createNestedHeaders_(sheet, sorted); // ヘッダー全体を書き直し
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
          createNestedHeaders_(sheet, sorted2);
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
  var current = readHeaderPaths_(sheet, headerRows);
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
  var current = readHeaderPaths_(sheet, headerRows);
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
