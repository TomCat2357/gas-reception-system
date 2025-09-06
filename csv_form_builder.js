/**
 * CSV→JSON→HTML フォーム生成（フォームビルダー用）
 * サーバ検証（Code.js/parseCSVFormDefinition）の仕様に準拠。
 * - 列は L1..L9 の階層（カンマ区切り行）
 * - 各セルは `title/type/hint` を '/' 区切りで記述
 * - selector は `selector:RADIO|CHECKBOX|DROPDOWN` のみ
 * - 同一親直下の selector グループは、全て同一種別の selector のみ（混在禁止）
 */

// --------- Utility: delimiter sniffing and CSV parsing ---------
function fb_sniffDelimiter(text) {
  if (!text) return ',';
  var firstLine = String(text).split(/\r?\n/)[0] || '';
  var candidates = ['\t', ';', ',', '/'];
  var counts = candidates.map(function(d){ return [d, (firstLine.match(new RegExp('\\' + d, 'g')) || []).length]; });
  counts.sort(function(a,b){ return b[1]-a[1]; });
  return counts[0][1] > 0 ? counts[0][0] : ',';
}

function fb_parseCsv(text, delimiter) {
  text = text || '';
  delimiter = delimiter || fb_sniffDelimiter(text);
  var lines = String(text).replace(/\r\n/g,'\n').replace(/\r/g,'\n').split('\n');
  var rows = [];
  for (var li=0; li<lines.length; li++) {
    var line = lines[li];
    if (!line || !line.trim()) continue;
    rows.push(fb_parseCsvLine(line, delimiter));
  }
  return rows;
}

function fb_parseCsvLine(line, delimiter) {
  var out = [];
  var i = 0, n = line.length;
  var buf = '';
  var inQuotes = false; var prevWasQuote = false;
  while (i < n) {
    var c = line[i];
    if (inQuotes) {
      if (c === '"') {
        if (line[i+1] === '"') { buf += '"'; i++; }
        else { inQuotes = false; }
      } else { buf += c; }
    } else {
      if (c === '"') { inQuotes = true; }
      else if (c === delimiter) { out.push(buf); buf = ''; }
      else { buf += c; }
    }
    i++;
  }
  out.push(buf);
  return out;
}

// --------- Type normalization ---------
// 型はサーバ仕様に正規化（selector/re/display のみ）
function fb_normalizeType(raw) {
  var s = String(raw || '').trim();
  if (!s) return null;
  // 旧来の別名を互換的に受理しつつ selector: に正規化
  var m = s.toUpperCase();
  if (m === 'RADIO' || m === 'SELECTOR:RADIO') return 'selector:RADIO';
  if (m === 'CHECKBOX' || m === 'SELECTOR:CHECKBOX') return 'selector:CHECKBOX';
  if (m === 'SELECT' || m === 'DROPDOWN' || m === 'SELECTOR:DROPDOWN') return 'selector:DROPDOWN';
  if (s.startsWith('selector:RADIO') || s.startsWith('selector:CHECKBOX') || s.startsWith('selector:DROPDOWN')) return s;
  if (s.startsWith('re:')) return s; // re:...
  if (s.startsWith('display:')) return s; // display:...
  return s; // その他はそのまま（レンダ時に扱い）
}

// --------- CSVセルを title/type/hint に分割（'/' 区切り） ---------
function fb_escapeSlashes(value) { if (!value || typeof value !== 'string') return value; return value.replace(/\\\//g,'/'); }
function fb_parseCSVCell(cellValue) {
  if (!cellValue || typeof cellValue !== 'string') return { title: null, type: null, hint: null };
  var unescaped = fb_escapeSlashes(cellValue);
  var parts = unescaped.split('/');
  return { title: parts[0] || null, type: fb_normalizeType(parts[1] || ''), hint: parts[2] || null };
}

// --------- L1..L9 配列からノードツリーを構築（サーバ準拠） ---------
function fb_getNodeKey(row, col, title, type){ return row + ':' + col + ':' + (title || '') + ':' + (type || ''); }
function fb_buildNodeTree(parsedRows){
  if (!parsedRows || parsedRows.length === 0) throw new Error('パース済みデータが空です');
  var nodes = new Map();
  var nodesByParent = new Map();
  for (var rowIdx=0; rowIdx<parsedRows.length; rowIdx++){
    var row = parsedRows[rowIdx];
    for (var colIdx=0; colIdx<9; colIdx++){
      var cell = row[colIdx];
      if (!cell || !cell.title) continue;
      var parentNode=null, parentKey=null;
      if (colIdx>0){
        // 同行の左列
        for (var leftCol=colIdx-1; leftCol>=0; leftCol--){
          if (row[leftCol] && row[leftCol].title){
            parentKey = fb_getNodeKey(rowIdx, leftCol, row[leftCol].title, row[leftCol].type);
            parentNode = nodes.get(parentKey);
            break;
          }
        }
        // 左列に親がない場合、上行の左列を遡る
        if (!parentNode && colIdx>0){
          for (var upRow=rowIdx-1; upRow>=0; upRow--){
            var upRowData = parsedRows[upRow];
            if (upRowData[colIdx-1] && upRowData[colIdx-1].title){
              parentKey = fb_getNodeKey(upRow, colIdx-1, upRowData[colIdx-1].title, upRowData[colIdx-1].type);
              parentNode = nodes.get(parentKey);
              break;
            }
          }
        }
      }
      var nodeKey = fb_getNodeKey(colIdx, cell.title, cell.type);
      // fix: include rowIdx in key
      nodeKey = fb_getNodeKey(rowIdx, colIdx, cell.title, cell.type);
      var existing = nodes.get(nodeKey);
      if (!existing){
        var newNode = { title: cell.title, children: [] };
        if (cell.type) newNode.type = cell.type;
        if (cell.hint) newNode.hint = cell.hint;
        nodes.set(nodeKey, newNode);
        if (parentNode){
          if (!nodesByParent.has(parentKey)) nodesByParent.set(parentKey, []);
          nodesByParent.get(parentKey).push(newNode);
          parentNode.children.push(newNode);
        }
      }
    }
  }
  // 親を持たないノード＝ルート候補
  var rootNodes=[];
  for (var kv of nodes){
    var node = kv[1];
    var hasParent=false;
    for (var byp of nodesByParent.values()){
      if (byp.includes(node)){ hasParent=true; break; }
    }
    if (!hasParent) rootNodes.push(node);
  }
  if (rootNodes.length===0) throw new Error('ルートノードが見つかりません');
  if (rootNodes.length===1) return rootNodes[0];
  return { title:'ROOT', children: rootNodes };
}

// --------- バリデーション（サーバ準拠） ---------
function fb_validateFormStructure(nodeTree){
  var errors=[];
  function walk(node, path){
    if (!node) return;
    var currentPath = path ? path + ' > ' + node.title : node.title;
    if (node.type){
      var validTypes = /^(selector:(RADIO|CHECKBOX|DROPDOWN)|re:.+|display:.+)$/;
      if (!validTypes.test(node.type)) errors.push('無効な型: '+node.type+' at '+currentPath);
    }
    if (node.children && node.children.length>0){
      var selectorChildren = node.children.filter(function(c){ return c.type && c.type.startsWith('selector:'); });
      if (selectorChildren.length>0){
        var selectorTypes = selectorChildren.map(function(c){ return c.type; });
        var uniqueTypes = selectorTypes.filter(function(t, i){ return selectorTypes.indexOf(t)===i; });
        if (uniqueTypes.length>1) errors.push('選択肢グループ内でselector種別が混在: '+uniqueTypes.join(', ')+' at '+currentPath);
        var nonSelectorChildren = node.children.filter(function(c){ return !c.type || !c.type.startsWith('selector:'); });
        if (nonSelectorChildren.length>0) errors.push('選択肢グループに非selector要素が混入 at '+currentPath);
      }
      node.children.forEach(function(ch){ walk(ch, currentPath); });
    }
  }
  walk(nodeTree, '');
  return { valid: errors.length===0, errors: errors };
}

// --------- CSV -> ノードツリー（サーバ準拠の簡易版） ---------
function fb_parseCSVFormDefinition(csvContent){
  try{
    if (!csvContent || typeof csvContent !== 'string') throw new Error('CSVコンテンツが無効です');
    var lines = csvContent.split('\n').filter(function(line){ return line.trim().length>0; });
    if (lines.length===0) throw new Error('CSVデータが空です');
    var headerLine = lines[0];
    if (!headerLine.includes('L1') || !headerLine.includes('L9')) throw new Error('CSVヘッダーが正しくありません。L1,L2,...,L9の形式である必要があります');
    var parsedRows=[];
    for (var i=1; i<lines.length; i++){
      var cells = fb_parseCsvLine(lines[i].trim(), ',');
      while (cells.length<9) cells.push('');
      var parsedRow=[];
      for (var j=0;j<9;j++){
        var cellValue = cells[j] ? cells[j].trim() : '';
        if (cellValue.startsWith('"') && cellValue.endsWith('"')){ cellValue = cellValue.slice(1,-1).replace(/""/g,'"'); }
        parsedRow.push(fb_parseCSVCell(cellValue));
      }
      parsedRows.push(parsedRow);
    }
    if (parsedRows.length===0) throw new Error('有効なデータ行がありません');
    var nodeTree = fb_buildNodeTree(parsedRows);
    var validation = fb_validateFormStructure(nodeTree);
    if (!validation.valid) throw new Error('バリデーションエラー: ' + validation.errors.join(', '));
    return { success:true, data: nodeTree };
  }catch(e){ return { success:false, error: e.message || String(e) }; }
}

// --------- CSV -> MODEL ---------
function convertFormCsvToModel(csvText) {
  // サーバ準拠のノードツリーモデルを返す
  var parsed = fb_parseCSVFormDefinition(csvText || '');
  if (!parsed.success) return { success:false, error: parsed.error };
  return { success:true, model: parsed.data };
}

// --------- MODEL -> HTML ---------
function generateFormHtml(model) {
  try {
    // model はノードツリー（title, type?, hint?, children[]）
    model = model || { title:'ROOT', children: [] };
    var esc = function(s){ return String(s == null ? '' : s).replace(/[&<>]/g, function(ch){ return ({'&':'&amp;','<':'&lt;','>':'&gt;'}[ch]); }); };
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
      '.flow{border-top:1px dashed var(--border);padding-top:10px;margin-top:10px;}',
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
    html.push('<title>フォーム生成 (template.csv より)</title>');
    html.push('<style>' + css + '</style></head>');
    html.push('<body><div class="container">');
    html.push('<div class="header"><div class="title">フォーム生成 (template.csv より)</div><div class="actions"><button class="btn" id="btn-reset">リセット</button><button class="btn" id="btn-export">入力値をJSONで取得</button></div></div>');
    html.push('<div class="blocks" id="blocks">');

    var idSeq = 0;
    function renderNode(node, depth){
      depth = depth || 0;
      var hasChildren = Array.isArray(node.children) && node.children.length>0;
      var indent = Math.max(0, depth) * 16;

      // depth=0 はブロック（L1）で区切る
      if (depth === 0){
        html.push('<div class="block">');
        if (node.title){ html.push('<h2>'+esc(node.title)+'</h2>'); }
        if (node.hint){ html.push('<div class="label">'+esc(node.hint)+'</div>'); }
        (node.children||[]).forEach(function(ch){ renderNode(ch, depth+1); });
        html.push('</div>');
        return;
      }
      // 子に selector があり（同種・2つ以上）かつ非selectorが無い場合はグループ描画
      if (hasChildren){
        var selectorChildren = node.children.filter(function(c){ return c.type && c.type.startsWith('selector:'); });
        var nonSelectorChildren = node.children.filter(function(c){ return !c.type || !c.type.startsWith('selector:'); });
        var uniq = selectorChildren.map(function(c){ return c.type; }).filter(function(t,i,a){ return a.indexOf(t)===i; });
        if (selectorChildren.length>=2 && nonSelectorChildren.length===0 && uniq.length===1){
          var kind = uniq[0];
          var groupId = 'g_'+(idSeq++);
          html.push('<div class="step" style="margin-left:'+indent+'px">');
          html.push('<div class="label">'+esc(node.title || '') + (node.hint? ' <span class="hint-badge" title="'+esc(node.hint)+'">hint</span>':'' ) + '</div>');
          if (kind==='selector:RADIO'){
            var name='r_'+groupId;
            html.push('<div class="options" data-group="'+groupId+'">');
            node.children.forEach(function(opt){
              var optId = 'opt_'+(idSeq++);
              var contId = 'cont_'+optId;
              html.push('<label class="chip"><input type="radio" name="'+name+'" value="'+esc(opt.title||'')+'" data-role="opt-radio" data-group="'+groupId+'" data-target="'+contId+'">'+esc(opt.title||'')+'</label>');
              // container for opt children
              html.push('<div id="'+contId+'" class="opt-children" data-group="'+groupId+'" style="display:none;">');
              // render nested groups starting from this opt node
              renderNode(opt, (depth+1));
              html.push('</div>');
            });
            html.push('</div>');
          } else if (kind==='selector:CHECKBOX'){
            var cbGroup = groupId;
            html.push('<div class="options" data-group="'+cbGroup+'">');
            node.children.forEach(function(opt){
              var optId = 'opt_'+(idSeq++);
              var contId = 'cont_'+optId;
              html.push('<label class="chip"><input type="checkbox" value="'+esc(opt.title||'')+'" data-role="opt-checkbox" data-group="'+cbGroup+'" data-target="'+contId+'">'+esc(opt.title||'')+'</label>');
              html.push('<div id="'+contId+'" class="opt-children" data-group="'+cbGroup+'" style="display:none;">');
              renderNode(opt, (depth+1));
              html.push('</div>');
            });
            html.push('</div>');
          } else if (kind==='selector:DROPDOWN'){
            var selId = 'sel_'+groupId;
            html.push('<select id="'+selId+'" data-role="opt-select" data-group="'+groupId+'">');
            node.children.forEach(function(opt){
              var optId = 'opt_'+(idSeq++);
              var contId = 'cont_'+optId;
              html.push('<option value="'+contId+'">'+esc(opt.title||'')+'</option>');
            });
            html.push('</select>');
            // containers in same order
            node.children.forEach(function(opt){
              var optId = 'opt_'+(idSeq++);
              var contId = 'cont_'+optId;
              html.push('<div id="'+contId+'" class="opt-children" data-group="'+groupId+'" style="display:none;">');
              renderNode(opt, (depth+1));
              html.push('</div>');
            });
          }
          html.push('</div>');
          return;
        }
      }
      // 非selector or 混在: 見出し/セクションとして描画し子を再帰
      if (node.title && (!node.type || node.type.startsWith('display:'))){
        html.push('<div class="step" style="margin-left:'+indent+'px">');
        html.push('<div class="label">' + esc(node.title) + (node.hint? ' <span class="hint-badge" title="'+esc(node.hint)+'">hint</span>':'' ) + '</div>');
        html.push('</div>');
        (node.children||[]).forEach(function(ch){ renderNode(ch, depth+1); });
        return;
      }
      // re: 入力項目
      if (node.type && node.type.startsWith('re:')){
        html.push('<div class="step" style="margin-left:'+indent+'px">');
        html.push('<div class="label">'+esc(node.title||'') + (node.hint? ' <span class="hint-badge" title="'+esc(node.hint)+'">hint</span>':'' ) + '</div>');
        html.push('<input type="text">');
        html.push('</div>');
        (node.children||[]).forEach(function(ch){ renderNode(ch, depth+1); });
        return;
      }
      // それ以外は子のみ描画
      (node.children||[]).forEach(function(ch){ renderNode(ch, depth+1); });
    }

    // ルート処理
    if (model.title==='ROOT' && Array.isArray(model.children)){
      model.children.forEach(function(ch){ renderNode(ch, 0); });
    } else {
      renderNode(model, 0);
    }

    // 付加: 最小のインラインJSで [リセット] と [JSON出力] を動作
    html.push('</div>');
    html.push('<script>(function(){\n' +
      'function getLabelText(label){ if(!label) return ""; var c=label.cloneNode(true); var badges=c.querySelectorAll(\'.hint-badge\'); for(var i=0;i<badges.length;i++){badges[i].remove();} return (c.textContent||\'\').trim(); }\n' +
      'function resetAll(){ var root=document.getElementById(\'blocks\'); if(!root) return; var inputs=root.querySelectorAll(\'input, textarea, select\'); for(var i=0;i<inputs.length;i++){ var el=inputs[i]; var t=(el.type||\'\').toLowerCase(); if(t===\'radio\'||t===\'checkbox\'){ el.checked=false; } else if(el.tagName===\'SELECT\'){ el.selectedIndex=0; } else { el.value=\'\'; } } }\n' +
      // 動的表示のバインド
      'function hideGroup(group){ var nodes=document.querySelectorAll(\'.opt-children[data-group="\'+group+\'"]\'); for(var i=0;i<nodes.length;i++){ nodes[i].style.display=\'none\'; } }\n' +
      'function bindToggles(){\n' +
      '  // radio\n' +
      '  var radios=document.querySelectorAll(\'input[data-role=opt-radio]\');\n' +
      '  for(var i=0;i<radios.length;i++){ radios[i].addEventListener(\'change\', function(e){ var g=this.getAttribute(\'data-group\'); hideGroup(g); var t=this.getAttribute(\'data-target\'); var el=document.getElementById(t); if(el) el.style.display=this.checked?\'block\':\'none\'; }); }\n' +
      '  // checkbox\n' +
      '  var cbs=document.querySelectorAll(\'input[data-role=opt-checkbox]\');\n' +
      '  for(var j=0;j<cbs.length;j++){ cbs[j].addEventListener(\'change\', function(){ var t=this.getAttribute(\'data-target\'); var el=document.getElementById(t); if(el) el.style.display=this.checked?\'block\':\'none\'; }); }\n' +
      '  // select\n' +
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
    return { success: true, html: html.join('') };
  } catch (e) {
    return { success: false, error: e.message || String(e) };
  }
}

// --------- Facade for client ---------
function generateFormFromCsv(csvText) {
  var m = convertFormCsvToModel(csvText || '');
  if (!m.success) return m;
  var h = generateFormHtml(m.model);
  if (!h.success) return h;
  return { success: true, model: m.model, html: h.html };
}
