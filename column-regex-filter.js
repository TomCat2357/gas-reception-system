/*
 * 列別フィルタ + 正規表現 + OR/AND + 日時比較（フロントエンド）
 *
 * 使い方（最小例）:
 *   // Google Apps Scriptでは関数を直接呼び出し
 *   const compiled = compileQuery(searchInput.value, { defaultFlags: 'i', normalize: true });
 *   const filtered = filterRows(allData, compiled, { fieldAliases: DEFAULT_FIELD_ALIASES });
 *   renderTable(filtered);
 *
 * クエリ構文:
 *   1) フィールド指定:  ステータス:完了
 *   2) 正規表現:        ID:/^REC-.*\/   作成日時:2024-08
 *   3) グローバル語句:  フィールド名が無い語は"どこかの列"にマッチで採用（正規表現も可）
 *   4) スペースを含む値: ステータス:"対応中"   個別案件タグ:/緊急\s+対応/
 *   5) AND/OR:         半角/全角スペース区切りは AND。`AND` と書いても AND。
 *                      句どうしを `OR`（大文字）で区切ると OR 。
 *                      例) ステータス:完了 作成日時:2024-08   ≡   ステータス:完了 AND 作成日時:2024-08
 *                          ステータス:完了 OR ステータス:対応中  →  どちらか満たせばOK
 *   6) 日時比較:        created:>=2024-05-01   updated:<2025-08-01T12:30:00
 *                      日付のみ指定時は 00:00:00 を付与（例: 2024-05-01 → 2024-05-01T00:00:00）
 *
 * マッチの論理:
 *   - クエリは OR で区切られた「句」の集合。句同士は OR。
 *   - 各句の中は AND（フィールド指定もグローバル語句も AND）。
 *   - 同一フィールドに複数条件が出た場合も AND（ステータス:完了 ステータス:/対応/）。
 *
 * 備考:
 *   - デフォルトで大文字小文字無視（flags = 'i'）。/pattern/ のように明示した場合はそのフラグを優先
 *   - 日本語向けに NFKC 正規化（全角/半角・一部合成文字）を軽く実施（必要に応じて拡張可能）
 */

// フィールドのエイリアス（key = 内部名、value = 外部名の配列）
const DEFAULT_FIELD_ALIASES = {
  作成日時: ['作成日時', '作成日', 'created'],
  更新日時: ['更新日時', '更新日'],
};

// 日時比較対象（正規化後のカノニカル名）
const DATE_FIELD_CANONICALS = new Set(['作成日時', '更新日時']);

function compileQuery(input, opts) {
  opts = opts || {};
  var defaultFlags = opts.defaultFlags || 'i';
  var normalize = opts.normalize !== false;
  const tokens = tokenize(input || '');
  const clauses = [];
  let cur = newClause(normalize);

  var pushIfNotEmpty = function() {
    if (cur.globals.length || Object.keys(cur.fields).length) clauses.push(cur);
  };

  for (var i = 0; i < tokens.length; i++) {
    var t = tokens[i];
    // 単独 AND は無視（空白と同義）、単独 OR は句区切り
    if (!t.key && t.value === 'AND') continue;
    if (!t.key && t.value === 'OR') { pushIfNotEmpty(); cur = newClause(normalize); continue; }

    if (t.key) {
      var canonical = toCanonicalFieldKey(t.key);
      // 日時フィールドかつ「比較」構文なら日時コンパイル
      if (DATE_FIELD_CANONICALS.has(canonical)) {
        var cmp = toDateComparator(t.value);
        if (cmp) {
          if (!cur.fields[canonical]) cur.fields[canonical] = [];
          cur.fields[canonical].push(cmp); // {__type:'date', op, ts}
          continue;
        }
      }
      var re = toRegExp(t.value, defaultFlags);
      if (re) {
        if (!cur.fields[canonical]) cur.fields[canonical] = [];
        cur.fields[canonical].push(re);
      }
    } else {
      var re = toRegExp(t.value, defaultFlags);
      if (re) cur.globals.push(re);
    }
  }
  pushIfNotEmpty();

  return { clauses, normalize };
}

function filterRows(rows, compiled, options) {
  options = options || {};
  var aliases = options.fieldAliases || DEFAULT_FIELD_ALIASES;
  var norm = (compiled && compiled.normalize ? normalizeJa : function(x) { return (x && x.toString) ? x.toString() : ''; });
  var allFieldKeys = Object.keys(aliases).filter(function(k) { return !DATE_FIELD_CANONICALS.has(k); }); // 日付以外

  // 句が無い（空入力）→ 全件
  if (!compiled || !Array.isArray(compiled.clauses) || compiled.clauses.length === 0) return rows || [];

  return (rows || []).filter((row) => {
    // どれかの句が真なら採用（OR）
    return compiled.clauses.some(function(clause) {
      // 1) フィールド縛り（AND）
      var fieldEntries = Object.entries ? Object.entries(clause.fields) : Object.keys(clause.fields).map(function(k) { return [k, clause.fields[k]]; });
      for (var j = 0; j < fieldEntries.length; j++) {
        var fieldKey = fieldEntries[j][0];
        var patterns = fieldEntries[j][1];
        var text = norm(getFieldValue(row, aliases, fieldKey));
        var ok = patterns.every(function(p) {
          if (p && p.__type === 'date') {
            var v = getFieldValue(row, aliases, fieldKey);
            var ts = toDateEpoch(v); // 行側日時
            if (ts == null) return false;
            return compareTs(ts, p.op, p.ts);
          }
          return safeTest(p, text);
        });
        if (!ok) return false;
      }
      // 2) グローバル語句（各語句は「どこかの列」に当たればOK）→ 句内 AND
      for (var k = 0; k < clause.globals.length; k++) {
        var re = clause.globals[k];
        var hit = false;
        for (var l = 0; l < allFieldKeys.length; l++) {
          var fk = allFieldKeys[l];
          var text2 = norm(getFieldValue(row, aliases, fk));
          if (safeTest(re, text2)) { hit = true; break; }
        }
        if (!hit) return false;
      }
      return true; // 句成立
    });
  });
}

// -------------------- 内部ヘルパー --------------------
function getFieldValue(row, aliases, fieldKey) {
  var keys = aliases[fieldKey] || [];
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    if (k in row) return row[k];
    var rowKeys = Object.keys(row);
    for (var j = 0; j < rowKeys.length; j++) {
      var kk = rowKeys[j];
      if (kk.toLowerCase() === k.toLowerCase()) return row[kk];
    }
  }
  return '';
}

function normalizeJa(s) {
  if (s == null) return '';
  try {
    return s.toString().normalize('NFKC');
  } catch {
    return String(s);
  }
}

function safeTest(re, text) {
  try { return re.test(text); } catch { return false; }
}

function toCanonicalFieldKey(key) {
  var k = String(key || '').toLowerCase();
  if (DEFAULT_FIELD_ALIASES[k]) return k;
  // エイリアス → カノニカル
  var entries = Object.entries ? Object.entries(DEFAULT_FIELD_ALIASES) : Object.keys(DEFAULT_FIELD_ALIASES).map(function(canon) { return [canon, DEFAULT_FIELD_ALIASES[canon]]; });
  for (var i = 0; i < entries.length; i++) {
    var canon = entries[i][0];
    var arr = entries[i][1] || [];
    for (var j = 0; j < arr.length; j++) {
      if (String(arr[j]).toLowerCase() === k) return canon;
    }
  }
  return k;
}

function toRegExp(value, defaultFlags) {
  defaultFlags = defaultFlags || 'i';
  var v = value.trim();
  if (!v) return null;
  // /.../flags の形式を検出
  if (v.indexOf('/') === 0 && v.lastIndexOf('/') > 0) {
    var lastSlash = v.lastIndexOf('/');
    var pattern = v.slice(1, lastSlash);
    var flags = v.slice(lastSlash + 1);
    try { return new RegExp(pattern, flags || defaultFlags); } catch { return null; }
  }
  // 通常のリテラル文字列
  try { return new RegExp(v, defaultFlags); } catch {
    var escaped = v.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    try { return new RegExp(escaped, defaultFlags); } catch { return null; }
  }
}

function tokenize(s) {
  // スペース区切り（半角/全角対応）。ただし "..." /'...'/ /.../ は 1 トークン。
  // 戻り値: { key?: string, value: string }[]
  var out = [];
  var i = 0, n = s.length;
  var buf = ''; var key = null;
  var inQuote = false; var quoteChar = null;
  var inRegex = false; var escape = false;

  var push = function() {
    var val = buf.trim();
    if (val) out.push({ key: key, value: val });
    buf = ''; key = null; inQuote = false; quoteChar = null; inRegex = false; escape = false;
  };

  while (i < n) {
    var c = s[i];
    if (inQuote) {
      if (escape) { buf += c; escape = false; }
      else if (c === '\\') { escape = true; }
      else if (c === quoteChar) { inQuote = false; }
      else { buf += c; }
    } else if (inRegex) {
      if (escape) { buf += c; escape = false; }
      else if (c === '\\') { escape = true; }
      else if (c === '/') { // 正規表現の終端
        // ここでは本体のみ格納。flags は toRegExp で拾うため末尾で再結合
        // flags を読み取る（/.../flags）
        var j = i + 1; var flags = '';
        while (j < n && /[gimsuy]/.test(s[j])) { flags += s[j]; j++; }
        buf = '/' + buf + '/' + flags; i = j - 1; inRegex = false;
      } else { buf += c; }
    } else {
      if (c === ' ' || c === '\t' || c === '\n' || c === '\u3000') { // 全角スペース対応
        if (buf.length) push();
      } else if (c === '"' || c === "'") {
        inQuote = true; quoteChar = c;
      } else if (c === '/' && (buf.slice(-1) === ':' || buf.length === 0)) {
        inRegex = true; // /.../ を開始
      } else if (c === ':' && key == null) {
        key = buf.trim().toLowerCase(); buf = '';
      } else {
        buf += c;
      }
    }
    i++;
  }
  if (buf.length) push();
  return out;
}

// ---- 日時比較：コンパイル/評価 ----
function toDateComparator(raw) {
  var s = String(raw || '').trim();
  var m = s.match(/^(<=|>=|<|>|!?=)?\s*(.+)$/);
  if (!m) return null;
  var op = (m[1] || '=').replace('==', '=');
  var ts = toDateEpoch(m[2]);
  if (ts == null) return null;
  return { __type: 'date', op: op, ts: ts };
}

function toDateEpoch(v) {
  if (v == null) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v.getTime();
  if (typeof v === 'number') return isFinite(v) ? v : null;
  var s = String(v).trim();
  if (!s) return null;
  // 正規化（全角→半角）
  var n = normalizeJa(s);
  // 形式: YYYY-MM-DD( / YYYY/MM/DD ) [T| ]HH:mm[:ss]
  var m = n.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})(?:[ T](\d{1,2})(?::(\d{1,2})(?::(\d{1,2}))?)?)?$/);
  if (m) {
    var Y = +m[1], M = +m[2]-1, D = +m[3];
    var hh = m[4] ? +m[4] : 0;
    var mm = m[5] ? +m[5] : 0;
    var ss = m[6] ? +m[6] : 0;
    return new Date(Y, M, D, hh, mm, ss).getTime(); // ローカル 00:00:00 既定
  }
  // ISO等は Date.parse に委ねる
  var t = Date.parse(n);
  return isNaN(t) ? null : t;
}

function compareTs(lhs, op, rhs) {
  switch (op) {
    case '<':  return lhs <  rhs;
    case '<=': return lhs <= rhs;
    case '>':  return lhs >  rhs;
    case '>=': return lhs >= rhs;
    case '!=': return lhs !== rhs;
    case '=':  return lhs === rhs;
    default:   return false;
  }
}

// ---- 句（clause）ユーティリティ ----
function newClause(normalize) {
  return { globals: [], fields: {}, normalize };
}

// ---- 入力デバウンス ----
function debounce(fn, wait) {
  wait = wait || 200;
  var t = null;
  return function debounced() {
    var ctx = this;
    var args = Array.prototype.slice.call(arguments);
    if (t) clearTimeout(t);
    t = setTimeout(function() { t = null; fn.apply(ctx, args); }, wait);
  };
}

/*
 * ---- 配線例（既存コード差し替えの目安） ----
 * 
 * // 1) 起動時などに全件取得済みとする
 * let allData = [];
 * let filteredData = [];
 * 
 * // 2) 入力イベントでフィルタ
 * const inputEl = document.getElementById('searchInput');
 * const onInput = debounce(function(e) {
 *   const compiled = compileQuery(e.target.value, { defaultFlags: 'i', normalize: true });
 *   filteredData = filterRows(allData, compiled, { fieldAliases: DEFAULT_FIELD_ALIASES });
 *   renderTable(filteredData);
 *   if (typeof updateDataCount === 'function') updateDataCount(filteredData.length, allData.length);
 * }, 200);
 * inputEl.addEventListener('input', onInput);
 * 
 * // 3) 列キーが違う場合は DEFAULT_FIELD_ALIASES を実データに合わせて調整してください
 */
