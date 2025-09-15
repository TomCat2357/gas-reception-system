以下を **「変更指示書（for Codex CLI）」** としてそのまま渡してください。
（目的：**Structure→CSV→JSON→HTML** の段階化と、**Form⇄DATAシート**の往復I/Oを分離してテスト可能にする）

---

# 変更指示書（for Codex CLI）

対象リポジトリ：`TomCat2357/gas-reception-system`
対象主要ファイル：`Code.js`（既存を温存しつつ追記）、`README.md`（導線追記）

## 1) 背景 / 目的

* 現状は `STRUCTURE` シートからフォーム生成までが一体化しがちで、段階的なテストがしにくい。
* **要件**：`STRUCTURE → CSV → JSON → HTML(=受付フォーム)` を **S1/S2/S3** の独立機能に分解。
  さらに **Form入力 → JSON → DATA**、**DATA → JSON → Form初期値** を **F1/F2** として分離。
* 目的は **各段階の入出力を個別テスト**できる構成に再編すること（回帰を起こさず現行画面は温存）。

## 2) 変更範囲（スコープ）

* `Code.js` に以下を**追記**（既存関数・既存ルーティングは削除しない）

  * 共通定数/ユーティリティ
  * **S1**: `structureToCsv_()`
  * **S2**: `csvToJson_(csvText)`
  * **S3**: `jsonToReceptionHtml_(fields, initialData)`
  * **F1**: `saveReceptionDataFromJson(json)`
  * **F2**: `getReceptionJsonByRow_(rowIndex)`
  * `doGet(e)` に **新規テスト用ページ**を追加
* `README.md` に新ルートの説明を**追記**

## 3) 用語・前提

* シート：`STRUCTURE`（定義, 10行ヘッダー）／`DATA`（データ本体, 同じ10行ヘッダーを持つ）
* 10行ヘッダー：**1行目=型**（TEXT/NUMBER/DATE/LIST/EXISTENCE など）、**2–10行目=L1…L9**
* dotted key：`L1.L2...` 連結（例：`受付.基本.氏名`）

## 4) 新規エンドポイント（doGetのpageパラメータ）

* `?page=test_structure_to_csv` … **S1** 出力（CSVプレビュー）
* `?page=test_csv_to_json` … **S2** 出力（JSONプレビュー）
* `?page=test_json_to_form` … **S3** 出力（空のフォーム）
* `?page=reception_from_structure` … **S1→S2→S3** をワンショットで実行してフォーム表示
* `?page=reception_edit&row=2` … **F2** 指定行をフォーム初期値に流し込み（例では2行目）

> 既存ページ（`reception`, `structure_form`, `debug`, `test_csv`）は残すこと。

## 5) 実装タスク（順序）

1. `Code.js` に **定数・ユーティリティ** を追加
2. **S1**: `structureToCsv_` を実装（10×Nヘッダーを列方向に走査し CSV 文字列生成）
3. **S2**: `csvToJson_` を実装（CSV→フィールド配列JSONへ変換）
4. **S3**: `jsonToReceptionHtml_` を実装（フィールド配列JSON→受付フォームHTMLへ）
5. **F1**: `saveReceptionDataFromJson` を実装（dotted key を `DATA` の列にマップして追記保存）
6. **F2**: `getReceptionJsonByRow_` を実装（`DATA` 指定行→ネストJSONへ復元）
7. `doGet` に **新規page** を追加（上記5種）
8. `README.md` に「段階テスト方法」を追記（下記「受け入れ条件／テスト」も参照）

## 6) 追記コード（そのまま貼り付け可）

> **注意**：既存コードと関数名が衝突しないか検索。衝突時は後置アンダースコア名を調整（例：`structureToCsv_v2_`）。
> **既存 `doGet`** がある場合は **switch の case をマージ**（default は既存のまま）。

```js
/***** === Constants === *****/
const STRUCTURE_SHEET_NAME = 'STRUCTURE';
const DATA_SHEET_NAME = 'DATA';
const HEADER_ROWS = 10; // 1: type, 2-10: L1..L9

// type→入力UIの簡易対応（必要に応じ拡張）
const TYPE_TO_WIDGET = {
  'TEXT': 'text',
  'TEXTAREA': 'textarea',
  'NUMBER': 'number',
  'DATE': 'date',
  'EXISTENCE': 'checkbox',
  'LIST': 'select'
};

/***** === Utils === *****/
const escHtml_ = s => String(s ?? '').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const joinPath_ = (a) => a.filter(Boolean).join('.');

function flattenJson_(obj, prefix='', out={}) {
  for (const [k,v] of Object.entries(obj||{})) {
    const key = prefix ? `${prefix}.${k}` : k;
    if (v && typeof v === 'object' && !Array.isArray(v)) flattenJson_(v, key, out);
    else out[key] = v;
  }
  return out;
}
function setByPath_(obj, segs, val){
  let cur = obj;
  for (let i=0;i<segs.length;i++){
    const k = segs[i];
    if (i === segs.length - 1) cur[k] = val;
    else { cur[k] = cur[k] || {}; cur = cur[k]; }
  }
}

/***** === S1: STRUCTURE → CSV === *****/
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
      return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
    }).join(',')
  );
  return lines.join('\n');
}

/***** === S2: CSV → JSON（フォーム仕様） === *****/
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

/***** === S3: JSON → HTML（受付フォーム生成） === *****/
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
        return `<label class="block"><div>${label}</div><textarea name="${name}" id="${id}">${val}</textarea></label>`;
      case 'checkbox': {
        const checked = (String(val) === 'true' || String(val) === '1') ? ' checked' : '';
        return `<label class="block"><input type="checkbox" name="${name}" id="${id}" value="true"${checked}/> ${label}</label>`;
      }
      case 'number':
        return `<label class="block"><div>${label}</div><input type="number" name="${name}" id="${id}" value="${val}"/></label>`;
      case 'date':
        return `<label class="block"><div>${label}</div><input type="date" name="${name}" id="${id}" value="${val}"/></label>`;
      case 'select':
        return `<label class="block"><div>${label}</div><select name="${name}" id="${id}"></select></label>`;
      default:
        return `<label class="block"><div>${label}</div><input type="text" name="${name}" id="${id}" value="${val}"/></label>`;
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
        .withSuccessHandler(() => { alert('保存しました'); })
        .withFailureHandler(err => { alert('保存失敗: ' + err); })
        .saveReceptionDataFromJson(data);
      return false;
    }
  </script>
</body></html>`;
  return html;
}

/***** === F1: JSON → DATAシート（保存） === *****/
function saveReceptionDataFromJson(json) {
  const sh = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET_NAME);
  if (!sh) throw new Error('DATAシートが見つかりません');
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(1, 1, HEADER_ROWS, lastCol).getValues(); // 10xN

  const keyToCol = new Map();
  for (let c=0;c<lastCol;c++){
    const path = [];
    for (let r=1;r<HEADER_ROWS;r++) {
      const v = header[r][c];
      if (v) path.push(String(v));
    }
    if (path.length) keyToCol.set(joinPath_(path), c+1); // 1-based
  }

  const flat = flattenJson_(json);
  const row = new Array(lastCol).fill('');
  for (const [k,v] of Object.entries(flat)) {
    const col = keyToCol.get(k);
    if (col) row[col-1] = (v == null) ? '' : v;
  }
  // 現段階は新規追加（更新は後述TODO）
  sh.appendRow(row);
}

/***** === F2: DATAシート → JSON（指定行を復元） === *****/
function getReceptionJsonByRow_(rowIndex) {
  const sh = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET_NAME);
  if (!sh) throw new Error('DATAシートが見つかりません');
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(1, 1, HEADER_ROWS, lastCol).getValues();
  const values = sh.getRange(rowIndex, 1, 1, lastCol).getValues()[0];

  const out = {};
  for (let c=0;c<lastCol;c++){
    const segs = [];
    for (let r=1;r<HEADER_ROWS;r++){
      const v = header[r][c];
      if (v) segs.push(String(v));
    }
    if (!segs.length) continue;
    setByPath_(out, segs, values[c]);
  }
  return out;
}

/***** === Routing: 新規テストページ群 === *****/
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'webapp';
  switch (page) {
    case 'test_structure_to_csv': {
      const csv = structureToCsv_();
      return HtmlService.createHtmlOutput('<pre>' + escHtml_(csv) + '</pre>');
    }
    case 'test_csv_to_json': {
      const json = csvToJson_(structureToCsv_());
      return HtmlService.createHtmlOutput('<pre>' + escHtml_(JSON.stringify(json, null, 2)) + '</pre>');
    }
    case 'test_json_to_form': {
      const json = csvToJson_(structureToCsv_());
      const html = jsonToReceptionHtml_(json, {}); // 初期値なし
      return HtmlService.createHtmlOutput(html);
    }
    case 'reception_from_structure': {
      const json = csvToJson_(structureToCsv_());
      const html = jsonToReceptionHtml_(json, {}); // 新規
      return HtmlService.createHtmlOutput(html);
    }
    case 'reception_edit': {
      const rowIndex = Number(e.parameter.row || '2'); // 例: 2行目
      const initial = getReceptionJsonByRow_(rowIndex);
      const json = csvToJson_(structureToCsv_());
      const html = jsonToReceptionHtml_(json, initial);
      return HtmlService.createHtmlOutput(html);
    }
    default:
      // 既存の既定ページへ（必要に応じて既存実装と統合）
      return HtmlService.createHtmlOutputFromFile('views/webapp');
  }
}
```

## 7) 受け入れ条件／テスト

* **S1**：`?page=test_structure_to_csv` で CSV が `<pre>` 表示される
* **S2**：`?page=test_csv_to_json` でフィールド配列 JSON が表示される
* **S3**：`?page=test_json_to_form` でフォームが表示され、DOMに input/textarea/select が出現
* **F1**：`?page=reception_from_structure` に入力→送信で `DATA` に1行追記
* **F2**：`?page=reception_edit&row=2` で `DATA` 2行目の内容がフォーム初期値に反映
* 既存ページ（`reception` 等）は従来どおり表示できる（回帰なし）

## 8) 非スコープ（今後の拡張/TODO）

* **更新（upsert）**：`DATA` に `meta.id` などのID列を設け、`id` 有無で `append` / `上書き` を分岐（後段で実装）
* **LIST選択肢**：`STRUCTURE` 側の追加行 or 別シートで `choices` を宣言し `csvToJson_` が `{choices:[...]}` を付与
* **バリデーション**：`required/min/max/pattern` 等を追加行で宣言→`jsonToReceptionHtml_` に反映
* **UI/レイアウト**：スタイル（CSS）・セクション見出し・折り畳み等の装飾は別PR

## 9) リスク／ガードレール

* 既存の `doGet` を**置換しない**（case の追加でマージ）
* 関数名が重複する場合は末尾に `_v2` を付して衝突回避
* 既存 `views/` 配下ファイルは削除・上書きしない
* `STRUCTURE`/`DATA` の 10行ヘッダー前提を崩さない

## 10) README追記（要約）

* 新ルートと段階テスト手順（上記受け入れ条件）を「開発者向けガイド」に追記
* 「STRUCTURE→CSV→JSON→HTML の段階テスト」「Form→DATA 書込」「DATA→Form 初期化」の手順を簡潔に

## 11) コミット計画（小さく分ける）

1. `Code.js`: S1/S2 追加 + `?page=test_structure_to_csv` / `?page=test_csv_to_json`

   * `feat(structure-pipeline): add S1/S2 and test pages`
2. `Code.js`: S3 追加 + `?page=test_json_to_form`

   * `feat(form-gen): add JSON→HTML generator and test page`
3. `Code.js`: F1/F2 追加 + `?page=reception_from_structure` / `?page=reception_edit`

   * `feat(data-io): add form save and row→json load with routes`
4. `README.md`: 開発者向けガイド追記

   * `docs: add staged testing routes and how-to`

---

以上。
