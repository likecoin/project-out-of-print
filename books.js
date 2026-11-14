/* 絕版重生計劃 — 共用書目資料層
 *
 * 讀取試算表「已發佈到網路」公開分頁的 CSV，供首頁的進度摘要與 progress.html
 * 共用，避免兩頁各自保存一份解析邏輯而逐漸不一致。
 *
 * 公開分頁只含書名、作者、出版社、出版年月、狀態、上架連結；工作分頁上的
 * 系列、作者聯絡線索、義工姓名與跟進記錄不會離開試算表。
 */
(function (global) {
  "use strict";

  var CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSJeLBu3-mpLoE-75hVT79uCl6PnOUzNU-jSE8dluYvuZl4jkuNHTOjL7OLVLVx0n4JUYeqsWx3zyop/pub?gid=1498874683&single=true&output=csv";

  // 各欄可接受的寫法；第一項為訊息中使用的名稱。多餘欄位一律忽略。
  var HEADERS = {
    title:     ["書名"],
    author:    ["作者", "作者／編者", "作者/編者", "作者編者"],
    publisher: ["出版社"],
    date:      ["出版年月", "出版年份", "出版日期"],
    status:    ["狀態", "進度"],
    url:       ["上架連結", "上架", "連結"]
  };
  var REQUIRED = ["title", "status"];   // 缺其一則整頁無話可說
  var COLS = {};                        // 對應到試算表實際使用的標題

  // 顯示次序與顏色；未認得的狀態照原文顯示，不會被靜默歸類。
  var STATES = [
    { key: "已上架",     label: "已上架",  color: "#50e3c2" },
    { key: "取得作者授權", label: "已獲授權", color: "#0a7059" },
    { key: "聯絡作者",    label: "聯絡中",  color: "#c89a3c" },
    { key: "",          label: "未開始",  color: "#b8a894" }
  ];

  // 按背景相對亮度選黑或白，日後在試算表新增狀態色仍可讀。
  function readableOn(hex) {
    var m = /^#?([0-9a-f]{6})$/i.exec(hex);
    if (!m) return "#06241e";
    var v = parseInt(m[1], 16);
    var ch = [(v >> 16) & 255, (v >> 8) & 255, v & 255].map(function (c) {
      c = c / 255;
      return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
    });
    var L = 0.2126 * ch[0] + 0.7152 * ch[1] + 0.0722 * ch[2];
    return (1.05 / (L + 0.05)) >= ((L + 0.05) / 0.0598) ? "#ffffff" : "#06241e";
  }

  function stateOf(raw) {
    var s = (raw || "").trim();
    for (var i = 0; i < STATES.length; i++) {
      if (STATES[i].key && s === STATES[i].key) return STATES[i];
    }
    if (s) return { key: s, label: s, color: "#8e7f6d" };
    return STATES[STATES.length - 1];
  }

  // RFC 4180：欄位可加引號，內含逗號、引號與換行。
  function parseCSV(text) {
    var rows = [], row = [], field = "", inQuotes = false, i = 0;
    if (text.charCodeAt(0) === 0xFEFF) i = 1;           // 去除 BOM
    for (; i < text.length; i++) {
      var c = text[i];
      if (inQuotes) {
        if (c === '"') {
          if (text[i + 1] === '"') { field += '"'; i++; }
          else inQuotes = false;
        } else field += c;
      } else if (c === '"') inQuotes = true;
      else if (c === ",") { row.push(field); field = ""; }
      else if (c === "\n") { row.push(field); rows.push(row); row = []; field = ""; }
      else if (c !== "\r") field += c;
    }
    if (field !== "" || row.length) { row.push(field); rows.push(row); }
    return rows;
  }

  function resolveHeaders(head) {
    var missing = [];
    Object.keys(HEADERS).forEach(function (field) {
      var found = null;
      HEADERS[field].forEach(function (name) {
        if (!found && head.indexOf(name) !== -1) found = name;
      });
      COLS[field] = found;
      if (!found && REQUIRED.indexOf(field) !== -1) missing.push(HEADERS[field][0]);
    });
    return missing;
  }

  function knownHeadings() {
    var all = [];
    Object.keys(HEADERS).forEach(function (f) { all = all.concat(HEADERS[f]); });
    return all;
  }

  // 標題列未必是第一列：公開分頁上方可能有標題或空白列。
  function findHeaderRow(rows) {
    var known = knownHeadings(), best = -1, bestScore = 0;
    for (var i = 0; i < rows.length && i < 10; i++) {
      var cells = rows[i].map(function (h) { return (h || "").trim(); });
      var score = cells.filter(function (c) { return c && known.indexOf(c) !== -1; }).length;
      if (score > bestScore) { bestScore = score; best = i; }
    }
    return bestScore >= 2 ? best : (rows.length ? 0 : -1);
  }

  function toObjects(rows) {
    if (!rows.length) return { books: [], missing: ["書名"], found: [] };
    var hi = findHeaderRow(rows);
    var head = rows[hi].map(function (h) { return (h || "").trim(); });
    var found = head.filter(function (h) { return h; });
    var missing = resolveHeaders(head);
    if (missing.length) return { books: [], missing: missing, found: found };
    var books = rows.slice(hi + 1).map(function (r) {
      var o = {};
      head.forEach(function (h, i) { o[h] = (r[i] || "").trim(); });
      return o;
    }).filter(function (o) { return o[COLS.title]; });   // 略過空白補位列
    return { books: books, missing: [], found: found };
  }

  // 依欄位名取值；選用欄位缺席時讀作空字串。
  function get(row, field) {
    var col = COLS[field];
    return col ? (row[col] || "") : "";
  }

  function counts(books) {
    var c = {};
    STATES.forEach(function (s) { c[s.label] = 0; });
    books.forEach(function (b) {
      var s = stateOf(get(b, "status"));
      c[s.label] = (c[s.label] || 0) + 1;
    });
    return c;
  }

  // 試算表公式出錯時，整格會變成 #VALUE! 之類的錯誤字；這與「欄位被改名」
  // 是兩回事，分開報告才說得出下一步該找誰修。
  var ERROR_CELL = /^#(VALUE!|REF!|N\/A|NAME\?|DIV\/0!|NUM!|NULL!|ERROR!)$/;

  function errorCell(rows) {
    for (var i = 0; i < rows.length && i < 5; i++) {
      for (var j = 0; j < rows[i].length; j++) {
        var c = (rows[i][j] || "").trim();
        if (ERROR_CELL.test(c)) return c;
      }
    }
    return null;
  }

  // 一律以 { books } 或 { error } 回覆，呼叫端不必各自判斷失敗原因。
  function load() {
    return fetch(CSV_URL, { cache: "no-store" })
      .then(function (r) {
        if (!r.ok) throw new Error("HTTP " + r.status);
        return r.text();
      })
      .then(function (text) {
        // 非 CSV 的發佈網址會回應網頁，若當成 CSV 解析會誤報為缺欄。
        if (/^\s*</.test(text) || /<html/i.test(text.slice(0, 400))) {
          return { error: "html" };
        }
        var rows = parseCSV(text);
        var out = toObjects(rows);
        if (out.missing.length) {
          var bad = errorCell(rows);
          if (bad) return { error: "formula", cell: bad };
          return { error: "missing", missing: out.missing, found: out.found };
        }
        if (!out.books.length) return { error: "empty" };
        return { books: out.books };
      })
      .catch(function () { return { error: "fetch" }; });
  }

  global.POOP = {
    STATES: STATES,
    stateOf: stateOf,
    readableOn: readableOn,
    get: get,
    counts: counts,
    load: load
  };
})(window);
