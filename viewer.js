/* Live workbook viewer: renders an exported Excel workbook (window.WB_DATA) and recalculates its formulas in the browser. */
(function () {
  "use strict";
  var D = window.WB_DATA; if (!D) return;
  var S = {}; D.sheets.forEach(function (s) { S[s.name] = s; });
  var ERR = function (e) { return { err: e }; };
  var isErr = function (v) { return v !== null && typeof v === "object" && v.err; };
  var isRng = function (v) { return v !== null && typeof v === "object" && v.rng; };

  /* ---------- A1 helpers ---------- */
  function colNum(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + s.charCodeAt(i) - 64; return n; }
  function colName(n) { var s = ""; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

  /* ---------- tokenizer / parser ---------- */
  var REF = /^((?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_\.]*)!)?\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?/;
  function tokenize(src) {
    var t = [], i = 0, m;
    while (i < src.length) {
      var ch = src[i];
      if (/\s/.test(ch)) { i++; continue; }
      var rest = src.slice(i);
      if (ch === '"') { var j = i + 1, s = ""; while (j < src.length) { if (src[j] === '"') { if (src[j + 1] === '"') { s += '"'; j += 2; continue; } break; } s += src[j++]; } t.push({ k: "str", v: s }); i = j + 1; continue; }
      if ((m = /^\d+(\.\d+)?([eE][+-]?\d+)?|^\.\d+/.exec(rest)) ) { t.push({ k: "num", v: parseFloat(m[0]) }); i += m[0].length; continue; }
      if ((m = /^[A-Za-z][A-Za-z0-9\.]*\(/.exec(rest))) { t.push({ k: "fn", v: m[0].slice(0, -1).toUpperCase() }); i += m[0].length; continue; }
      if ((m = /^(TRUE|FALSE)\b/i.exec(rest))) { t.push({ k: "bool", v: m[1].toUpperCase() === "TRUE" }); i += m[0].length; continue; }
      if ((m = REF.exec(rest))) {
        var sh = m[1] ? m[1].slice(0, -1) : null; if (sh && sh[0] === "'") sh = sh.slice(1, -1).replace(/''/g, "'");
        var c1 = colNum(m[2]), r1 = +m[3], c2 = m[4] ? colNum(m[4]) : c1, r2 = m[5] ? +m[5] : r1;
        t.push({ k: "ref", sheet: sh, r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) }); i += m[0].length; continue;
      }
      if ((m = /^(<>|<=|>=|[=<>+\-*\/^&%(),])/.exec(rest))) { t.push({ k: "op", v: m[0] }); i += m[0].length; continue; }
      throw new Error("bad token at " + rest.slice(0, 12));
    }
    return t;
  }
  function parse(src) {
    var t = tokenize(src), p = 0;
    function pk() { return t[p]; }
    function isOp(v) { return t[p] && t[p].k === "op" && t[p].v === v; }
    function cmp() { var l = cat(); while (t[p] && t[p].k === "op" && /^(=|<>|<=|>=|<|>)$/.test(t[p].v)) { var o = t[p++].v; l = { t: "bin", o: o, l: l, r: cat() }; } return l; }
    function cat() { var l = add(); while (isOp("&")) { p++; l = { t: "bin", o: "&", l: l, r: add() }; } return l; }
    function add() { var l = mul(); while (isOp("+") || isOp("-")) { var o = t[p++].v; l = { t: "bin", o: o, l: l, r: mul() }; } return l; }
    function mul() { var l = pow(); while (isOp("*") || isOp("/")) { var o = t[p++].v; l = { t: "bin", o: o, l: l, r: pow() }; } return l; }
    function pow() { var l = un(); while (isOp("^")) { p++; l = { t: "bin", o: "^", l: l, r: un() }; } return l; }
    function un() { if (isOp("-")) { p++; return { t: "un", e: un() }; } if (isOp("+")) { p++; return un(); } return post(); }
    function post() { var e = prim(); while (isOp("%")) { p++; e = { t: "bin", o: "/", l: e, r: { t: "num", v: 100 } }; } return e; }
    function prim() {
      var k = pk(); if (!k) throw new Error("unexpected end");
      if (k.k === "num") { p++; return { t: "num", v: k.v }; }
      if (k.k === "str") { p++; return { t: "str", v: k.v }; }
      if (k.k === "bool") { p++; return { t: "bool", v: k.v }; }
      if (k.k === "ref") { p++; return { t: "ref", sheet: k.sheet, r1: k.r1, c1: k.c1, r2: k.r2, c2: k.c2 }; }
      if (k.k === "fn") {
        p++; var args = [];
        if (isOp(")")) { p++; return { t: "fn", n: k.v, a: args }; }
        for (;;) {
          if (isOp(",")) { args.push({ t: "empty" }); p++; continue; }
          if (isOp(")")) { args.push({ t: "empty" }); p++; break; }
          args.push(cmp());
          if (isOp(",")) { p++; if (isOp(")")) { args.push({ t: "empty" }); p++; break; } continue; }
          if (isOp(")")) { p++; break; }
          throw new Error("expected , or )");
        }
        return { t: "fn", n: k.v, a: args };
      }
      if (isOp("(")) { p++; var e = cmp(); if (!isOp(")")) throw new Error("expected )"); p++; return e; }
      throw new Error("unexpected " + JSON.stringify(k));
    }
    var ast = cmp(); if (p < t.length) throw new Error("trailing tokens"); return ast;
  }

  /* ---------- evaluation ---------- */
  var cache = {}, busy = {};
  function rec(sheet, r, c) { var s = S[sheet]; return s ? s.cells[r + "," + c] : undefined; }
  function cellVal(sheet, r, c) {
    var rc = rec(sheet, r, c); if (!rc) return null;
    if (rc[1]) {
      var key = sheet + "!" + r + "," + c;
      if (key in cache) return cache[key];
      if (busy[key]) return ERR("#REF!");
      busy[key] = 1;
      var v;
      try { if (!rc.ast) rc.ast = parse(rc[1]); v = scalar(ev(rc.ast, sheet)); } catch (e) { v = ERR("#NAME?"); }
      delete busy[key]; cache[key] = v; return v;
    }
    return rc[0] === undefined ? null : rc[0];
  }
  function scalar(v) { if (isRng(v)) { return cellVal(v.sheet, v.r1, v.c1); } return v; }
  function toNum(v) {
    if (isErr(v)) return v; if (v === null || v === "") return 0; if (typeof v === "boolean") return v ? 1 : 0;
    if (typeof v === "number") return v; var n = Number(v); return isNaN(n) ? ERR("#VALUE!") : n;
  }
  function toBool(v) { if (isErr(v)) return v; if (typeof v === "boolean") return v; if (typeof v === "number") return v !== 0; if (v === null) return false; if (typeof v === "string") { var u = v.toUpperCase(); if (u === "TRUE") return true; if (u === "FALSE") return false; } return ERR("#VALUE!"); }
  function toStr(v) { if (isErr(v)) return v; if (v === null) return ""; if (typeof v === "boolean") return v ? "TRUE" : "FALSE"; return String(v); }
  function cmpVals(a, b) {
    if (a === null) a = (typeof b === "string") ? "" : (typeof b === "boolean" ? false : 0);
    if (b === null) b = (typeof a === "string") ? "" : (typeof a === "boolean" ? false : 0);
    var ra = typeof a === "number" ? 0 : typeof a === "string" ? 1 : 2, rb = typeof b === "number" ? 0 : typeof b === "string" ? 1 : 2;
    if (ra !== rb) return ra < rb ? -1 : 1;
    if (ra === 1) { a = a.toLowerCase(); b = b.toLowerCase(); }
    return a < b ? -1 : a > b ? 1 : 0;
  }
  function cells(rng, fn) { for (var r = rng.r1; r <= rng.r2; r++) for (var c = rng.c1; c <= rng.c2; c++) fn(cellVal(rng.sheet, r, c), r, c); }
  function critFn(crit) {
    crit = scalar(crit); var op = "=", val = crit;
    if (typeof crit === "string") { var m = /^(<>|<=|>=|=|<|>)(.*)$/.exec(crit); if (m) { op = m[1]; val = m[2]; } }
    var num = (typeof val === "string" && val !== "" && !isNaN(Number(val))) ? Number(val) : val;
    return function (v) {
      if (op === "=" || op === "<>") {
        var eq;
        if (num === "") eq = (v === null || v === "");
        else if (typeof num === "number") eq = (typeof v === "number" && v === num) || (typeof v === "string" && v !== "" && Number(v) === num);
        else if (typeof num === "string" && /[*?]/.test(num)) { var rx = new RegExp("^" + num.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".") + "$", "i"); eq = typeof v === "string" && rx.test(v); }
        else eq = (typeof v === "string" && typeof num === "string" && v.toLowerCase() === num.toLowerCase()) || (typeof v === "boolean" && v === num);
        if (op === "<>") { if (num === "") return !(v === null || v === ""); return !eq; }
        return eq;
      }
      if (v === null || typeof v === "boolean" || isErr(v)) return false;
      if (typeof num === "number") { if (typeof v !== "number") return false; }
      else if (typeof v !== "string") return false;
      var cc = cmpVals(v, num);
      return op === ">" ? cc > 0 : op === "<" ? cc < 0 : op === ">=" ? cc >= 0 : cc <= 0;
    };
  }
  function flat(args, ctx, onlyNums) {
    var out = [], err = null;
    args.forEach(function (a) {
      var v = ev(a, ctx);
      if (isRng(v)) cells(v, function (x) { if (isErr(x)) { err = err || x; return; } if (onlyNums) { if (typeof x === "number") out.push(x); } else out.push(x); });
      else { if (isErr(v)) err = err || v; else if (onlyNums) { if (v === null) return; var n = toNum(v); if (isErr(n)) err = err || n; else out.push(n); } else out.push(v); }
    });
    return err ? { err: err.err } : out;
  }
  var FN = {
    IF: function (a, ctx) { var c = toBool(scalar(ev(a[0], ctx))); if (isErr(c)) return c; var br = c ? a[1] : a[2]; if (!br || br.t === "empty") return c ? 0 : false; var r = scalar(ev(br, ctx)); return r === null ? 0 : r; },
    IFERROR: function (a, ctx) { var v = scalar(ev(a[0], ctx)); return isErr(v) ? scalar(ev(a[1], ctx)) : v; },
    ISNUMBER: function (a, ctx) { var v = scalar(ev(a[0], ctx)); return typeof v === "number"; },
    NOT: function (a, ctx) { var b = toBool(scalar(ev(a[0], ctx))); return isErr(b) ? b : !b; },
    AND: function (a, ctx) { var f = flat(a, ctx, false); if (isErr(f)) return f; var res = true, any = false; f.forEach(function (x) { if (typeof x === "boolean" || typeof x === "number") { any = true; if (!toBool(x)) res = false; } }); return any ? res : ERR("#VALUE!"); },
    OR: function (a, ctx) { var f = flat(a, ctx, false); if (isErr(f)) return f; var res = false, any = false; f.forEach(function (x) { if (typeof x === "boolean" || typeof x === "number") { any = true; if (toBool(x)) res = true; } }); return any ? res : ERR("#VALUE!"); },
    SUM: function (a, ctx) { var f = flat(a, ctx, true); if (isErr(f)) return f; return f.reduce(function (s, x) { return s + x; }, 0); },
    COUNT: function (a, ctx) { var n = 0; a.forEach(function (x) { var v = ev(x, ctx); if (isRng(v)) cells(v, function (q) { if (typeof q === "number") n++; }); else if (typeof v === "number") n++; }); return n; },
    COUNTA: function (a, ctx) { var n = 0; a.forEach(function (x) { var v = ev(x, ctx); if (isRng(v)) cells(v, function (q) { if (q !== null) n++; }); else if (v !== null) n++; }); return n; },
    AVERAGE: function (a, ctx) { var f = flat(a, ctx, true); if (isErr(f)) return f; return f.length ? f.reduce(function (s, x) { return s + x; }, 0) / f.length : ERR("#DIV/0!"); },
    MAX: function (a, ctx) { var f = flat(a, ctx, true); return isErr(f) ? f : (f.length ? Math.max.apply(null, f) : 0); },
    MIN: function (a, ctx) { var f = flat(a, ctx, true); return isErr(f) ? f : (f.length ? Math.min.apply(null, f) : 0); },
    ABS: function (a, ctx) { var n = toNum(scalar(ev(a[0], ctx))); return isErr(n) ? n : Math.abs(n); },
    ROUND: function (a, ctx) { var n = toNum(scalar(ev(a[0], ctx))), d = a[1] ? toNum(scalar(ev(a[1], ctx))) : 0; if (isErr(n)) return n; if (isErr(d)) return d; var f = Math.pow(10, d); return (n < 0 ? -1 : 1) * Math.round(Math.abs(n) * f + 1e-9) / f; },
    COUNTIF: function (a, ctx) { var r = ev(a[0], ctx), fn = critFn(ev(a[1], ctx)), n = 0; if (!isRng(r)) return fn(r) ? 1 : 0; cells(r, function (v) { if (fn(v)) n++; }); return n; },
    COUNTIFS: function (a, ctx) { return ifs(null, a, ctx); },
    SUMIFS: function (a, ctx) { return ifs(ev(a[0], ctx), a.slice(1), ctx); },
    INDEX: function (a, ctx) {
      var r = ev(a[0], ctx); if (!isRng(r)) return r;
      var rows = r.r2 - r.r1 + 1, cols = r.c2 - r.c1 + 1, ri = a[1] ? toNum(scalar(ev(a[1], ctx))) : 1, ci = a[2] ? toNum(scalar(ev(a[2], ctx))) : null;
      if (isErr(ri)) return ri; if (ci !== null && isErr(ci)) return ci;
      if (ci === null) { if (rows === 1 && cols > 1) { ci = ri; ri = 1; } else ci = 1; }
      if (ri < 1 || ci < 1 || ri > rows || ci > cols) return ERR("#REF!");
      return cellVal(r.sheet, r.r1 + ri - 1, r.c1 + ci - 1);
    },
    MATCH: function (a, ctx) {
      var v = scalar(ev(a[0], ctx)), r = ev(a[1], ctx), t = a[2] ? toNum(scalar(ev(a[2], ctx))) : 1; if (isErr(v)) return v;
      if (!isRng(r)) return ERR("#N/A");
      var list = []; cells(r, function (x) { list.push(x); });
      var i, best = -1;
      if (t === 0) { var fn = critFn(typeof v === "string" ? "=" + v : v); for (i = 0; i < list.length; i++) { if (list[i] !== null && fn(list[i])) return i + 1; } return ERR("#N/A"); }
      for (i = 0; i < list.length; i++) { if (list[i] === null) continue; var c = cmpVals(list[i], v); if (t > 0 ? c <= 0 : c >= 0) best = i; else if (t > 0) break; }
      return best < 0 ? ERR("#N/A") : best + 1;
    }
  };
  function ifs(sumR, a, ctx) {
    var pairs = [], i; for (i = 0; i + 1 < a.length; i += 2) { var rg = ev(a[i], ctx); if (!isRng(rg)) return ERR("#VALUE!"); pairs.push([rg, critFn(ev(a[i + 1], ctx))]); }
    var base = pairs[0][0], rows = base.r2 - base.r1 + 1, cols = base.c2 - base.c1 + 1, tot = 0;
    for (var r = 0; r < rows; r++) for (var c = 0; c < cols; c++) {
      var ok = true; for (i = 0; i < pairs.length && ok; i++) { var pr = pairs[i][0]; ok = pairs[i][1](cellVal(pr.sheet, pr.r1 + r, pr.c1 + c)); }
      if (!ok) continue;
      if (sumR) { var sv = cellVal(sumR.sheet, sumR.r1 + r, sumR.c1 + c); if (typeof sv === "number") tot += sv; } else tot++;
    }
    return tot;
  }
  function ev(n, ctx, off) {
    switch (n.t) {
      case "num": case "str": case "bool": return n.v;
      case "empty": return null;
      case "ref": var dr = off ? off[0] : 0, dc = off ? off[1] : 0; return { rng: 1, sheet: n.sheet || ctx, r1: n.r1 + dr, c1: n.c1 + dc, r2: n.r2 + dr, c2: n.c2 + dc };
      case "un": var x = toNum(scalar(ev(n.e, ctx, off))); return isErr(x) ? x : -x;
      case "bin":
        var l = scalar(ev(n.l, ctx, off)), r = scalar(ev(n.r, ctx, off));
        if (isErr(l)) return l; if (isErr(r)) return r;
        if (n.o === "&") return toStr(l) + toStr(r);
        if (/^(=|<>|<|>|<=|>=)$/.test(n.o)) { var c = cmpVals(l, r); return n.o === "=" ? c === 0 : n.o === "<>" ? c !== 0 : n.o === "<" ? c < 0 : n.o === ">" ? c > 0 : n.o === "<=" ? c <= 0 : c >= 0; }
        var a = toNum(l), b = toNum(r); if (isErr(a)) return a; if (isErr(b)) return b;
        if (n.o === "+") return a + b; if (n.o === "-") return a - b; if (n.o === "*") return a * b;
        if (n.o === "/") return b === 0 ? ERR("#DIV/0!") : a / b; return Math.pow(a, b);
      case "fn":
        var f = FN[n.n]; if (!f) return ERR("#NAME?");
        if (off) { var shifted = { t: "fn", n: n.n, a: n.a.map(function (q) { return { t: "off", q: q, off: off }; }) }; return f(shifted.a, ctx); }
        return f(n.a, ctx);
      case "off": return ev(n.q, ctx, n.off);
    }
    return ERR("#VALUE!");
  }

  /* ---------- number formats ---------- */
  function fmtNum(v, nf) {
    if (!nf) { if (Number.isInteger(v)) return String(v); var s = parseFloat(v.toPrecision(10)).toString(); return s; }
    if (/y/i.test(nf) && /d/i.test(nf)) {
      var d = new Date(Math.round((v - 25569) * 86400000)); if (isNaN(d)) return String(v);
      var p = function (x) { return (x < 10 ? "0" : "") + x; }, mo = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
      return nf.replace(/\\/g, "").replace(/yyyy/g, d.getUTCFullYear()).replace(/yy/g, String(d.getUTCFullYear()).slice(2)).replace(/mmm/g, mo[d.getUTCMonth()]).replace(/mm/g, p(d.getUTCMonth() + 1)).replace(/dd/g, p(d.getUTCDate()));
    }
    var pct = /%/.test(nf), val = pct ? v * 100 : v, m = /0\.(0+)/.exec(nf), dec = m ? m[1].length : 0, th = /#,##/.test(nf);
    var neg = val < 0, s2 = Math.abs(val).toFixed(dec); if (th) { var pp = s2.split("."); pp[0] = pp[0].replace(/\B(?=(\d{3})+(?!\d))/g, ","); s2 = pp.join("."); }
    var pre = (/^"([^"]*)"/.exec(nf) || [, ""])[1];
    return (neg ? "-" : "") + pre + s2 + (pct ? "%" : "");
  }
  function show(v, nf) {
    if (v === null || v === undefined) return "";
    if (isErr(v)) return v.err;
    if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
    if (typeof v === "number") return fmtNum(v, nf);
    return String(v);
  }

  /* ---------- conditional formatting ---------- */
  function cfStyle(sheet, r, c) {
    var s = S[sheet], out = null;
    for (var i = 0; i < s.cf.length; i++) {
      var k = s.cf[i]; if (r < k.r1 || r > k.r2 || c < k.c1 || c > k.c2) continue;
      var v = cellVal(sheet, r, c), hit = false;
      try {
        if (k.t === "cellIs") {
          var f = k.f.map(function (x) { if (!k.ast) k.ast = {}; if (!k.ast[x]) k.ast[x] = parse(x); return scalar(ev(k.ast[x], sheet)); });
          if (v === null || isErr(v)) hit = false; else {
            var op = k.op, cc = function (x) { if (typeof x !== typeof v && !(typeof x === "number" && typeof v === "number")) return null; return cmpVals(v, x); };
            if (op === "equal") hit = cc(f[0]) === 0; else if (op === "notEqual") hit = cc(f[0]) !== 0;
            else if (op === "lessThan") hit = cc(f[0]) !== null && cc(f[0]) < 0; else if (op === "lessThanOrEqual") hit = cc(f[0]) !== null && cc(f[0]) <= 0;
            else if (op === "greaterThan") hit = cc(f[0]) !== null && cc(f[0]) > 0; else if (op === "greaterThanOrEqual") hit = cc(f[0]) !== null && cc(f[0]) >= 0;
            else if (op === "between") hit = cc(f[0]) !== null && cc(f[0]) >= 0 && cc(f[1]) <= 0; else if (op === "notBetween") hit = !(cc(f[0]) >= 0 && cc(f[1]) <= 0);
          }
        } else if (k.t === "expression") {
          if (!k.ast) k.ast = {}; var x0 = k.f[0]; if (!k.ast[x0]) k.ast[x0] = parse(x0);
          var res = scalar(ev(k.ast[x0], sheet, [r - k.r1, c - k.c1])); hit = res === true;
        }
      } catch (e) { hit = false; }
      if (hit) { out = out || {}; if (k.bg && !out.bg) out.bg = k.bg; if (k.fg && !out.fg) out.fg = k.fg; }
    }
    return out;
  }

  /* ---------- rendering ---------- */
  var root = document.getElementById("wb"), cur = 0, bar, addr, fx, tabsEl, scroller, editAll = false;
  var DEFH = 22;
  function dvFor(s, r, c) { for (var i = 0; i < s.dv.length; i++) { var d = s.dv[i]; if (r >= d.r1 && r <= d.r2 && c >= d.c1 && c <= d.c2) return d; } return null; }
  function dvList(s, d) {
    if (d.list) return d.list;
    var m = /^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/.exec((d.ref || "").replace(/^.*!/, "")); if (!m) return [];
    var out = []; for (var r = +m[2]; r <= +m[4]; r++) for (var c = colNum(m[1]); c <= colNum(m[3]); c++) { var v = cellVal(s.name, r, c); if (v !== null) out.push(String(v)); } return out;
  }
  function render() {
    var s = D.sheets[cur]; cache = {};
    var merged = {}, covered = {};
    s.merges.forEach(function (m) { var mm = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(m); if (!mm) return; var c1 = colNum(mm[1]), r1 = +mm[2], c2 = colNum(mm[3]), r2 = +mm[4]; merged[r1 + "," + c1] = [r2 - r1 + 1, c2 - c1 + 1]; for (var r = r1; r <= r2; r++) for (var c = c1; c <= c2; c++) if (r !== r1 || c !== c1) covered[r + "," + c] = 1; });
    var fr = s.freeze ? s.freeze[0] : 0, fc = s.freeze ? s.freeze[1] : 0;
    var colLeft = [0], x = 40; for (var c = 1; c <= s.cols; c++) { colLeft[c] = x; x += s.widths[c - 1]; }
    var rowTop = [0], y = 24; for (var r = 1; r <= s.rows; r++) { rowTop[r] = y; y += (s.heights[r] || DEFH); }
    var h = '<table class="grid"><colgroup><col style="width:40px">';
    for (c = 1; c <= s.cols; c++) h += '<col style="width:' + s.widths[c - 1] + 'px' + (s.widths[c - 1] === 0 ? ";visibility:collapse" : "") + '">';
    h += "</colgroup><thead><tr><th class=\"corner\"></th>";
    for (c = 1; c <= s.cols; c++) h += '<th class="ch' + (c <= fc ? " fz" : "") + '"' + (c <= fc ? ' style="left:' + colLeft[c] + 'px"' : "") + ">" + colName(c) + "</th>";
    h += "</tr></thead><tbody>";
    for (r = 1; r <= s.rows; r++) {
      var rh = s.heights[r] || DEFH;
      h += '<tr style="height:' + rh + 'px"><th class="rh' + (r <= fr ? " fz" : "") + '"' + (r <= fr ? ' style="top:' + rowTop[r] + 'px"' : "") + ">" + r + "</th>";
      for (c = 1; c <= s.cols; c++) {
        var key = r + "," + c; if (covered[key]) continue;
        var rc = s.cells[key], st = rc ? D.styles[rc[2]] || {} : {}, span = merged[key];
        var cls = "c" + (rc && rc[4] ? " ed" : "") + (rc && rc[1] ? " fm" : "") + (r <= fr ? " fzr" : "") + (c <= fc ? " fzc" : "");
        var css = "";
        if (st.bg) css += "background:" + st.bg + ";"; if (st.c) css += "color:" + st.c + ";"; if (st.b) css += "font-weight:700;"; if (st.i) css += "font-style:italic;"; if (st.sz) css += "font-size:" + Math.round(st.sz * 1.2) + "px;";
        css += "text-align:" + (st.h === "center" ? "center" : st.h === "right" ? "right" : st.h === "left" ? "left" : "") + ";";
        if (st.v === "center") css += "vertical-align:middle;"; else if (st.v === "top") css += "vertical-align:top;";
        css += st.w ? "white-space:normal;" : "white-space:nowrap;overflow:hidden;text-overflow:ellipsis;";
        if (st.bd) cls += " bd";
        if (r <= fr) css += "top:" + rowTop[r] + "px;"; if (c <= fc) css += "left:" + colLeft[c] + "px;";
        h += '<td class="' + cls + '" data-r="' + r + '" data-c="' + c + '"' + (span ? ' rowspan="' + span[0] + '" colspan="' + span[1] + '"' : "") + ' style="' + css + '" data-bg="' + (st.bg || "") + '" data-fg="' + (st.c || "") + '"></td>';
      }
      h += "</tr>";
    }
    h += "</tbody></table>";
    scroller.innerHTML = h;
    paint();
  }
  function paint() {
    var s = D.sheets[cur]; cache = {};
    var tds = scroller.querySelectorAll("td.c");
    for (var i = 0; i < tds.length; i++) {
      var td = tds[i], r = +td.getAttribute("data-r"), c = +td.getAttribute("data-c"), rc = s.cells[r + "," + c];
      if (td.firstChild && td.firstChild.nodeType === 1) continue; /* being edited */
      var v = cellVal(s.name, r, c), nf = rc ? rc[3] : 0;
      td.textContent = show(v, nf);
      if (typeof v === "number" && !(rc && D.styles[rc[2]] && D.styles[rc[2]].h)) td.style.textAlign = "right";
      var k = cfStyle(s.name, r, c);
      td.style.background = k && k.bg ? k.bg : td.getAttribute("data-bg") || "";
      td.style.color = k && k.fg ? k.fg : td.getAttribute("data-fg") || "";
      if (isErr(v)) td.classList.add("er"); else td.classList.remove("er");
    }
  }
  function setVal(sheetName, r, c, val) {
    var s = S[sheetName], key = r + "," + c, rc = s.cells[key];
    if (val === "" || val === null) val = null; else if (typeof val === "string" && val.trim() !== "" && !isNaN(Number(val))) val = Number(val);
    if (!rc) { rc = s.cells[key] = [val, null, 0, 0, 1]; } else rc[0] = val;
  }
  function selectCell(td) {
    var prev = scroller.querySelector("td.sel"); if (prev) prev.classList.remove("sel");
    td.classList.add("sel");
    var s = D.sheets[cur], r = +td.getAttribute("data-r"), c = +td.getAttribute("data-c"), rc = s.cells[r + "," + c];
    addr.textContent = colName(c) + r; fx.textContent = rc && rc[1] ? "=" + rc[1] : (rc && rc[0] !== null && rc[0] !== undefined ? String(rc[0]) : "");
  }
  function beginEdit(td) {
    if (td.querySelector("input,select")) return;
    var s = D.sheets[cur], r = +td.getAttribute("data-r"), c = +td.getAttribute("data-c"), rc = s.cells[r + "," + c], d = dvFor(s, r, c), cv = rc && rc[0] !== null && rc[0] !== undefined ? String(rc[0]) : "", el;
    if (d) {
      var list = dvList(s, d); el = document.createElement("select");
      var o = document.createElement("option"); o.value = ""; o.textContent = ""; el.appendChild(o);
      list.forEach(function (x) { var op = document.createElement("option"); op.value = x; op.textContent = x; if (x === cv) op.selected = true; el.appendChild(op); });
      el.addEventListener("change", function () { commit(td, r, c, el.value); });
    } else { el = document.createElement("input"); el.type = "text"; el.value = cv; el.addEventListener("keydown", function (e) { if (e.key === "Enter") { commit(td, r, c, el.value); } else if (e.key === "Escape") { el._done = true; td.removeChild(el); paint(); } }); }
    el.addEventListener("blur", function () { if (!el._done && td.contains(el)) commit(td, r, c, el.value); });
    td.textContent = ""; td.appendChild(el); el.focus();
  }
  function commit(td, r, c, val) {
    var el = td.firstChild; if (!el || el.nodeType !== 1 || el._done) return;
    el._done = true; if (el.parentNode === td) td.removeChild(el);
    setVal(D.sheets[cur].name, r, c, val); paintAll();
  }
  function paintAll() { cache = {}; paint(); }
  function hasRendered() { return !!scroller.querySelector("table"); }

  /* ---------- sample data ---------- */
  function rng32(a) { return function () { a |= 0; a = a + 0x6D2B79F5 | 0; var t = Math.imul(a ^ a >>> 15, 1 | a); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; }; }
  function loadExample() {
    var R = rng32(7);
    D.sheets.forEach(function (s) {
      if (D.sample === "ohs") {
        for (var r = 1; r <= s.rows; r++) { var a = s.cells[r + ",1"], b = s.cells[r + ",2"]; if (a && a[0] && b && b[4]) { var p = R(); setVal(s.name, r, p < .82 ? 2 : p < .93 ? 3 : 4, "X"); } }
        return;
      }
      s.dv.forEach(function (d) {
        var list = dvList(s, d); if (!list.length || list.length > 6) return;
        for (var r = d.r1; r <= d.r2; r++) for (var c = d.c1; c <= d.c2; c++) {
          var rc = s.cells[r + "," + c]; if (!rc || !rc[4]) continue;
          /* only rate rows that have an item text in column B or A */
          var lab = s.cells[r + ",2"] || s.cells[r + ",3"] || s.cells[r + ",1"]; if (!lab || !lab[0] && !lab[1]) continue;
          if (R() < .12) continue;
          var pool = list.filter(function (x) { return x !== "N/A"; }), pick = R() < .06 && list.indexOf("N/A") > -1 ? "N/A" : pool[Math.floor(Math.pow(R(), .75) * pool.length)];
          setVal(s.name, r, c, pick);
        }
      });
    });
    paintAll();
  }

  /* ---------- UI shell ---------- */
  function build() {
    var tabs = D.sheets.map(function (s, i) { return '<button class="tab' + (i === 0 ? " on" : "") + '" data-i="' + i + '">' + s.name + "</button>"; }).join("");
    root.innerHTML = '<div class="bar"><div class="tabs">' + tabs + '</div><div class="acts">' +
      (D.sample ? '<button class="b" id="ex">Load example data</button>' : "") +
      '<button class="b" id="ea">Edit any cell: off</button><button class="b" id="rs">Reset</button><a class="b dl" href="' + D.file + '">&darr; Download Excel</a></div></div>' +
      '<div class="fxbar"><span id="addr">A1</span><span id="fx"></span></div><div class="scroller" id="sc"></div>';
    scroller = document.getElementById("sc"); addr = document.getElementById("addr"); fx = document.getElementById("fx");
    root.querySelectorAll(".tab").forEach(function (b) { b.addEventListener("click", function () { cur = +b.getAttribute("data-i"); root.querySelectorAll(".tab").forEach(function (x) { x.classList.toggle("on", x === b); }); render(); scroller.scrollTop = 0; scroller.scrollLeft = 0; }); });
    var ex = document.getElementById("ex"); if (ex) ex.addEventListener("click", loadExample);
    document.getElementById("rs").addEventListener("click", function () { location.reload(); });
    var ea = document.getElementById("ea"); ea.addEventListener("click", function () { editAll = !editAll; ea.textContent = "Edit any cell: " + (editAll ? "on" : "off"); ea.classList.toggle("onb", editAll); });
    scroller.addEventListener("click", function (e) {
      var td = e.target.closest ? e.target.closest("td.c") : null; if (!td) return;
      selectCell(td);
      var s = D.sheets[cur], r = +td.getAttribute("data-r"), c = +td.getAttribute("data-c"), rc = s.cells[r + "," + c];
      if ((rc && rc[4]) || (editAll && !(rc && rc[1]))) beginEdit(td);
    });
    render();
  }
  window.WBV = { evalAll: function () { cache = {}; var out = {}; D.sheets.forEach(function (s) { out[s.name] = {}; Object.keys(s.cells).forEach(function (k) { var rc = s.cells[k]; if (rc[1]) { var rr = k.split(","); var v = cellVal(s.name, +rr[0], +rr[1]); out[s.name][colName(+rr[1]) + rr[0]] = v; } }); }); return out; }, set: function (sh, a1, v) { var m = /^([A-Z]+)(\d+)$/.exec(a1); setVal(sh, +m[2], colNum(m[1]), v); cache = {}; }, example: loadExample };
  if (root) build();
})();
