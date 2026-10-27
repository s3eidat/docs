/* Docs: the document model. Pure functions (no DOM, no storage) so the self-test can prove them.
 *
 * Doc    {v, id, title, created, updated, page, hf, base, blocks[]}
 * page   {size, orient, w, h (mm, as laid out), mt, mr, mb, ml, hd, fd (mm)}
 * hf     {header, footer (text with {page} {pages} {title} {date} tokens), ha, fa (align), first (different first page)}
 * base   {font, size (pt), line (multiple), after (pt)}   = the Normal style
 * Blocks p     {id, type:"p", style, align, dir, ind, fi, lh, sb, sa, list:{k:"ul"|"ol"|"todo", l, c}|null, m, text:spans}
 *        table {id, type:"table", rows:[[cell]], cw:[fractions], hdr}   cell {text, bg, al, cs, rs, m (covered by a merge)}
 *        img   {id, type:"img", blob, w, h (CSS px at 96 dpi), align, alt}
 *        pb    {id, type:"pb"}   page break
 *        toc   {id, type:"toc"}  table of contents, built from headings at layout time
 * Spans  [{t, b, i, u, s, sup, sub, f font, z size pt, c colour #rrggbb, h highlight #rrggbb, a url}]
 * A position is {b: block index, r, c (table cell or -1), o: offset in that text}. Nothing here makes HTML.
 */
(function () {
  "use strict";

  var ID_RE = /^[a-z0-9]{1,24}$/;
  var BOOL = ["b", "i", "u", "s", "sup", "sub"];
  var KEYS = BOOL.concat(["f", "z", "c", "h", "a"]);
  var HEX = /^#[0-9a-f]{6}$/i;
  var FONT_RE = /^[\p{L}\p{N} ._\-]{1,48}$/u;
  var MM = 96 / 25.4, PT = 96 / 72;
  // Named highlight colours (the standard set word processors share), so highlights survive a .docx round trip.
  var HL = { yellow: "#ffff00", green: "#00ff00", cyan: "#00ffff", magenta: "#ff00ff", blue: "#0000ff", red: "#ff0000",
    darkBlue: "#000080", darkCyan: "#008080", darkGreen: "#008000", darkMagenta: "#800080", darkRed: "#800000",
    darkYellow: "#808000", darkGray: "#808080", lightGray: "#c0c0c0", black: "#000000", white: "#ffffff" };
  var SIZES = { a4: [210, 297], letter: [215.9, 279.4], legal: [215.9, 355.6], a5: [148, 210], b5: [176, 250] };
  // Styles. size null = base size. sb/sa in pt, lh null = base line spacing.
  var STYLES = {
    normal: { name: "Normal" },
    title: { name: "Title", size: 26, sb: 0, sa: 6, lh: 1.05, color: "#17233a", font: "" },
    subtitle: { name: "Subtitle", size: 14, sb: 0, sa: 12, color: "#5a6372" },
    h1: { name: "Heading 1", size: 18, bold: true, sb: 18, sa: 6, lh: 1.1, color: "#1d3557", next: true },
    h2: { name: "Heading 2", size: 14, bold: true, sb: 14, sa: 4, lh: 1.1, color: "#1d3557", next: true },
    h3: { name: "Heading 3", size: 12, bold: true, sb: 12, sa: 3, lh: 1.1, color: "#34495e", next: true },
    quote: { name: "Quote", italic: true, color: "#4b5563", ind: 1, sb: 6, sa: 10, bar: true },
    code: { name: "Code", size: 10, font: "Consolas", mono: true, lh: 1.15, sa: 0, sb: 0, bg: "#f3f4f6" }
  };
  var STYLE_KEYS = Object.keys(STYLES);
  var HEADINGS = { title: 0, h1: 1, h2: 2, h3: 3 };
  var MAX_LEVEL = 8;

  function uid() {
    var a = new Uint8Array(9), s = "";
    (window.crypto || window.msCrypto).getRandomValues(a);
    for (var i = 0; i < a.length; i++) s += (a[i] % 36).toString(36);
    return s + Date.now().toString(36).slice(-3);
  }
  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); }
  function str(v, max) { return typeof v === "string" ? v.slice(0, max) : ""; }
  function num(v, lo, hi, def) { v = typeof v === "string" && v.trim() !== "" ? Number(v) : v; return typeof v === "number" && isFinite(v) && v >= lo && v <= hi ? v : def; }
  function clamp(v, lo, hi) { return Math.max(lo, Math.min(hi, v)); }

  // ---------- links: http(s) and mailto only. Bare domains get https, bare emails get mailto. ----------
  function safeUrl(raw) {
    if (typeof raw !== "string") return null;
    var u = raw.replace(/[\u0000-\u001F\u007F\s]+/g, function (m) { return /^ +$/.test(m) ? m : ""; }).trim();
    if (!u || u.length > 2048) return null;
    var m = /^([a-z][a-z0-9+.\-]*):/i.exec(u);
    if (m) {
      var sc = m[1].toLowerCase();
      if (sc === "http" || sc === "https") return /^https?:\/\/[^\s\/?#]+/i.test(u) ? u.replace(/ /g, "%20") : null;
      if (sc === "mailto") return /^mailto:[^\s]+@[^\s]+$/i.test(u.replace(/ /g, "%20")) ? u.replace(/ /g, "%20") : null;
      return null;
    }
    if (/^[\w.+-]+@[\w-]+\.[\w.-]+$/.test(u)) return "mailto:" + u;
    if (/^(www\.)?[\p{L}\p{N}-]+(\.[\p{L}\p{N}-]+)+(:\d+)?([\/?#]\S*)?$/u.test(u)) return "https://" + u;
    return null;
  }

  // ---------- spans ----------
  function same(a, b) { for (var k = 0; k < KEYS.length; k++) if ((a[KEYS[k]] || "") !== (b[KEYS[k]] || "")) return false; return true; }
  function marksOf(s) { var o = {}; KEYS.forEach(function (k) { if (s && s[k]) o[k] = s[k]; }); return o; }
  function withText(m, t) { var o = marksOf(m); o.t = t; return o; }
  function norm(spans) {
    var out = [];
    (spans || []).forEach(function (s) {
      if (!s || !s.t) return;
      var last = out[out.length - 1];
      if (last && same(last, s)) last.t += s.t; else out.push(withText(s, s.t));
    });
    return out;
  }
  function plain(spans) { var s = ""; (spans || []).forEach(function (x) { s += x.t; }); return s; }
  function len(spans) { var n = 0; (spans || []).forEach(function (x) { n += x.t.length; }); return n; }
  function slice(spans, a, b) {
    var out = [], pos = 0;
    if (b == null) b = Infinity;
    (spans || []).forEach(function (s) {
      var L = s.t.length, s0 = pos, s1 = pos + L;
      pos = s1;
      if (s1 <= a || s0 >= b) return;
      out.push(withText(s, s.t.slice(Math.max(0, a - s0), Math.min(L, b - s0))));
    });
    return norm(out);
  }
  function concat(x, y) { return norm((x || []).concat(y || [])); }
  function cut(spans, a, b) { return concat(slice(spans, 0, a), slice(spans, b)); }
  function insertSpans(spans, at, ins) { return norm(slice(spans, 0, at).concat(ins || [], slice(spans, at))); }
  // Marks a character typed at `at` gets: the character before it (or after it at the very start). A link does not
  // grow when you type at its end.
  function marksAt(spans, at) {
    var L = len(spans);
    if (!L) return {};
    var before = at > 0 ? slice(spans, at - 1, at)[0] : null, after = at < L ? slice(spans, at, at + 1)[0] : null;
    var m = marksOf(before || after);
    if (m.a && (!before || !after || before.a !== after.a)) delete m.a;
    return m;
  }
  function insertText(spans, at, text, marks) { return insertSpans(spans, at, [withText(marks || marksAt(spans, at), text)]); }
  function valOk(k, v) {
    if (BOOL.indexOf(k) >= 0) return v === true;
    if (k === "f") return typeof v === "string" && FONT_RE.test(v);
    if (k === "z") return typeof v === "number" && v >= 1 && v <= 400;
    if (k === "c" || k === "h") return typeof v === "string" && HEX.test(v);
    if (k === "a") return typeof v === "string" && !!safeUrl(v);
    return false;
  }
  function setMark(spans, a, b, key, value) {
    if (a >= b) return norm(spans);
    var mid = slice(spans, a, b).map(function (s) {
      var o = withText(s, s.t);
      if (value != null && value !== false && value !== "") o[key] = value; else delete o[key];
      if (key === "sup" && value) delete o.sub;
      if (key === "sub" && value) delete o.sup;
      return o;
    });
    return norm(slice(spans, 0, a).concat(mid, slice(spans, b)));
  }
  function hasMark(spans, a, b, key) {
    var part = slice(spans, a, b);
    return part.length > 0 && part.every(function (s) { return !!s[key]; });
  }
  function cleanSpans(arr, budget) {
    var max = budget || 200000;
    if (!Array.isArray(arr)) return typeof arr === "string" ? norm([{ t: arr.slice(0, max) }]) : [];
    var out = [], total = 0;
    for (var k = 0; k < arr.length && k < 20000 && total < max; k++) {
      var s = arr[k];
      if (!s || typeof s !== "object") continue;
      var t = str(s.t, max - total).replace(/\r\n?/g, "\n").replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
      if (!t) continue;
      var o = { t: t };
      KEYS.forEach(function (key) {
        var v = s[key];
        if (key === "z") v = num(v, 1, 400, null);
        if (key === "a" && typeof v === "string") v = safeUrl(v);
        if (key === "c" || key === "h") v = typeof v === "string" ? v.toLowerCase() : v;
        if (v != null && valOk(key, v)) o[key] = v;
      });
      if (o.sup && o.sub) delete o.sub;
      out.push(o); total += t.length;
    }
    return norm(out);
  }
  function spansOf(text, marks) { return text ? norm([withText(marks || {}, String(text))]) : []; }

  // ---------- blocks ----------
  function para(style, text, props) {
    var b = { id: uid(), type: "p", style: STYLES[style] ? style : "normal", align: "", dir: "", ind: 0, fi: "", lh: 0, sb: null, sa: null, list: null, text: [] };
    if (text != null) b.text = Array.isArray(text) ? norm(text) : spansOf(text);
    if (props) for (var k in props) b[k] = props[k];
    return b;
  }
  function cell(text) { return { text: text ? (Array.isArray(text) ? norm(text) : spansOf(text)) : [], bg: "", al: "", cs: 1, rs: 1 }; }
  function table(rows, cols, data) {
    var t = { id: uid(), type: "table", rows: [], cw: [], hdr: true };
    for (var r = 0; r < rows; r++) { var row = []; for (var c = 0; c < cols; c++) row.push(cell(data && data[r] ? data[r][c] : "")); t.rows.push(row); }
    for (var k = 0; k < cols; k++) t.cw.push(1 / cols);
    return t;
  }
  function image(blob, w, h, props) { var b = { id: uid(), type: "img", blob: blob, w: w, h: h, align: "center", alt: "" }; if (props) for (var k in props) b[k] = props[k]; return b; }
  function pagebreak() { return { id: uid(), type: "pb" }; }
  function toc() { return { id: uid(), type: "toc" }; }
  function isText(b) { return b && b.type === "p"; }

  function cleanList(l) {
    if (!l || typeof l !== "object") return null;
    var k = l.k === "ul" || l.k === "ol" || l.k === "todo" ? l.k : null;
    if (!k) return null;
    var o = { k: k, l: Math.floor(num(l.l, 0, MAX_LEVEL, 0)) };
    if (k === "todo") o.c = l.c === true;
    return o;
  }
  function cleanBlock(x, budget) {
    x = x && typeof x === "object" ? x : {};
    var id = typeof x.id === "string" && ID_RE.test(x.id) ? x.id : uid();
    if (x.type === "table") {
      var rows = Array.isArray(x.rows) ? x.rows.slice(0, 1000) : [];
      var cols = Math.min(63, Math.max.apply(null, [1].concat(rows.map(function (r) { return Array.isArray(r) ? r.length : 0; }))));
      var t = { id: id, type: "table", rows: [], cw: [], hdr: x.hdr === true };
      rows.forEach(function (r) {
        var row = [];
        for (var c = 0; c < cols; c++) {
          var s = Array.isArray(r) && r[c] && typeof r[c] === "object" ? r[c] : {};
          var cl = cell(cleanSpans(s.text, 20000));
          if (typeof s.bg === "string" && HEX.test(s.bg)) cl.bg = s.bg.toLowerCase();
          if (s.al === "left" || s.al === "center" || s.al === "right" || s.al === "justify") cl.al = s.al;
          cl.cs = Math.floor(num(s.cs, 1, 63, 1)); cl.rs = Math.floor(num(s.rs, 1, 1000, 1));
          if (s.m === true) cl.m = true;
          row.push(cl);
        }
        t.rows.push(row);
      });
      if (!t.rows.length) return table(2, 2);
      var cw = Array.isArray(x.cw) ? x.cw.slice(0, cols).map(function (v) { return num(v, 0.001, 1, 0); }) : [];
      while (cw.length < cols) cw.push(0);
      var known = cw.filter(Boolean), rest = Math.max(0, 1 - known.reduce(function (a, b) { return a + b; }, 0)), missing = cw.length - known.length;
      cw = cw.map(function (v) { return v || (missing ? Math.max(0.03, rest / missing) : 0.1); });
      var sum = cw.reduce(function (a, b) { return a + b; }, 0);
      t.cw = cw.map(function (v) { return v / sum; });
      fixMerges(t);
      return t;
    }
    if (x.type === "img") {
      return { id: id, type: "img", blob: typeof x.blob === "string" && ID_RE.test(x.blob) ? x.blob : "", w: Math.round(num(x.w, 4, 4000, 200)), h: Math.round(num(x.h, 4, 6000, 150)),
        align: x.align === "left" || x.align === "right" ? x.align : "center", alt: str(x.alt, 500) };
    }
    if (x.type === "pb") return { id: id, type: "pb" };
    if (x.type === "toc") return { id: id, type: "toc" };
    var b = para(STYLES[x.style] ? x.style : "normal", null, { id: id });
    b.text = cleanSpans(x.text, budget);
    b.align = ["left", "center", "right", "justify"].indexOf(x.align) >= 0 ? x.align : "";
    b.dir = x.dir === "rtl" || x.dir === "ltr" ? x.dir : "";
    b.ind = Math.floor(num(x.ind, 0, MAX_LEVEL, 0));
    b.fi = x.fi === "first" || x.fi === "hang" ? x.fi : "";
    b.lh = num(x.lh, 0.5, 5, 0);
    b.sb = num(x.sb, 0, 200, null); b.sa = num(x.sa, 0, 200, null);
    b.list = cleanList(x.list);
    if (x.m && typeof x.m === "object") { var mk = cleanSpans([withText(x.m, "x")])[0]; if (mk) { delete mk.t; delete mk.a; if (Object.keys(mk).length) b.m = mk; } }
    return b;
  }
  function cleanPage(p) {
    p = p && typeof p === "object" ? p : {};
    var size = SIZES[p.size] ? p.size : p.size === "custom" ? "custom" : "a4";
    var orient = p.orient === "landscape" ? "landscape" : "portrait";
    var dims = size === "custom" ? [num(p.w, 50, 1200, 210), num(p.h, 50, 1200, 297)] : SIZES[size].slice();
    if (size !== "custom") { dims.sort(function (a, b) { return a - b; }); if (orient === "landscape") dims.reverse(); }
    else orient = dims[0] > dims[1] ? "landscape" : "portrait";
    var o = { size: size, orient: orient, w: dims[0], h: dims[1] };
    ["mt", "mr", "mb", "ml"].forEach(function (k) { o[k] = num(p[k], 0, 150, 25.4); });
    // margins can't eat the page: keep at least 30 mm of text area each way
    if (o.w - o.ml - o.mr < 30) { o.ml = o.mr = Math.max(0, (o.w - 30) / 2); }
    if (o.h - o.mt - o.mb < 30) { o.mt = o.mb = Math.max(0, (o.h - 30) / 2); }
    o.hd = num(p.hd, 0, 100, 12.7); o.fd = num(p.fd, 0, 100, 12.7);
    return o;
  }
  function cleanHf(h) {
    h = h && typeof h === "object" ? h : {};
    var al = function (v, d) { return v === "left" || v === "center" || v === "right" ? v : d; };
    return { header: str(h.header, 500).replace(/[\r\n]+/g, " "), footer: str(h.footer, 500).replace(/[\r\n]+/g, " "), ha: al(h.ha, "right"), fa: al(h.fa, "center"), first: h.first === true };
  }
  function cleanBase(b) {
    b = b && typeof b === "object" ? b : {};
    return { font: typeof b.font === "string" && FONT_RE.test(b.font) ? b.font : "Calibri", size: num(b.size, 4, 72, 11), line: num(b.line, 0.8, 4, 1.15), after: num(b.after, 0, 72, 8) };
  }
  function cleanDoc(x) {
    x = x && typeof x === "object" ? x : {};
    var d = { v: 1, id: typeof x.id === "string" && ID_RE.test(x.id) ? x.id : uid(), title: str(x.title, 200), created: num(x.created, 0, 1e14, Date.now()), updated: num(x.updated, 0, 1e14, Date.now()),
      page: cleanPage(x.page), hf: cleanHf(x.hf), base: cleanBase(x.base), blocks: [] };
    var budget = 5e6;
    (Array.isArray(x.blocks) ? x.blocks.slice(0, 50000) : []).forEach(function (b) {
      if (budget <= 0) return;
      var c = cleanBlock(b, Math.min(200000, budget));
      budget -= c.type === "p" ? len(c.text) + 1 : 50;
      d.blocks.push(c);
    });
    return normalize(d);
  }
  function newDoc(props) {
    var d = cleanDoc(props || {});
    if (!d.blocks.length) d.blocks.push(para("normal"));
    return d;
  }
  // Invariants: at least one block, the last block is a paragraph, tables are rectangular with sane merges.
  function normalize(d) {
    if (!d.blocks.length || !isText(d.blocks[d.blocks.length - 1])) d.blocks.push(para("normal"));
    return d;
  }
  function fixMerges(t) {
    var R = t.rows.length, C = t.rows[0].length;
    t.rows.forEach(function (row) { row.forEach(function (c) { delete c.m; }); });
    for (var r = 0; r < R; r++) for (var c = 0; c < C; c++) {
      var cl = t.rows[r][c];
      if (cl.m) { cl.cs = cl.rs = 1; continue; }
      cl.cs = clamp(cl.cs || 1, 1, C - c); cl.rs = clamp(cl.rs || 1, 1, R - r);
      // shrink a merge that would overlap one already placed
      for (var rr = r; rr < r + cl.rs; rr++) for (var cc = c; cc < c + cl.cs; cc++) {
        if ((rr !== r || cc !== c) && t.rows[rr][cc].m) { cl.cs = 1; cl.rs = 1; }
      }
      for (rr = r; rr < r + cl.rs; rr++) for (cc = c; cc < c + cl.cs; cc++) if (rr !== r || cc !== c) { t.rows[rr][cc].m = true; t.rows[rr][cc].cs = t.rows[rr][cc].rs = 1; }
    }
    return t;
  }

  // ---------- positions ----------
  function holder(d, p) {
    var b = d.blocks[p.b];
    if (!b) return null;
    if (b.type === "p") return b;
    if (b.type === "table" && p.r >= 0 && b.rows[p.r] && b.rows[p.r][p.c]) return b.rows[p.r][p.c];
    return null;
  }
  function P(b, o, r, c) { return { b: b, r: r == null ? -1 : r, c: c == null ? -1 : c, o: o || 0 }; }
  function cmp(a, b) { return a.b - b.b || (a.r - b.r) || (a.c - b.c) || (a.o - b.o); }
  function order(a, f) { return cmp(a, f) <= 0 ? [a, f] : [f, a]; }
  function holderLen(d, p) { var h = holder(d, p); return h ? len(h.text) : 1; }
  function startPos(d) { return firstIn(d, 0); }
  function endPos(d) { return lastIn(d, d.blocks.length - 1); }
  function firstIn(d, bi) {
    var b = d.blocks[bi];
    if (!b) return P(0, 0);
    if (b.type === "table") { for (var r = 0; r < b.rows.length; r++) for (var c = 0; c < b.rows[r].length; c++) if (!b.rows[r][c].m) return P(bi, 0, r, c); }
    return P(bi, 0);
  }
  function lastIn(d, bi) {
    var b = d.blocks[bi];
    if (!b) return P(0, 0);
    if (b.type === "table") { for (var r = b.rows.length - 1; r >= 0; r--) for (var c = b.rows[r].length - 1; c >= 0; c--) if (!b.rows[r][c].m) return P(bi, len(b.rows[r][c].text), r, c); }
    return P(bi, b.type === "p" ? len(b.text) : 1);
  }
  // Every text holder (paragraphs and table cells) touched by [a, f), in order, with the offsets inside it.
  function segments(d, a, f) {
    var s = order(a, f), A = s[0], F = s[1], out = [];
    for (var bi = A.b; bi <= F.b; bi++) {
      var b = d.blocks[bi];
      if (!b) continue;
      if (b.type === "p") {
        var L = len(b.text);
        out.push({ b: bi, r: -1, c: -1, h: b, from: bi === A.b && A.r < 0 ? Math.min(A.o, L) : 0, to: bi === F.b && F.r < 0 ? Math.min(F.o, L) : L });
      } else if (b.type === "table") {
        for (var r = 0; r < b.rows.length; r++) for (var c = 0; c < b.rows[r].length; c++) {
          var cl = b.rows[r][c];
          if (cl.m) continue;
          var here = P(bi, 0, r, c), CL = len(cl.text);
          if (bi === A.b && A.r >= 0 && cmp(P(bi, CL, r, c), A) < 0) continue;
          if (bi === F.b && F.r >= 0 && cmp(here, F) > 0) continue;
          out.push({ b: bi, r: r, c: c, h: cl, from: bi === A.b && A.r === r && A.c === c ? A.o : 0, to: bi === F.b && F.r === r && F.c === c ? F.o : CL });
        }
      } else out.push({ b: bi, atom: true });
    }
    return out;
  }
  function textBetween(d, a, f) {
    return segments(d, a, f).map(function (s) { return s.atom ? "" : plain(slice(s.h.text, s.from, s.to)); }).join("\n");
  }

  // ---------- edits (mutate d, return the new caret position) ----------
  function insertTextAt(d, p, text, marks) {
    var h = holder(d, p);
    if (!h) return p;
    var t = String(text).replace(/\r\n?/g, "\n");
    if (p.r < 0) t = t; // paragraphs may hold soft line breaks too
    var mk = marks || (len(h.text) ? marksAt(h.text, p.o) : h.m || {});
    h.text = insertText(h.text, p.o, t, mk);
    if (h.m && len(h.text)) delete h.m;
    return P(p.b, p.o + t.length, p.r, p.c);
  }
  function deleteRange(d, a, f) {
    var s = order(a, f), A = s[0], F = s[1];
    if (!cmp(A, F)) return A;
    if (A.b === F.b && A.r === F.r && A.c === F.c) {
      var h = holder(d, A);
      if (h) { h.text = cut(h.text, A.o, F.o); return A; }
      if (d.blocks[A.b] && d.blocks[A.b].type !== "table" && F.o > A.o) { d.blocks.splice(A.b, 1); normalize(d); return firstIn(d, Math.min(A.b, d.blocks.length - 1)); }
      return A;
    }
    // clear the text of the table cells that are covered; paragraphs at the ends are trimmed and joined
    var segs = segments(d, A, F), remove = [];
    function tableCovered(bi) {
      var s0 = firstIn(d, bi), s1 = lastIn(d, bi);
      return (A.b < bi || (A.b === bi && cmp(A, s0) <= 0)) && (F.b > bi || (F.b === bi && cmp(F, s1) >= 0));
    }
    segs.forEach(function (sg) {
      if (sg.atom) { var cover = (sg.b > A.b || A.o === 0) && (sg.b < F.b || F.o >= 1); if (cover) remove.push(sg.b); return; }
      if (sg.r >= 0) { if (tableCovered(sg.b)) { if (remove.indexOf(sg.b) < 0) remove.push(sg.b); } else sg.h.text = cut(sg.h.text, sg.from, sg.to); return; }
      if (sg.b !== A.b && sg.b !== F.b) remove.push(sg.b);
    });
    var first = d.blocks[A.b], last = d.blocks[F.b], caret = A;
    if (first.type === "p" && last.type === "p" && A.b !== F.b) {
      var mk = len(first.text) === A.o && !A.o ? marksAt(last.text, F.o) : null;
      first.text = concat(slice(first.text, 0, A.o), slice(last.text, F.o));
      if (!len(first.text) && mk && Object.keys(mk).length) first.m = mk;
      remove.push(F.b);
    } else if (first.type === "p" && A.b !== F.b) { first.text = slice(first.text, 0, A.o); }
    else if (last.type === "p" && A.b !== F.b) { last.text = slice(last.text, F.o); }
    if (first.type !== "p" && remove.indexOf(A.b) >= 0) caret = null;
    remove.sort(function (x, y) { return y - x; }).filter(function (v, i, arr) { return arr.indexOf(v) === i; }).forEach(function (bi) { d.blocks.splice(bi, 1); });
    normalize(d);
    if (!caret) { var nb = Math.min(A.b, d.blocks.length - 1); caret = firstIn(d, nb); }
    return caret;
  }
  function continuation(b) {
    var nb = para(STYLES[b.style].next || b.style === "title" || b.style === "subtitle" ? "normal" : b.style, null, { align: b.style === "title" || STYLES[b.style].next ? "" : b.align, dir: b.dir, ind: b.ind, fi: b.fi, lh: b.lh });
    if (b.style === "code" || b.style === "quote") { nb.sb = b.sb; nb.sa = b.sa; }
    if (b.list) nb.list = { k: b.list.k, l: b.list.l, c: b.list.k === "todo" ? false : undefined };
    if (nb.list && nb.list.c === undefined) delete nb.list.c;
    return nb;
  }
  // Enter. In a table cell it is a line break; on an empty list item it ends the list (or steps out a level).
  function splitAt(d, p) {
    var b = d.blocks[p.b];
    if (p.r >= 0) return insertTextAt(d, p, "\n");
    if (!b || b.type !== "p") {
      var np = para("normal");
      d.blocks.splice(p.b + (p.o > 0 ? 1 : 0), 0, np);
      return P(p.b + (p.o > 0 ? 1 : 0), 0);
    }
    var L = len(b.text);
    if (b.list && !L) {
      if (b.list.l > 0) b.list.l--; else b.list = null;
      return P(p.b, 0);
    }
    var nb = continuation(b), mk = marksAt(b.text, p.o);
    if (p.o === 0 && L > 0) {
      // Enter at the start keeps this paragraph as it is and opens an empty one above it
      var top = continuation(b);
      top.style = b.style; top.align = b.align;
      if (b.list) top.list = { k: b.list.k, l: b.list.l };
      d.blocks.splice(p.b, 0, top);
      return P(p.b + 1, 0);
    }
    nb.text = slice(b.text, p.o);
    b.text = slice(b.text, 0, p.o);
    if (!len(nb.text) && Object.keys(mk).length) { delete mk.a; if (Object.keys(mk).length) nb.m = mk; }
    d.blocks.splice(p.b + 1, 0, nb);
    return P(p.b + 1, 0);
  }
  // Backspace at the very start of a text holder: first remove the list / indent, then join with the paragraph above.
  function backAtStart(d, p) {
    var b = d.blocks[p.b];
    if (p.r >= 0) return p;
    if (b.list) { b.list = null; return p; }
    if (b.ind > 0) { b.ind--; return p; }
    if (b.fi) { b.fi = ""; return p; }
    if (p.b === 0) { if (b.style !== "normal" && !len(b.text)) b.style = "normal"; return p; }
    var prev = d.blocks[p.b - 1];
    if (prev.type === "p") {
      var off = len(prev.text);
      if (!off && len(b.text)) { d.blocks.splice(p.b - 1, 1); return P(p.b - 1, 0); }
      prev.text = concat(prev.text, b.text);
      d.blocks.splice(p.b, 1);
      normalize(d);
      return P(p.b - 1, off);
    }
    if (prev.type === "table") { if (!len(b.text) && p.b < d.blocks.length - 1) { d.blocks.splice(p.b, 1); } return lastIn(d, p.b - 1); }
    d.blocks.splice(p.b - 1, 1); // an image, page break or contents block above: remove it
    return P(p.b - 1, 0);
  }
  // Delete at the very end: join the next paragraph in.
  function deleteAtEnd(d, p) {
    if (p.r >= 0) return p;
    var b = d.blocks[p.b], next = d.blocks[p.b + 1];
    if (!next) return p;
    if (next.type === "p") { var off = len(b.text); b.text = concat(b.text, next.text); d.blocks.splice(p.b + 1, 1); normalize(d); return P(p.b, off); }
    if (next.type !== "table") { d.blocks.splice(p.b + 1, 1); normalize(d); }
    return p;
  }
  // Paste / insert a list of blocks at the caret. The first and last pasted paragraphs join the text around the caret.
  function insertBlocks(d, p, blocks) {
    blocks = (blocks || []).filter(Boolean);
    if (!blocks.length) return p;
    var h = holder(d, p);
    if (p.r >= 0) {
      // inside a cell only text fits: flatten
      var txt = [];
      blocks.forEach(function (b, k) { if (b.type === "p") { if (k) txt.push({ t: "\n" }); txt = txt.concat(b.text); } });
      h.text = insertSpans(h.text, p.o, txt);
      return P(p.b, p.o + len(txt), p.r, p.c);
    }
    var b = d.blocks[p.b];
    if (!b || b.type !== "p") { Array.prototype.splice.apply(d.blocks, [p.b + 1, 0].concat(blocks)); normalize(d); return lastIn(d, p.b + blocks.length); }
    if (blocks.length === 1 && blocks[0].type === "p") {
      b.text = insertSpans(b.text, p.o, blocks[0].text);
      if (!len(slice(b.text, 0, p.o)) && p.o === 0 && b.style === "normal" && blocks[0].style !== "normal") b.style = blocks[0].style;
      return P(p.b, p.o + len(blocks[0].text));
    }
    var right = slice(b.text, p.o);
    b.text = slice(b.text, 0, p.o);
    var at = p.b, rest = blocks.slice();
    if (rest[0].type === "p") {
      var f = rest.shift();
      if (!len(b.text)) { f.text = concat(b.text, f.text); d.blocks[at] = f; } else b.text = concat(b.text, f.text);
    } else if (!len(b.text)) { d.blocks.splice(at, 1); at--; }
    Array.prototype.splice.apply(d.blocks, [at + 1, 0].concat(rest));
    var lastIdx = at + rest.length, last = d.blocks[lastIdx];
    var caret;
    if (last && last.type === "p" && rest.length) { caret = P(lastIdx, len(last.text)); last.text = concat(last.text, right); }
    else if (len(right) || !rest.length) {
      var tail = para(b.style || "normal", right, { dir: b.dir });
      d.blocks.splice(lastIdx + 1, 0, tail); caret = P(lastIdx + 1, 0);
      if (!rest.length) caret = P(at, len(d.blocks[at].text));
    } else caret = firstIn(d, Math.min(lastIdx + 1, d.blocks.length - 1));
    normalize(d);
    return caret;
  }
  function insertBlockAt(d, p, block) {
    // splits the paragraph at the caret when needed, puts block between, returns a caret after it
    var b = d.blocks[p.b];
    if (p.r >= 0) { p = P(p.b, 0); b = d.blocks[p.b]; d.blocks.splice(p.b + 1, 0, block); normalize(d); return firstIn(d, p.b + 2 < d.blocks.length ? p.b + 2 : d.blocks.length - 1); }
    if (b && b.type === "p") {
      var L = len(b.text);
      if (!L) { d.blocks.splice(p.b, 0, block); normalize(d); return P(p.b + 1, 0); }
      if (p.o >= L) { d.blocks.splice(p.b + 1, 0, block); normalize(d); var nxt = d.blocks[p.b + 2]; if (!nxt || nxt.type !== "p") d.blocks.splice(p.b + 2, 0, para("normal")); return P(p.b + 2, 0); }
      if (p.o === 0) { d.blocks.splice(p.b, 0, block); return P(p.b + 1, 0); }
      var right = continuation(b); right.style = b.style; right.text = slice(b.text, p.o); b.text = slice(b.text, 0, p.o);
      d.blocks.splice(p.b + 1, 0, block, right);
      return P(p.b + 2, 0);
    }
    d.blocks.splice(p.b + 1, 0, block); normalize(d);
    return firstIn(d, p.b + 2 < d.blocks.length ? p.b + 2 : d.blocks.length - 1);
  }

  // ---------- formatting over a range ----------
  function setMarkRange(d, a, f, key, value) {
    segments(d, a, f).forEach(function (s) { if (!s.atom && s.to > s.from) s.h.text = setMark(s.h.text, s.from, s.to, key, value); });
  }
  function hasMarkRange(d, a, f, key) {
    var segs = segments(d, a, f).filter(function (s) { return !s.atom && s.to > s.from; });
    return segs.length > 0 && segs.every(function (s) { return hasMark(s.h.text, s.from, s.to, key); });
  }
  function toggleMarkRange(d, a, f, key) { var on = !hasMarkRange(d, a, f, key); setMarkRange(d, a, f, key, on ? true : null); return on; }
  // the marks every character of the range shares (for the toolbar); a collapsed range reports the typing marks
  function marksInRange(d, a, f) {
    if (!cmp(a, f)) { var h = holder(d, a); return h ? (len(h.text) ? marksAt(h.text, a.o) : marksOf(h.m)) : {}; }
    var common = null;
    segments(d, a, f).forEach(function (s) {
      if (s.atom || s.to <= s.from) return;
      slice(s.h.text, s.from, s.to).forEach(function (sp) {
        if (!common) { common = marksOf(sp); return; }
        Object.keys(common).forEach(function (k) { if (common[k] !== sp[k]) delete common[k]; });
      });
    });
    return common || {};
  }
  function clearFormat(d, a, f) {
    segments(d, a, f).forEach(function (s) {
      if (s.atom || s.to <= s.from) return;
      var mid = slice(s.h.text, s.from, s.to).map(function (sp) { var o = { t: sp.t }; if (sp.a) o.a = sp.a; return o; });
      s.h.text = norm(slice(s.h.text, 0, s.from).concat(mid, slice(s.h.text, s.to)));
    });
  }
  function parasIn(d, a, f) {
    var s = order(a, f), out = [];
    for (var bi = s[0].b; bi <= s[1].b; bi++) if (d.blocks[bi] && d.blocks[bi].type === "p") out.push(bi);
    return out;
  }
  function cellsIn(d, a, f) { return segments(d, a, f).filter(function (s) { return !s.atom && s.r >= 0; }).map(function (s) { return s.h; }); }
  function setStyle(d, a, f, style) {
    parasIn(d, a, f).forEach(function (bi) {
      var b = d.blocks[bi];
      b.style = STYLES[style] ? style : "normal";
      if (HEADINGS[style] != null) { b.list = null; }
    });
  }
  function setAlign(d, a, f, al) {
    parasIn(d, a, f).forEach(function (bi) { d.blocks[bi].align = al; });
    cellsIn(d, a, f).forEach(function (c) { c.al = al; });
  }
  function toggleList(d, a, f, kind) {
    var list = parasIn(d, a, f);
    var all = list.length && list.every(function (bi) { var l = d.blocks[bi].list; return l && l.k === kind; });
    list.forEach(function (bi) {
      var b = d.blocks[bi];
      if (all) b.list = null;
      else { b.list = { k: kind, l: b.list ? b.list.l : Math.min(b.ind, MAX_LEVEL) }; if (kind === "todo") b.list.c = false; b.ind = 0; if (HEADINGS[b.style] != null || b.style === "code") b.style = "normal"; }
    });
    return !all;
  }
  function indent(d, a, f, delta) {
    parasIn(d, a, f).forEach(function (bi) {
      var b = d.blocks[bi];
      if (b.list) b.list.l = clamp(b.list.l + delta, 0, MAX_LEVEL);
      else b.ind = clamp(b.ind + delta, 0, MAX_LEVEL);
    });
  }

  // ---------- tables ----------
  function tableUnmergeAll(t, test) {
    t.rows.forEach(function (row, r) { row.forEach(function (c, ci) { if (!c.m && (c.cs > 1 || c.rs > 1) && test(r, ci, c)) { c.cs = 1; c.rs = 1; } }); });
    fixMerges(t);
  }
  function addRow(t, at) {
    var C = t.rows[0].length, row = [];
    tableUnmergeAll(t, function (r, c, cl) { return r < at && r + cl.rs > at; });
    for (var c = 0; c < C; c++) { var nc = cell(); var src = t.rows[Math.min(Math.max(0, at - 1), t.rows.length - 1)][c]; if (src && src.bg && at > 0 && !(t.hdr && at - 1 === 0)) nc.bg = src.bg; row.push(nc); }
    t.rows.splice(clamp(at, 0, t.rows.length), 0, row);
    fixMerges(t);
  }
  function delRow(t, at) {
    if (t.rows.length <= 1) return false;
    tableUnmergeAll(t, function (r, c, cl) { return r <= at && r + cl.rs > at && cl.rs > 1; });
    t.rows.splice(at, 1); fixMerges(t); return true;
  }
  function addCol(t, at) {
    var C = t.rows[0].length;
    if (C >= 63) return false;
    tableUnmergeAll(t, function (r, c, cl) { return c < at && c + cl.cs > at; });
    t.rows.forEach(function (row) { row.splice(clamp(at, 0, C), 0, cell()); });
    var w = 1 / (C + 1);
    t.cw = t.cw.map(function (v) { return v * (1 - w); });
    t.cw.splice(clamp(at, 0, C), 0, w);
    fixMerges(t); return true;
  }
  function delCol(t, at) {
    var C = t.rows[0].length;
    if (C <= 1) return false;
    tableUnmergeAll(t, function (r, c, cl) { return c <= at && c + cl.cs > at && cl.cs > 1; });
    t.rows.forEach(function (row) { row.splice(at, 1); });
    var gone = t.cw.splice(at, 1)[0], rest = 1 - gone;
    t.cw = t.cw.map(function (v) { return v / rest; });
    fixMerges(t); return true;
  }
  function mergeCells(t, r0, c0, r1, c1) {
    var R0 = Math.min(r0, r1), R1 = Math.max(r0, r1), C0 = Math.min(c0, c1), C1 = Math.max(c0, c1);
    if (R0 === R1 && C0 === C1) return false;
    tableUnmergeAll(t, function (r, c, cl) { return !(r >= R0 && r + cl.rs - 1 <= R1 && c >= C0 && c + cl.cs - 1 <= C1) && r <= R1 && r + cl.rs - 1 >= R0 && c <= C1 && c + cl.cs - 1 >= C0; });
    var top = t.rows[R0][C0], text = top.text;
    for (var r = R0; r <= R1; r++) for (var c = C0; c <= C1; c++) {
      if (r === R0 && c === C0) continue;
      var cl = t.rows[r][c];
      if (len(cl.text)) text = concat(concat(text, len(text) ? [{ t: "\n" }] : []), cl.text);
      cl.text = []; cl.cs = cl.rs = 1;
    }
    top.text = text; top.cs = C1 - C0 + 1; top.rs = R1 - R0 + 1;
    fixMerges(t); return true;
  }
  function splitCell(t, r, c) { var cl = t.rows[r][c]; if (cl.cs === 1 && cl.rs === 1) return false; cl.cs = cl.rs = 1; fixMerges(t); return true; }

  // ---------- numbering, headings, contents ----------
  var BULLETS = ["•", "◦", "▪"];
  function roman(n) { var map = [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"], [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]], r = ""; map.forEach(function (p) { while (n >= p[0]) { r += p[1]; n -= p[0]; } }); return r; }
  function alpha(n) { var s = ""; while (n > 0) { n--; s = String.fromCharCode(97 + (n % 26)) + s; n = Math.floor(n / 26); } return s; }
  function numLabel(n, level) { var k = level % 3; return (k === 0 ? String(n) : k === 1 ? alpha(n) : roman(n)) + "."; }
  // list markers for every paragraph: contiguous list paragraphs form one list; deeper levels restart under a new parent
  function listLabels(d) {
    var labels = [], counters = [];
    d.blocks.forEach(function (b, i) {
      if (b.type !== "p" || !b.list) { if (!(b.type === "p" && !len(b.text) && d.blocks[i + 1] && d.blocks[i + 1].list && i > 0 && d.blocks[i - 1].list)) counters = []; labels[i] = ""; return; }
      var l = b.list.l;
      counters.length = l + 1;
      if (b.list.k === "ol") { counters[l] = (counters[l] || 0) + 1; labels[i] = numLabel(counters[l], l); }
      else { counters[l] = counters[l] || 0; labels[i] = b.list.k === "todo" ? (b.list.c ? "☑" : "☐") : BULLETS[l % 3]; }
    });
    return labels;
  }
  function headings(d) {
    var out = [];
    d.blocks.forEach(function (b, i) { if (b.type === "p" && HEADINGS[b.style] != null && b.style !== "title") { var t = plain(b.text).replace(/\s+/g, " ").trim(); if (t) out.push({ b: i, id: b.id, level: HEADINGS[b.style], text: t }); } });
    return out;
  }
  // contents entries: the headings after the contents block (or all of them), with page numbers from a layout
  function tocEntries(d, pageOf) {
    return headings(d).map(function (h) { return { b: h.b, id: h.id, level: h.level, text: h.text, page: pageOf ? pageOf[h.id] || 0 : 0 }; });
  }

  // ---------- counting, finding ----------
  var WORD = /[\p{L}\p{N}\p{M}]+(?:['’\-.][\p{L}\p{N}\p{M}]+)*/gu;
  function countText(t) {
    var w = (t.match(WORD) || []).length;
    return { words: w, chars: t.replace(/\n/g, "").length, charsNoSpaces: t.replace(/\s/g, "").length };
  }
  function allText(d) {
    var parts = [];
    d.blocks.forEach(function (b) {
      if (b.type === "p") parts.push(plain(b.text));
      else if (b.type === "table") b.rows.forEach(function (row) { row.forEach(function (c) { if (!c.m) parts.push(plain(c.text)); }); });
    });
    return parts;
  }
  function count(d) {
    var parts = allText(d), c = countText(parts.join("\n"));
    c.paragraphs = d.blocks.filter(function (b) { return b.type === "p" && len(b.text); }).length;
    return c;
  }
  function fold(s, caseSensitive) { s = String(s || ""); return caseSensitive ? s : s.toLowerCase(); }
  function findAll(d, q, opts) {
    opts = opts || {};
    var out = [];
    if (!q) return out;
    var needle = fold(q, opts.matchCase), word = /[\p{L}\p{N}_]/u;
    function scan(h, b, r, c) {
      var t = fold(plain(h.text), opts.matchCase), from = 0, at;
      while ((at = t.indexOf(needle, from)) >= 0 && out.length < 5000) {
        var ok = !opts.whole || ((at === 0 || !word.test(t[at - 1])) && (at + needle.length >= t.length || !word.test(t[at + needle.length])));
        if (ok) out.push({ a: P(b, at, r, c), f: P(b, at + needle.length, r, c) });
        from = at + Math.max(1, needle.length);
      }
    }
    d.blocks.forEach(function (b, bi) {
      if (b.type === "p") scan(b, bi, -1, -1);
      else if (b.type === "table") b.rows.forEach(function (row, r) { row.forEach(function (cl, c) { if (!cl.m) scan(cl, bi, r, c); }); });
    });
    return out;
  }
  function replaceAt(d, m, text) {
    var h = holder(d, m.a);
    if (!h) return m.a;
    var mk = slice(h.text, m.a.o, m.a.o + 1)[0];
    h.text = cut(h.text, m.a.o, m.f.o);
    h.text = insertSpans(h.text, m.a.o, text ? [withText(mk, text)] : []);
    return P(m.a.b, m.a.o + String(text).length, m.a.r, m.a.c);
  }
  function replaceAll(d, q, text, opts) {
    var ms = findAll(d, q, opts);
    for (var k = ms.length - 1; k >= 0; k--) replaceAt(d, ms[k], text);
    return ms.length;
  }

  // ---------- grapheme / word steps (for Backspace and Delete) ----------
  var seg = typeof Intl !== "undefined" && Intl.Segmenter ? new Intl.Segmenter(undefined, { granularity: "grapheme" }) : null;
  function prevStop(t, o) {
    if (o <= 0) return 0;
    if (seg) { var last = 0; for (var it = seg.segment(t.slice(0, o))[Symbol.iterator](), x = it.next(); !x.done; x = it.next()) last = x.value.index; return last; }
    var c = t.charCodeAt(o - 1); return c >= 0xdc00 && c <= 0xdfff && o > 1 ? o - 2 : o - 1;
  }
  function nextStop(t, o) {
    if (o >= t.length) return t.length;
    if (seg) { var it = seg.segment(t.slice(o))[Symbol.iterator](), x = it.next(); x = it.next(); return x.done ? t.length : o + x.value.index; }
    var c = t.charCodeAt(o); return c >= 0xd800 && c <= 0xdbff ? o + 2 : o + 1;
  }
  function wordBack(t, o) { var s = t.slice(0, o), m = /[\p{L}\p{N}\p{M}_]+[^\p{L}\p{N}\p{M}_\n]*$|[^\p{L}\p{N}\p{M}_\n]+$|\n$/u.exec(s); return m ? o - m[0].length : Math.max(0, o - 1); }
  function wordFwd(t, o) { var s = t.slice(o), m = /^[^\p{L}\p{N}\p{M}_\n]*[\p{L}\p{N}\p{M}_]+|^[^\p{L}\p{N}\p{M}_\n]+|^\n/u.exec(s); return m ? o + m[0].length : Math.min(t.length, o + 1); }

  // ---------- page geometry ----------
  function geometry(d) {
    var p = d.page;
    return { w: p.w * MM, h: p.h * MM, mt: p.mt * MM, mr: p.mr * MM, mb: p.mb * MM, ml: p.ml * MM, hd: p.hd * MM, fd: p.fd * MM,
      cw: (p.w - p.ml - p.mr) * MM, ch: (p.h - p.mt - p.mb) * MM };
  }
  function styleOf(d, b) {
    var s = STYLES[b.style] || STYLES.normal, base = d.base;
    return {
      size: s.size || base.size, bold: !!s.bold, italic: !!s.italic, color: s.color || "", font: s.font || "", mono: !!s.mono, bg: s.bg || "",
      lh: b.lh || s.lh || base.line, sb: b.sb != null ? b.sb : s.sb != null ? s.sb : 0, sa: b.sa != null ? b.sa : s.sa != null ? s.sa : base.after,
      ind: (b.list ? b.list.l + 1 : b.ind) + (s.ind || 0), bar: !!s.bar
    };
  }
  function fillTokens(text, ctx) {
    return String(text || "").replace(/\{(page|pages|title|date)\}/g, function (m, k) { return k === "page" ? String(ctx.page) : k === "pages" ? String(ctx.pages) : k === "title" ? ctx.title || "" : ctx.date || ""; });
  }
  function clone(o) { return JSON.parse(JSON.stringify(o)); }

  window.DM = {
    ID_RE: ID_RE, BOOL: BOOL, KEYS: KEYS, HL: HL, SIZES: SIZES, STYLES: STYLES, STYLE_KEYS: STYLE_KEYS, HEADINGS: HEADINGS, MM: MM, PT: PT, MAX_LEVEL: MAX_LEVEL, HEX: HEX, FONT_RE: FONT_RE,
    uid: uid, esc: esc, safeUrl: safeUrl, clamp: clamp, num: num, clone: clone,
    norm: norm, plain: plain, len: len, slice: slice, concat: concat, cut: cut, insertSpans: insertSpans, insertText: insertText, marksAt: marksAt, marksOf: marksOf,
    setMark: setMark, hasMark: hasMark, cleanSpans: cleanSpans, spansOf: spansOf, withText: withText,
    para: para, cell: cell, table: table, image: image, pagebreak: pagebreak, toc: toc, isText: isText,
    cleanBlock: cleanBlock, cleanDoc: cleanDoc, cleanPage: cleanPage, cleanHf: cleanHf, cleanBase: cleanBase, newDoc: newDoc, normalize: normalize, fixMerges: fixMerges,
    P: P, cmp: cmp, order: order, holder: holder, holderLen: holderLen, startPos: startPos, endPos: endPos, firstIn: firstIn, lastIn: lastIn, segments: segments, textBetween: textBetween,
    insertTextAt: insertTextAt, deleteRange: deleteRange, splitAt: splitAt, backAtStart: backAtStart, deleteAtEnd: deleteAtEnd, insertBlocks: insertBlocks, insertBlockAt: insertBlockAt,
    setMarkRange: setMarkRange, hasMarkRange: hasMarkRange, toggleMarkRange: toggleMarkRange, marksInRange: marksInRange, clearFormat: clearFormat,
    parasIn: parasIn, setStyle: setStyle, setAlign: setAlign, toggleList: toggleList, indent: indent,
    addRow: addRow, delRow: delRow, addCol: addCol, delCol: delCol, mergeCells: mergeCells, splitCell: splitCell,
    listLabels: listLabels, numLabel: numLabel, headings: headings, tocEntries: tocEntries, count: count, countText: countText, allText: allText,
    findAll: findAll, replaceAt: replaceAt, replaceAll: replaceAll, prevStop: prevStop, nextStop: nextStop, wordBack: wordBack, wordFwd: wordFwd,
    geometry: geometry, styleOf: styleOf, fillTokens: fillTokens
  };
})();
