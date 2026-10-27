/* Docs: text formats. Paste (HTML / plain text) into blocks through a whitelist walker, and export blocks to
 * HTML, Markdown and plain text. HTML is parsed with DOMParser (inert: no scripts run, nothing loads) and only
 * the tags and the handful of inline style properties listed here are read; everything else is dropped. */
(function () {
  "use strict";
  var M = window.DM, esc = M.esc;
  var SKIP = { SCRIPT: 1, STYLE: 1, TEMPLATE: 1, NOSCRIPT: 1, IFRAME: 1, OBJECT: 1, EMBED: 1, SVG: 1, HEAD: 1, TITLE: 1, META: 1, LINK: 1, BUTTON: 1, SELECT: 1, INPUT: 1, TEXTAREA: 1, CANVAS: 1, VIDEO: 1, AUDIO: 1 };
  var BLOCK = { P: 1, DIV: 1, H1: 1, H2: 1, H3: 1, H4: 1, H5: 1, H6: 1, LI: 1, BLOCKQUOTE: 1, PRE: 1, TABLE: 1, UL: 1, OL: 1, HR: 1, SECTION: 1, ARTICLE: 1, HEADER: 1, FOOTER: 1, MAIN: 1, ASIDE: 1, NAV: 1, FIGURE: 1, FIGCAPTION: 1, DL: 1, DT: 1, DD: 1, ADDRESS: 1, CENTER: 1 };
  var NAMED = { black: "#000000", white: "#ffffff", red: "#ff0000", green: "#008000", blue: "#0000ff", yellow: "#ffff00", gray: "#808080", grey: "#808080", orange: "#ffa500", purple: "#800080" };

  function hex(v) {
    if (!v) return null;
    v = String(v).trim().toLowerCase();
    var m;
    if ((m = /^#([0-9a-f]{3})$/.exec(v))) return "#" + m[1].split("").map(function (c) { return c + c; }).join("");
    if (/^#[0-9a-f]{6}$/.test(v)) return v;
    if ((m = /^rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s\/]+([\d.]+%?))?\s*\)$/.exec(v))) {
      if (m[4] != null && parseFloat(m[4]) === 0) return null;
      return "#" + [m[1], m[2], m[3]].map(function (x) { return ("0" + Math.min(255, +x).toString(16)).slice(-2); }).join("");
    }
    return NAMED[v] || null;
  }
  function sizePt(v) {
    var m = /^([\d.]+)(pt|px|em|rem)?$/.exec(String(v || "").trim());
    if (!m) return null;
    var n = parseFloat(m[1]), u = m[2] || "px";
    var p = u === "pt" ? n : u === "px" ? n * 0.75 : n * 11;
    return p >= 4 && p <= 200 ? Math.round(p * 2) / 2 : null;
  }
  function inlineMarks(el, m) {
    var o = Object.assign({}, m), tag = el.nodeName, st = el.style;
    if (tag === "B" || tag === "STRONG") o.b = true;
    if (tag === "I" || tag === "EM" || tag === "CITE") o.i = true;
    if (tag === "U" || tag === "INS") o.u = true;
    if (tag === "S" || tag === "STRIKE" || tag === "DEL") o.s = true;
    if (tag === "SUP") { o.sup = true; delete o.sub; }
    if (tag === "SUB") { o.sub = true; delete o.sup; }
    if (tag === "MARK") o.h = "#ffff00";
    if (tag === "CODE" || tag === "KBD" || tag === "SAMP") o.f = "Consolas";
    if (tag === "A") { var u = M.safeUrl(el.getAttribute("href") || ""); if (u) o.a = u; }
    if (tag === "FONT") { var c0 = hex(el.getAttribute("color")); if (c0) o.c = c0; var f0 = (el.getAttribute("face") || "").split(",")[0].replace(/["']/g, "").trim(); if (M.FONT_RE.test(f0)) o.f = f0; }
    if (st) {
      var fw = st.fontWeight; if (fw === "bold" || fw === "bolder" || +fw >= 600) o.b = true; else if (fw === "normal" || (+fw && +fw < 600)) delete o.b;
      if (st.fontStyle === "italic" || st.fontStyle === "oblique") o.i = true; else if (st.fontStyle === "normal") delete o.i;
      var td = (st.textDecorationLine || st.textDecoration || "");
      if (/underline/.test(td)) o.u = true;
      if (/line-through/.test(td)) o.s = true;
      var c = hex(st.color); if (c && c !== "#000000") o.c = c;
      var bg = hex(st.backgroundColor); if (bg && bg !== "#ffffff") o.h = bg;
      var z = sizePt(st.fontSize); if (z) o.z = z;
      var ff = (st.fontFamily || "").split(",")[0].replace(/["']/g, "").trim();
      if (ff && M.FONT_RE.test(ff) && !/^(inherit|initial|serif|sans-serif|monospace|system-ui)$/i.test(ff)) o.f = ff;
      if (st.verticalAlign === "super") o.sup = true; else if (st.verticalAlign === "sub") o.sub = true;
    }
    return o;
  }
  function alignOf(el) {
    var a = (el.style && el.style.textAlign) || el.getAttribute("align") || "";
    a = a.toLowerCase();
    return a === "center" ? "center" : a === "right" || a === "end" ? "right" : a === "justify" ? "justify" : a === "left" || a === "start" ? "left" : "";
  }

  // HTML -> blocks (+ the data: images found, for the app to store)
  function htmlToBlocks(html) {
    var dom = new DOMParser().parseFromString(String(html || "").slice(0, 4e6), "text/html");
    var blocks = [], images = [], cur = null, total = 0;
    function flush() { if (cur) { cur.text = M.norm(cur.text); blocks.push(cur); cur = null; } }
    function open(props) { flush(); cur = M.para(props.style || "normal", null, { align: props.align || "", dir: props.dir || "", list: props.list || null }); }
    function add(t, m) {
      if (!t || total > 1e6) return;
      if (!cur) open({});
      t = t.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
      cur.text.push(M.withText(m, t)); total += t.length;
    }
    function walk(node, m, ctx) {
      for (var c = node.firstChild; c; c = c.nextSibling) {
        if (blocks.length > 20000) return;
        if (c.nodeType === 3) {
          var t = ctx.pre ? c.data : c.data.replace(/[\t\n\r ]+/g, " ");
          if (!ctx.pre && (!cur || !M.len(cur.text))) t = t.replace(/^ +/, "");
          if (ctx.pre) t.split("\n").forEach(function (line, k) { if (k) open({ style: "code" }); add(line, m); });
          else add(t, m);
          continue;
        }
        if (c.nodeType !== 1) continue;
        var tag = c.nodeName.toUpperCase();
        if (SKIP[tag] || c.namespaceURI && !/xhtml/.test(c.namespaceURI)) continue;
        if (tag === "BR") { if (cur) add("\n", m); else open(ctx); continue; }
        if (tag === "IMG") {
          var src = c.getAttribute("src") || "";
          if (/^data:image\/(png|jpeg|jpg|gif|webp|bmp);base64,/i.test(src) && src.length < 30e6) {
            flush();
            var w = Math.min(2000, +c.getAttribute("width") || 300), hh = Math.min(3000, +c.getAttribute("height") || Math.round(w * 0.66));
            var b = M.image("", w, hh, { alt: (c.getAttribute("alt") || "").slice(0, 300) });
            blocks.push(b); images.push({ block: b, src: src });
          }
          continue;
        }
        if (tag === "HR") { flush(); continue; }
        if (tag === "TABLE") { flush(); var tb = tableFrom(c); if (tb) blocks.push(tb); continue; }
        if (tag === "UL" || tag === "OL") {
          flush();
          var lvl = ctx.list ? Math.min(M.MAX_LEVEL, ctx.list.l + 1) : 0;
          var isTodo = !!c.querySelector(":scope > li > input[type=checkbox]");
          walk(c, m, Object.assign({}, ctx, { list: { k: isTodo ? "todo" : tag === "OL" ? "ol" : "ul", l: lvl }, inList: true }));
          flush();
          continue;
        }
        if (BLOCK[tag]) {
          var nctx = Object.assign({}, ctx);
          if (/^H[1-6]$/.test(tag)) nctx.style = tag === "H1" ? "h1" : tag === "H2" ? "h2" : "h3";
          if (tag === "BLOCKQUOTE") nctx.style = "quote";
          if (tag === "PRE") { nctx.style = "code"; nctx.pre = true; }
          var al = alignOf(c); if (al) nctx.align = al;
          var d = (c.getAttribute("dir") || "").toLowerCase(); if (d === "rtl" || d === "ltr") nctx.dir = d;
          var props = { style: nctx.style, align: nctx.align, dir: nctx.dir };
          if (tag === "LI") {
            var cb = c.querySelector(":scope > input[type=checkbox]");
            props.list = ctx.list ? { k: ctx.list.k, l: ctx.list.l } : { k: "ul", l: 0 };
            if (props.list.k === "todo") props.list.c = !!(cb && cb.checked);
            nctx.list = props.list; open(props);
            walk(c, inlineMarks(c, m), Object.assign(nctx, { style: "normal" }));
            flush(); continue;
          }
          if (tag === "DIV" && cur && !c.querySelector("p,div,h1,h2,h3,h4,h5,h6,li,table,blockquote,pre")) { open(props); walk(c, inlineMarks(c, m), nctx); flush(); continue; }
          flush();
          if (!/^(DIV|SECTION|ARTICLE|MAIN|HEADER|FOOTER|ASIDE|NAV|FIGURE|CENTER|DL)$/.test(tag)) open(props);
          walk(c, inlineMarks(c, m), nctx);
          flush();
          continue;
        }
        walk(c, inlineMarks(c, m), ctx);
      }
    }
    function cellSpans(td) {
      var out = [];
      (function w(n, m) {
        for (var c = n.firstChild; c; c = c.nextSibling) {
          if (c.nodeType === 3) { var t = c.data.replace(/[\t\n\r ]+/g, " "); if (t) out.push(M.withText(m, t)); continue; }
          if (c.nodeType !== 1 || SKIP[c.nodeName.toUpperCase()] || c.nodeName === "TABLE" || (c.namespaceURI && !/xhtml/.test(c.namespaceURI))) continue;
          if (c.nodeName === "BR") { out.push({ t: "\n" }); continue; }
          if (BLOCK[c.nodeName] && out.length && !/\n$/.test(out[out.length - 1].t)) out.push({ t: "\n" });
          w(c, inlineMarks(c, m));
        }
      })(td, {});
      var s = M.norm(out), L = M.len(s);
      while (L && /\s$/.test(M.plain(s))) { s = M.slice(s, 0, L - 1); L--; }
      return s;
    }
    function tableFrom(tbl) {
      var trs = Array.prototype.filter.call(tbl.querySelectorAll("tr"), function (tr) { return tr.closest("table") === tbl; }).slice(0, 500);
      if (!trs.length) return null;
      var grid = [], cols = 0;
      trs.forEach(function (tr, r) {
        grid[r] = grid[r] || [];
        var c = 0;
        Array.prototype.forEach.call(tr.children, function (td) {
          if (td.nodeName !== "TD" && td.nodeName !== "TH") return;
          while (grid[r][c]) c++;
          var cs = Math.max(1, Math.min(40, +td.getAttribute("colspan") || 1)), rs = Math.max(1, Math.min(trs.length - r, +td.getAttribute("rowspan") || 1));
          var cl = M.cell(cellSpans(td)); cl.cs = cs; cl.rs = rs;
          var bg = hex(td.style && td.style.backgroundColor) || hex(td.getAttribute("bgcolor")); if (bg) cl.bg = bg;
          for (var rr = r; rr < r + rs; rr++) { grid[rr] = grid[rr] || []; for (var cc = c; cc < c + cs; cc++) grid[rr][cc] = rr === r && cc === c ? cl : { text: [], bg: "", al: "", cs: 1, rs: 1, m: true }; }
          c += cs;
        });
      });
      grid.forEach(function (row) { cols = Math.max(cols, row.length); });
      cols = Math.min(cols, 40);
      if (!cols) return null;
      var t = M.table(grid.length, cols);
      t.rows = grid.map(function (row) { var out = []; for (var c = 0; c < cols; c++) out.push(row[c] || M.cell()); return out; });
      t.hdr = !!tbl.querySelector("th");
      return M.cleanBlock(t);
    }
    walk(dom.body || dom.documentElement, {}, { style: "normal" });
    flush();
    // trailing whitespace-only paragraphs from markup noise
    blocks = blocks.filter(function (b, k) { return b.type !== "p" || M.plain(b.text).trim() || b.list || (k > 0 && k < blocks.length - 1); });
    blocks.forEach(function (b) { if (b.type === "p") { b.text = trimSpans(b.text); } });
    return { blocks: blocks, images: images };
  }
  function trimSpans(s) {
    var t = M.plain(s), a = t.length - t.replace(/^[ \t]+/, "").length, z = t.replace(/[ \t]+$/, "").length;
    return M.slice(s, a, z);
  }
  function textToBlocks(text) {
    return String(text || "").replace(/\r\n?/g, "\n").slice(0, 2e6).split("\n").map(function (line) { return M.para("normal", line.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")); });
  }

  // ---------- Markdown in ----------
  function inlineMd(s) {
    var out = [], re = /(\*\*([^*]+)\*\*|__([^_]+)__|\*([^*\s][^*]*)\*|_([^_\s][^_]*)_|~~([^~]+)~~|`([^`]+)`|\[([^\]]+)\]\(([^)\s]+)\))/g, last = 0, m;
    while ((m = re.exec(s))) {
      if (m.index > last) out.push({ t: s.slice(last, m.index) });
      if (m[2] || m[3]) out.push({ t: m[2] || m[3], b: true });
      else if (m[4] || m[5]) out.push({ t: m[4] || m[5], i: true });
      else if (m[6]) out.push({ t: m[6], s: true });
      else if (m[7]) out.push({ t: m[7], f: "Consolas" });
      else if (m[8]) { var u = M.safeUrl(m[9]); out.push(u ? { t: m[8], a: u, u: true, c: "#1155cc" } : { t: m[8] }); }
      last = re.lastIndex;
    }
    if (last < s.length) out.push({ t: s.slice(last) });
    return M.cleanSpans(out);
  }
  function mdToBlocks(md) {
    var lines = String(md || "").replace(/\r\n?/g, "\n").slice(0, 2e6).split("\n"), blocks = [], i = 0, m;
    while (i < lines.length && blocks.length < 20000) {
      var line = lines[i];
      if (/^```/.test(line)) { i++; while (i < lines.length && !/^```/.test(lines[i])) { blocks.push(M.para("code", lines[i])); i++; } i++; continue; }
      if (/^\s*\|.*\|\s*$/.test(line) && i + 1 < lines.length && /^\s*\|?\s*:?-{2,}/.test(lines[i + 1])) {
        var rows = [];
        while (i < lines.length && /^\s*\|.*\|\s*$/.test(lines[i])) { if (!/^\s*\|?\s*:?-{2,}/.test(lines[i])) rows.push(lines[i].trim().replace(/^\||\|$/g, "").split("|").map(function (c) { return c.trim(); })); i++; }
        var cols = Math.max.apply(null, rows.map(function (r) { return r.length; })), t = M.table(rows.length, Math.min(cols, 40));
        rows.forEach(function (r, ri) { r.slice(0, 40).forEach(function (c, ci) { t.rows[ri][ci].text = inlineMd(c); }); });
        blocks.push(t); continue;
      }
      if ((m = /^(#{1,6})\s+(.*)$/.exec(line))) { blocks.push(M.para(m[1].length === 1 ? "h1" : m[1].length === 2 ? "h2" : "h3", inlineMd(m[2]))); i++; continue; }
      if ((m = /^(\s*)[-*+]\s+\[([ xX])\]\s+(.*)$/.exec(line))) { blocks.push(M.para("normal", inlineMd(m[3]), { list: { k: "todo", l: Math.min(8, Math.floor(m[1].length / 2)), c: m[2] !== " " } })); i++; continue; }
      if ((m = /^(\s*)[-*+]\s+(.*)$/.exec(line))) { blocks.push(M.para("normal", inlineMd(m[2]), { list: { k: "ul", l: Math.min(8, Math.floor(m[1].length / 2)) } })); i++; continue; }
      if ((m = /^(\s*)\d{1,4}[.)]\s+(.*)$/.exec(line))) { blocks.push(M.para("normal", inlineMd(m[2]), { list: { k: "ol", l: Math.min(8, Math.floor(m[1].length / 3)) } })); i++; continue; }
      if ((m = /^>\s?(.*)$/.exec(line))) { blocks.push(M.para("quote", inlineMd(m[1]))); i++; continue; }
      if (/^(-{3,}|\*{3,}|_{3,})\s*$/.test(line)) { i++; continue; }
      if (!line.trim()) { i++; continue; }
      var buf = [line]; i++;
      while (i < lines.length && lines[i].trim() && !/^(#{1,6}\s|```|\s*[-*+]\s|\s*\d{1,4}[.)]\s|>|\s*\|)/.test(lines[i])) { buf.push(lines[i]); i++; }
      blocks.push(M.para("normal", inlineMd(buf.join(" "))));
    }
    return blocks;
  }

  // ---------- export ----------
  function spanHtml(s) {
    var st = [];
    if (s.f) st.push("font-family:" + window.DR.fontStack(s.f).replace(/"/g, "'"));
    if (s.z) st.push("font-size:" + s.z + "pt");
    if (s.c) st.push("color:" + s.c);
    if (s.h) st.push("background-color:" + s.h);
    var t = esc(s.t).replace(/\n/g, "<br>");
    if (st.length) t = '<span style="' + st.join(";") + '">' + t + "</span>";
    if (s.b) t = "<strong>" + t + "</strong>";
    if (s.i) t = "<em>" + t + "</em>";
    if (s.u) t = "<u>" + t + "</u>";
    if (s.s) t = "<s>" + t + "</s>";
    if (s.sup) t = "<sup>" + t + "</sup>";
    if (s.sub) t = "<sub>" + t + "</sub>";
    if (s.a) t = '<a href="' + esc(s.a) + '">' + t + "</a>";
    return t;
  }
  function spansHtml(spans) { return spans.map(spanHtml).join(""); }
  function blocksToHtml(blocks, opts) {
    opts = opts || {};
    var out = [], i = 0, labels = M.listLabels({ blocks: blocks });
    var doc = opts.doc || { base: { size: 11, line: 1.15, after: 8 }, title: "" };
    while (i < blocks.length) {
      var b = blocks[i];
      if (b.type === "p" && b.list) {
        // nested lists from flat levels
        var stack = [], liOpen = [];
        while (i < blocks.length && blocks[i].type === "p" && blocks[i].list) {
          var x = blocks[i], tag = x.list.k === "ol" ? "ol" : "ul", lvl = x.list.l + 1;
          while (stack.length > lvl) { if (liOpen.pop()) out.push("</li>"); out.push("</" + stack.pop() + ">"); }
          while (stack.length < lvl) {
            if (stack.length && !liOpen[liOpen.length - 1]) { out.push("<li>"); liOpen[liOpen.length - 1] = true; }
            out.push("<" + tag + (x.list.k === "todo" ? ' style="list-style:none"' : "") + ">"); stack.push(tag); liOpen.push(false);
          }
          if (liOpen[liOpen.length - 1]) out.push("</li>");
          var box = x.list.k === "todo" ? (x.list.c ? "☑ " : "☐ ") : "";
          out.push("<li" + (x.dir ? ' dir="' + x.dir + '"' : ' dir="auto"') + ">" + box + spansHtml(x.text));
          liOpen[liOpen.length - 1] = true;
          i++;
        }
        while (stack.length) { if (liOpen.pop()) out.push("</li>"); out.push("</" + stack.pop() + ">"); }
        continue;
      }
      if (b.type === "p") {
        var tagp = { title: "h1", subtitle: "p", h1: "h1", h2: "h2", h3: "h3", quote: "blockquote", code: "pre" }[b.style] || "p";
        if (b.style === "title" && opts.full) tagp = "h1";
        var stl = [];
        if (b.align) stl.push("text-align:" + b.align);
        if (b.ind) stl.push("margin-left:" + (b.ind * 0.5) + "in");
        if (b.fi === "first") stl.push("text-indent:0.5in");
        if (b.lh) stl.push("line-height:" + b.lh);
        var cls = b.style === "title" ? ' class="title"' : b.style === "subtitle" ? ' class="subtitle"' : "";
        out.push("<" + tagp + cls + (b.dir ? ' dir="' + b.dir + '"' : ' dir="auto"') + (stl.length ? ' style="' + stl.join(";") + '"' : "") + ">" + (M.len(b.text) ? spansHtml(b.text) : "<br>") + "</" + tagp + ">");
      } else if (b.type === "table") {
        var rows = b.rows.map(function (row, r) {
          return "<tr>" + row.map(function (cl) {
            if (cl.m) return "";
            var tg = b.hdr && r === 0 ? "th" : "td";
            return "<" + tg + (cl.cs > 1 ? ' colspan="' + cl.cs + '"' : "") + (cl.rs > 1 ? ' rowspan="' + cl.rs + '"' : "") + (cl.bg ? ' style="background-color:' + cl.bg + '"' : "") + ' dir="auto">' + spansHtml(cl.text) + "</" + tg + ">";
          }).join("") + "</tr>";
        });
        out.push('<table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse;width:100%">' + rows.join("") + "</table>");
      } else if (b.type === "img") {
        var src = opts.images && opts.images[b.blob];
        if (src) out.push('<p style="text-align:' + b.align + '"><img src="' + esc(src) + '" width="' + b.w + '" height="' + b.h + '" alt="' + esc(b.alt) + '"></p>');
      } else if (b.type === "pb") out.push('<div style="break-after:page;page-break-after:always"></div>');
      else if (b.type === "toc" && opts.doc) {
        out.push('<nav class="toc"><p><strong>Contents</strong></p>' + M.tocEntries(opts.doc, opts.pageOf).map(function (e) { return '<p style="margin-left:' + (e.level - 1) * 1.5 + 'em">' + esc(e.text) + (e.page ? " … " + e.page : "") + "</p>"; }).join("") + "</nav>");
      }
      i++;
    }
    return out.join("\n");
  }
  function docToHtml(doc, images, pageOf) {
    var p = doc.page, fs = window.DR.fontStack(doc.base.font).replace(/"/g, "'");
    var css = "@page{size:" + p.w + "mm " + p.h + "mm;margin:" + p.mt + "mm " + p.mr + "mm " + p.mb + "mm " + p.ml + "mm}" +
      "body{font-family:" + fs + ";font-size:" + doc.base.size + "pt;line-height:" + doc.base.line + ";max-width:" + (p.w - p.ml - p.mr) + "mm;margin:2em auto;padding:0 1em;color:#111}" +
      "p,li{margin:0 0 " + doc.base.after + "pt}h1,h2,h3{color:#1d3557;line-height:1.15}h1{font-size:18pt}h2{font-size:14pt}h3{font-size:12pt}.title{font-size:26pt;margin:0 0 6pt}.subtitle{font-size:14pt;color:#5a6372}" +
      "blockquote{margin:0 0 10pt 0.25in;padding-left:15px;border-left:3px solid #c5ccd6;color:#4b5563;font-style:italic}pre{font-family:Consolas,monospace;font-size:10pt;background:#f3f4f6;margin:0;padding:0 4px;white-space:pre-wrap}" +
      "table{margin:0 0 " + doc.base.after + "pt}td,th{vertical-align:top;border:1px solid #9aa3ad}th{background:#eef1f5;text-align:start}img{max-width:100%;height:auto}";
    return "<!doctype html>\n<html lang=\"en\">\n<head>\n<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n<title>" + esc(doc.title || "Document") + "</title>\n<style>" + css + "</style>\n</head>\n<body>\n" +
      blocksToHtml(doc.blocks, { images: images, doc: doc, full: true, pageOf: pageOf }) + "\n</body>\n</html>\n";
  }
  function mdEsc(t) { return t.replace(/([\\`*_~\[\]#|])/g, "\\$1"); }
  function spansMd(spans) {
    return spans.map(function (s) {
      var t = s.f && /consol|courier|mono/i.test(s.f) && !/`/.test(s.t) ? "`" + s.t + "`" : mdEsc(s.t);
      var pre = "", lead = /^\s*/.exec(t)[0], trail = /\s*$/.exec(t)[0], core = t.trim();
      if (!core) return t;
      if (s.s) core = "~~" + core + "~~";
      if (s.i) core = "*" + core + "*";
      if (s.b) core = "**" + core + "**";
      if (s.a) core = "[" + core + "](" + s.a.replace(/\)/g, "%29") + ")";
      return pre + lead + core.replace(/\n/g, "  \n") + trail;
    }).join("");
  }
  function docToMarkdown(doc, imgName) {
    var out = [], labels = M.listLabels(doc), inCode = false;
    doc.blocks.forEach(function (b, i) {
      if (b.type === "p" && b.style === "code") { if (!inCode) { out.push("```"); inCode = true; } out.push(M.plain(b.text)); return; }
      if (inCode) { out.push("```", ""); inCode = false; }
      if (b.type === "p") {
        var t = spansMd(b.text);
        if (b.list) {
          var pad = new Array(b.list.l + 1).join(b.list.k === "ol" ? "   " : "  ");
          out.push(pad + (b.list.k === "ol" ? (labels[i].replace(/[^0-9]/g, "") || "1") + ". " : b.list.k === "todo" ? "- [" + (b.list.c ? "x" : " ") + "] " : "- ") + t);
          if (!doc.blocks[i + 1] || !doc.blocks[i + 1].list) out.push("");
          return;
        }
        var pre = { title: "# ", h1: "# ", h2: "## ", h3: "### ", quote: "> " }[b.style] || "";
        if (b.style === "title" || b.style === "subtitle") pre = b.style === "title" ? "# " : "*";
        out.push(pre + t + (b.style === "subtitle" && t ? "*" : ""), "");
      } else if (b.type === "table") {
        b.rows.forEach(function (row, r) {
          out.push("| " + row.map(function (c) { return c.m ? "" : spansMd(c.text).replace(/\n/g, "<br>").replace(/\|/g, "\\|"); }).join(" | ") + " |");
          if (r === 0) out.push("|" + row.map(function () { return " --- "; }).join("|") + "|");
        });
        out.push("");
      } else if (b.type === "img") out.push("![" + mdEsc(b.alt || "image") + "](" + (imgName ? imgName(b) : "image") + ")", "");
      else if (b.type === "pb") out.push("---", "");
      else if (b.type === "toc") { M.headings(doc).forEach(function (h) { out.push(new Array(h.level).join("  ") + "- " + mdEsc(h.text)); }); out.push(""); }
    });
    if (inCode) out.push("```");
    return out.join("\n").replace(/\n{3,}/g, "\n\n").trim() + "\n";
  }
  function docToText(doc) {
    var out = [], labels = M.listLabels(doc);
    doc.blocks.forEach(function (b, i) {
      if (b.type === "p") out.push((b.list ? new Array(b.list.l + 1).join("    ") + labels[i] + " " : "") + M.plain(b.text));
      else if (b.type === "table") b.rows.forEach(function (row) { out.push(row.map(function (c) { return c.m ? "" : M.plain(c.text).replace(/\n/g, " "); }).join("\t")); });
      else if (b.type === "pb") out.push("\f");
      else if (b.type === "toc") M.headings(doc).forEach(function (h) { out.push(new Array(h.level).join("  ") + h.text); });
    });
    return out.join("\n").replace(/\n?\f\n?/g, "\n\n") + "\n";
  }

  window.DIO = { htmlToBlocks: htmlToBlocks, textToBlocks: textToBlocks, mdToBlocks: mdToBlocks, blocksToHtml: blocksToHtml, docToHtml: docToHtml, docToMarkdown: docToMarkdown, docToText: docToText, hex: hex };
})();
