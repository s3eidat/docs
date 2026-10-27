/* Docs: turns model blocks into DOM. Text only ever goes in through text nodes; links only get hrefs that passed
 * DM.safeUrl; colours, fonts and sizes were validated by the model, and are set through the CSSOM anyway. */
(function () {
  "use strict";
  var M = window.DM;
  var IND = 48; // one indent step = 0.5 in at 96 dpi
  var FALLBACK = {
    calibri: "Carlito, 'Segoe UI', Arial, sans-serif", cambria: "Caladea, Georgia, serif", "times new roman": "'Liberation Serif', Tinos, Times, serif",
    arial: "'Liberation Sans', Arimo, Helvetica, sans-serif", helvetica: "Arial, sans-serif", georgia: "Gelasio, serif", garamond: "'EB Garamond', Georgia, serif",
    "courier new": "'Liberation Mono', Cousine, monospace", consolas: "'Cascadia Mono', Menlo, 'Courier New', monospace", verdana: "sans-serif", tahoma: "sans-serif",
    "segoe ui": "system-ui, sans-serif", inter: "system-ui, sans-serif"
  };
  var SERIF = /times|georgia|garamond|cambria|book|palatino|serif|naskh|amiri|traditional|minion|baskerville|didot|caladea/i;
  var MONO = /consol|courier|mono|code/i;
  function fontStack(f) {
    if (!f) return "";
    var fb = FALLBACK[f.toLowerCase()] || (MONO.test(f) ? "monospace" : SERIF.test(f) ? "serif" : "sans-serif");
    return '"' + f.replace(/["\\]/g, "") + '", ' + fb;
  }
  function pt(v) { return v + "pt"; }

  // ---------- text ----------
  function spanEl(s) {
    var el = document.createElement("span"), st = el.style, deco = [];
    if (s.b) st.fontWeight = "700";
    if (s.i) st.fontStyle = "italic";
    if (s.u) deco.push("underline");
    if (s.s) deco.push("line-through");
    if (deco.length) st.textDecorationLine = deco.join(" ");
    if (s.f) st.fontFamily = fontStack(s.f);
    if (s.z) st.fontSize = pt(s.z);
    if (s.c) st.color = s.c;
    if (s.h) st.backgroundColor = s.h;
    if (s.sup || s.sub) { st.verticalAlign = s.sup ? "super" : "sub"; st.fontSize = s.z ? pt(Math.round(s.z * 0.65 * 2) / 2) : "0.65em"; st.lineHeight = "0"; }
    var parts = s.t.split("\n");
    parts.forEach(function (p, k) {
      if (k) el.appendChild(document.createElement("br"));
      if (p) el.appendChild(document.createTextNode(p));
    });
    return el;
  }
  function renderSpans(el, spans) {
    el.textContent = "";
    var link = null, linkUrl = null;
    spans.forEach(function (s) {
      var node = spanEl(s);
      if (s.a) {
        if (!link || linkUrl !== s.a) {
          link = document.createElement("a"); link.className = "dc-a"; link.href = s.a; link.rel = "noopener noreferrer"; link.target = "_blank";
          link.title = s.a + "  (Ctrl+click to open)"; linkUrl = s.a; el.appendChild(link);
        }
        link.appendChild(node);
      } else { link = null; linkUrl = null; el.appendChild(node); }
    });
    if (!spans.length || /\n$/.test(spans[spans.length - 1].t)) { var br = document.createElement("br"); br.setAttribute("data-tail", ""); el.appendChild(br); }
  }
  // text length / offsets inside a rendered text holder: text nodes count, <br> counts 1, the tail <br> counts 0
  function walkLen(n) {
    var L = 0;
    for (var c = n.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 3) L += c.data.length;
      else if (c.nodeType === 1 && !c.hasAttribute("data-x")) { if (c.nodeName === "BR") L += c.hasAttribute("data-tail") ? 0 : 1; else L += walkLen(c); }
    }
    return L;
  }
  function textOf(n) {
    var s = "";
    for (var c = n.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 3) s += c.data;
      else if (c.nodeType === 1 && !c.hasAttribute("data-x")) { if (c.nodeName === "BR") s += c.hasAttribute("data-tail") ? "" : "\n"; else s += textOf(c); }
    }
    return s;
  }
  function localOffset(root, node, off) {
    if (node === root) { var L = 0; for (var k = 0; k < off && k < root.childNodes.length; k++) { var ch = root.childNodes[k]; L += ch.nodeType === 3 ? ch.data.length : ch.nodeName === "BR" ? (ch.hasAttribute("data-tail") ? 0 : 1) : ch.nodeType === 1 ? walkLen(ch) : 0; } return L; }
    var acc = 0, found = false;
    (function walk(n) {
      for (var c = n.firstChild; c && !found; c = c.nextSibling) {
        if (c === node) {
          if (c.nodeType === 3) acc += Math.min(off, c.data.length);
          else for (var k = 0; k < off && k < c.childNodes.length; k++) { var x = c.childNodes[k]; acc += x.nodeType === 3 ? x.data.length : x.nodeName === "BR" ? (x.hasAttribute && x.hasAttribute("data-tail") ? 0 : 1) : x.nodeType === 1 ? walkLen(x) : 0; }
          found = true; return;
        }
        if (c.nodeType === 3) acc += c.data.length;
        else if (c.nodeType === 1 && !c.hasAttribute("data-x")) {
          if (c.nodeName === "BR") acc += c.hasAttribute("data-tail") ? 0 : 1;
          else if (c.contains(node)) { walk(c); return; } else acc += walkLen(c);
        }
      }
    })(root);
    return acc;
  }
  function domPoint(root, o) {
    var acc = 0, res = null;
    (function walk(n) {
      for (var c = n.firstChild; c && !res; c = c.nextSibling) {
        if (c.nodeType === 3) {
          if (o <= acc + c.data.length) { res = { node: c, off: o - acc }; return; }
          acc += c.data.length;
        } else if (c.nodeType === 1 && !c.hasAttribute("data-x")) {
          if (c.nodeName === "BR") {
            var idx = Array.prototype.indexOf.call(n.childNodes, c);
            if (o === acc) { res = { node: n, off: idx }; return; }
            if (!c.hasAttribute("data-tail")) acc += 1;
            if (o === acc && (!c.nextSibling || c.nextSibling.nodeName === "BR")) { res = { node: n, off: idx + 1 }; return; }
          } else walk(c);
        }
      }
    })(root);
    return res || { node: root, off: root.childNodes.length };
  }

  // ---------- blocks ----------
  function dirOf(b, text) {
    if (b.dir) return b.dir;
    var m = /[\p{L}]/u.exec(text || M.plain(b.text));
    return m && /[\u0590-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFF]/.test(m[0]) ? "rtl" : "ltr";
  }
  function paraEl(doc, b, label, from, to, o) {
    o = o || {};
    var st = M.styleOf(doc, b), L = M.len(b.text);
    if (to == null) to = L;
    var first = from === 0, last = to >= L;
    var el = document.createElement("p");
    el.className = "dc-p dc-s-" + b.style + (first ? "" : " dc-cont") + (last ? "" : " dc-more");
    el.setAttribute("data-b", b.id); el.setAttribute("data-o", String(from));
    el.dir = dirOf(b);
    var s = el.style, indPx = st.ind * IND;
    s.paddingTop = first ? pt(st.sb) : "0";
    s.paddingBottom = last ? pt(st.sa) : "0";
    s.lineHeight = String(st.lh);
    if (b.align) s.textAlign = b.align === "left" ? "start" : b.align === "right" ? "end" : b.align;
    if (b.align === "justify" && !last) s.textAlignLast = "justify";
    if (st.bar) { s.marginInlineStart = Math.max(0, indPx - 18) + "px"; s.paddingInlineStart = "15px"; }
    else if (indPx) s.paddingInlineStart = indPx + "px";
    if (first && b.fi === "first") s.textIndent = IND + "px";
    if (b.fi === "hang") { s.paddingInlineStart = (indPx + IND) + "px"; if (first) s.textIndent = -IND + "px"; }
    if (b.list && first && label) {
      el.classList.add("dc-li", "dc-li-" + b.list.k);
      if (b.list.k === "todo" && b.list.c) el.classList.add("dc-done");
      el.setAttribute("data-mk", label);
      s.setProperty("--mk", (indPx - 24) + "px");
      s.setProperty("--sb", pt(st.sb));
      s.setProperty("--lh", String(st.lh));
    }
    if (b.m && !L) {
      // an empty paragraph shows the size it will type in
      if (b.m.z) s.fontSize = pt(b.m.z);
      if (b.m.f) s.fontFamily = fontStack(b.m.f);
    }
    if (!L && o.placeholder) { el.classList.add("dc-ph"); el.setAttribute("data-ph", o.placeholder); }
    renderSpans(el, M.slice(b.text, from, to));
    return el;
  }
  function cellEl(doc, t, r, c, rep) {
    var cl = t.rows[r][c], td = document.createElement(rep || (t.hdr && r === 0) ? "th" : "td");
    if (cl.cs > 1) td.colSpan = cl.cs;
    if (cl.rs > 1) td.rowSpan = cl.rs;
    if (cl.bg) td.style.backgroundColor = cl.bg;
    var div = document.createElement("div");
    div.className = "dc-cell";
    if (!rep) { div.setAttribute("data-b", t.id); div.setAttribute("data-r", String(r)); div.setAttribute("data-c", String(c)); div.setAttribute("data-o", "0"); td.setAttribute("data-r", String(r)); td.setAttribute("data-c", String(c)); }
    div.dir = dirOf({ dir: "", text: cl.text });
    if (cl.al) div.style.textAlign = cl.al === "left" ? "start" : cl.al === "right" ? "end" : cl.al;
    renderSpans(div, cl.text);
    td.appendChild(div);
    return td;
  }
  function tableEl(doc, t, r0, r1, rep) {
    var el = document.createElement("table");
    el.className = "dc-tbl" + (t.hdr ? " dc-hdr-row" : "");
    el.setAttribute("data-b", t.id); el.setAttribute("data-r0", String(r0));
    var dirs = 0;
    t.rows.forEach(function (row) { row.forEach(function (cl) { if (dirOf({ dir: "", text: cl.text }) === "rtl" && M.len(cl.text)) dirs++; }); });
    if (dirs && dirs * 2 >= t.rows.length * t.rows[0].length * 0.5) el.dir = "rtl";
    var cg = document.createElement("colgroup");
    t.cw.forEach(function (w) { var col = document.createElement("col"); col.style.width = (w * 100).toFixed(3) + "%"; cg.appendChild(col); });
    el.appendChild(cg);
    var body = document.createElement("tbody");
    function row(r, isRep) {
      var tr = document.createElement("tr");
      if (!isRep) tr.setAttribute("data-r", String(r)); else { tr.className = "dc-rep"; tr.setAttribute("contenteditable", "false"); tr.setAttribute("data-x", ""); }
      t.rows[r].forEach(function (cl, c) {
        if (cl.m) return;
        var td = cellEl(doc, t, r, c, isRep);
        if (cl.rs > 1 && r + cl.rs > r1) td.rowSpan = Math.max(1, r1 - r);
        tr.appendChild(td);
      });
      body.appendChild(tr);
    }
    if (rep && t.hdr && r0 > 0) row(0, true);
    for (var r = r0; r < r1; r++) row(r, false);
    el.appendChild(body);
    el.style.paddingBottom = "0";
    el.style.marginBottom = pt(doc.base.after);
    return el;
  }
  function imgEl(doc, b, url, maxW, maxH) {
    var fig = document.createElement("figure");
    fig.className = "dc-fig dc-al-" + b.align;
    fig.setAttribute("data-b", b.id); fig.setAttribute("contenteditable", "false");
    var sz = imgSize(b, maxW, maxH);
    var box = document.createElement("span"); box.className = "dc-imgbox";
    box.style.width = sz.w + "px"; box.style.height = sz.h + "px";
    var img = document.createElement("img");
    img.alt = b.alt || ""; img.draggable = false; img.width = sz.w; img.height = sz.h;
    if (url) img.src = url; else box.classList.add("dc-imgmissing");
    box.appendChild(img);
    var hd = document.createElement("span"); hd.className = "dc-handle no-print"; hd.title = "Drag to resize"; box.appendChild(hd);
    fig.appendChild(box);
    fig.style.paddingBottom = pt(doc.base.after);
    return fig;
  }
  function imgSize(b, maxW, maxH) {
    var w = b.w, h = b.h, k = Math.min(1, maxW / w, maxH ? (maxH - 12) / h : 1);
    return { w: Math.max(4, Math.round(w * k)), h: Math.max(4, Math.round(h * k)) };
  }
  function pbEl(b) {
    var el = document.createElement("div");
    el.className = "dc-pbk"; el.setAttribute("data-b", b.id); el.setAttribute("contenteditable", "false");
    var s = document.createElement("span"); s.textContent = "Page break"; el.appendChild(s);
    return el;
  }
  function tocEl(doc, b, entries) {
    var el = document.createElement("div");
    el.className = "dc-toc"; el.setAttribute("data-b", b.id); el.setAttribute("contenteditable", "false");
    var hd = document.createElement("div"); hd.className = "dc-toc-h"; hd.textContent = "Contents"; el.appendChild(hd);
    if (!entries.length) { var e = document.createElement("div"); e.className = "dc-toc-empty"; e.textContent = "Headings (Heading 1 to 3) will be listed here with their page numbers."; el.appendChild(e); }
    entries.forEach(function (en) {
      var row = document.createElement("div");
      row.className = "dc-toc-e dc-toc-l" + en.level; row.setAttribute("data-go", en.id);
      var t = document.createElement("span"); t.className = "dc-toc-t"; t.textContent = en.text; t.dir = "auto";
      var dots = document.createElement("span"); dots.className = "dc-toc-d";
      var n = document.createElement("span"); n.className = "dc-toc-n"; n.textContent = en.page ? String(en.page) : "";
      row.appendChild(t); row.appendChild(dots); row.appendChild(n);
      el.appendChild(row);
    });
    el.style.paddingBottom = pt(doc.base.after);
    return el;
  }

  // ---------- styles for one document (base font, size and the style table) ----------
  function css() {
    var out = [];
    M.STYLE_KEYS.forEach(function (k) {
      var s = M.STYLES[k], r = [];
      if (s.size) r.push("font-size:" + s.size + "pt");
      if (s.bold) r.push("font-weight:700");
      if (s.italic) r.push("font-style:italic");
      if (s.color) r.push("color:" + s.color);
      if (s.font) r.push("font-family:" + fontStack(s.font));
      if (s.bg) r.push("background:" + s.bg);
      if (s.bar) r.push("border-inline-start:3px solid #c5ccd6");
      if (k === "title") r.push("letter-spacing:-.01em");
      if (r.length) out.push(".dc-s-" + k + "{" + r.join(";") + "}");
    });
    return out.join("\n");
  }

  window.DR = {
    IND: IND, fontStack: fontStack, renderSpans: renderSpans, textOf: textOf, walkLen: walkLen, localOffset: localOffset, domPoint: domPoint,
    dirOf: dirOf, paraEl: paraEl, tableEl: tableEl, imgEl: imgEl, imgSize: imgSize, pbEl: pbEl, tocEl: tocEl, css: css
  };
})();
