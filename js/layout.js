/* Docs: the pagination engine. Measures blocks in a hidden column exactly as wide as the page's text area and
 * fills fixed-size pages. Paragraphs split at line boundaries (found by measuring prefixes that end at break
 * opportunities, so both halves wrap exactly like the whole), tables split between rows (never inside a merged
 * row group, header row repeated), headings keep with the next line, single orphan lines move down.
 * Measurements are cached per block content, so typing re-measures only the paragraph being typed in. */
(function () {
  "use strict";
  var M = window.DM, R = window.DR;

  function Layout(measurer) {
    var cache = {}; // block id -> {sig, w, h:{}, t:{}}
    function entry(b, sig, w) {
      var e = cache[b.id];
      if (!e || e.sig !== sig || e.w !== w) e = cache[b.id] = { sig: sig, w: w, h: {}, t: null };
      e.used = true;
      return e;
    }
    function measureEl(el) {
      measurer.appendChild(el);
      var h = el.getBoundingClientRect().height;
      var cs = getComputedStyle(el);
      h += (parseFloat(cs.marginTop) || 0) + (parseFloat(cs.marginBottom) || 0);
      measurer.removeChild(el);
      return h;
    }

    function paginate(doc, opts) {
      opts = opts || {};
      var g = M.geometry(doc), H = g.ch, W = g.cw;
      measurer.style.width = W + "px";
      measurer.style.fontFamily = R.fontStack(doc.base.font); measurer.style.fontSize = doc.base.size + "pt"; measurer.style.lineHeight = String(doc.base.line);
      Object.keys(cache).forEach(function (k) { cache[k].used = false; });
      var labels = opts.labels || M.listLabels(doc), entries = opts.entries || M.tocEntries(doc);
      var pages = [], cur = null, y = 0, pageOf = {}, blockPage = {};
      function newPage() { cur = { items: [] }; pages.push(cur); y = 0; }
      newPage();
      function place(it, h) { it.h = h; cur.items.push(it); y += h; var b = doc.blocks[it.bi]; if (!(b.id in blockPage)) blockPage[b.id] = pages.length; }

      function pH(e, b, label, from, to) {
        var key = from + "|" + to;
        if (!(key in e.h)) e.h[key] = measureEl(R.paraEl(doc, b, label, from, to));
        return e.h[key];
      }
      function split(e, b, label, from, avail, st) {
        var text = M.plain(b.text), L = text.length, C = [];
        for (var k = from + 1; k <= L; k++) {
          var ch = text[k - 1];
          if (k === L || ch === "\n" || ((/\s/.test(ch) || ch === "-" || ch === "​") && !/\s/.test(text[k]))) C.push(k);
        }
        if (C.length < 2) { C = []; for (k = from + 1; k <= L; k++) C.push(k); }
        if (pH(e, b, label, from, C[0]) > avail) return null;
        var lo = 0, hi = C.length - 1;
        if (pH(e, b, label, from, C[hi]) <= avail) return null;
        while (lo < hi) { var mid = (lo + hi) >> 1; if (pH(e, b, label, from, C[mid]) > avail) hi = mid; else lo = mid + 1; }
        var i = lo, target = pH(e, b, label, from, C[i]);
        lo = 1; hi = i;
        while (lo < hi) { mid = (lo + hi) >> 1; if (pH(e, b, label, from, C[mid]) >= target - 0.5) hi = mid; else lo = mid + 1; }
        var to = C[lo - 1];
        if (to <= from) return null;
        var h = pH(e, b, label, from, to);
        // orphan control: don't leave a single line at the bottom of a page
        var lineH = st.size * M.PT * st.lh, content = h - (from === 0 ? st.sb * M.PT : 0);
        if (y > 0 && content < lineH * 1.6) return null;
        return { to: to, h: h };
      }
      function tableInfo(e, t) {
        if (!e.t) {
          var el = R.tableEl(doc, t, 0, t.rows.length, false);
          measurer.appendChild(el);
          var rows = Array.prototype.map.call(el.querySelectorAll("tbody > tr"), function (tr) { return tr.getBoundingClientRect().height; });
          var total = el.getBoundingClientRect().height, mb = parseFloat(getComputedStyle(el).marginBottom) || 0;
          measurer.removeChild(el);
          var sum = rows.reduce(function (a, b) { return a + b; }, 0);
          // a break is allowed before row r when no merged cell above runs into it
          var ok = [true], reach = 0;
          t.rows.forEach(function (row, r) { if (r) ok[r] = reach <= r; row.forEach(function (c) { if (!c.m) reach = Math.max(reach, r + c.rs); }); });
          ok[t.rows.length] = true;
          e.t = { rows: rows, extra: Math.max(0, total - sum), mb: mb, ok: ok };
        }
        return e.t;
      }

      doc.blocks.forEach(function (b, bi) {
        if (b.type === "pb") { place({ k: "pb", bi: bi }, 0); newPage(); return; }
        if (b.type === "img") {
          var sz = R.imgSize(b, W, H), ih = sz.h + doc.base.after * M.PT;
          if (y > 0 && y + sz.h > H + 0.5) newPage();
          place({ k: "img", bi: bi }, ih); return;
        }
        if (b.type === "toc") {
          var et = entry(b, JSON.stringify(entries.map(function (x) { return [x.level, x.text]; })) + doc.base.size + doc.base.font, W);
          if (!("toc" in et.h)) et.h.toc = measureEl(R.tocEl(doc, b, entries));
          var th = et.h.toc;
          if (y > 0 && y + th > H + 0.5) newPage();
          place({ k: "toc", bi: bi }, th); return;
        }
        if (b.type === "table") {
          var e2 = entry(b, JSON.stringify(b) + doc.base.size + doc.base.font, W), info = tableInfo(e2, b), R0 = 0, NR = b.rows.length;
          var guard = 0;
          while (R0 < NR && guard++ < 10000) {
            var rep = R0 > 0 && b.hdr ? info.rows[0] : 0, acc = rep + info.extra, best = R0;
            for (var k = R0; k < NR; k++) { acc += info.rows[k]; if (acc > H - y + 0.5) break; if (info.ok[k + 1]) best = k + 1; }
            if (best === R0) {
              if (y > 0) { newPage(); continue; }
              best = R0 + 1; while (best < NR && !info.ok[best]) best++;
            }
            var hgt = rep + info.extra; for (k = R0; k < best; k++) hgt += info.rows[k];
            if (best >= NR) hgt += info.mb;
            place({ k: "table", bi: bi, r0: R0, r1: best, rep: R0 > 0 }, hgt);
            R0 = best;
            if (R0 < NR) newPage();
          }
          return;
        }
        // paragraph
        var label = labels[bi] || "", e = entry(b, JSON.stringify(b) + "|" + label + "|" + doc.base.size + doc.base.font + doc.base.line + doc.base.after, W);
        var st = M.styleOf(doc, b), L = M.len(b.text), from = 0, sa = st.sa * M.PT, loops = 0, heading = M.HEADINGS[b.style] != null;
        while (loops++ < 5000) {
          var full = pH(e, b, label, from, L), fit = full - sa;
          if (from === 0 && heading && y > 0 && st.size) {
            var nextLine = doc.base.size * M.PT * doc.base.line * 1.2;
            if (y + fit + nextLine > H + 0.5 && fit + nextLine <= H) { newPage(); continue; }
          }
          if (y + fit <= H + 0.5) { place({ k: "p", bi: bi, from: from, to: L }, full); if (from === 0 && heading) pageOf[b.id] = pages.length; break; }
          var sp = L - from > 1 ? split(e, b, label, from, H - y, st) : null;
          if (sp && sp.to > from && sp.to < L) {
            place({ k: "p", bi: bi, from: from, to: sp.to }, sp.h);
            if (from === 0 && heading) pageOf[b.id] = pages.length;
            newPage(); from = sp.to; continue;
          }
          if (y === 0) { place({ k: "p", bi: bi, from: from, to: L }, full); if (from === 0 && heading) pageOf[b.id] = pages.length; break; }
          newPage();
        }
      });
      // drop measurements of blocks that no longer exist
      Object.keys(cache).forEach(function (k) { if (!cache[k].used) delete cache[k]; });
      return { pages: pages, pageOf: pageOf, blockPage: blockPage, labels: labels, entries: entries.map(function (x) { x.page = pageOf[x.id] || 0; return x; }), g: g };
    }
    return { paginate: paginate, reset: function () { cache = {}; } };
  }
  window.DL = { Layout: Layout };
})();
