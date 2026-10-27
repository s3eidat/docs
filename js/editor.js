/* Docs: the editor. All pages live in ONE contenteditable host so the caret and selections cross paragraphs and
 * pages natively, but the document model is the source of truth: every edit arrives as a beforeinput event, is
 * applied to the model, and the affected pages are re-rendered from it. The only edits the browser performs itself
 * are IME compositions (and anything uncancelable); those are read back as plain text diffs, so no foreign markup
 * ever reaches the model. Header, footer, images, page breaks and the contents block are non-editable islands. */
(function () {
  "use strict";
  var M = window.DM, R = window.DR;

  function Editor(opts) {
    var stack = opts.stack, measurer = opts.measurer, layout = window.DL.Layout(measurer);
    var doc = null, plan = null, index = {}, els = {}, pageEls = [];
    var lastSel = null, atom = -1, pending = null, pendingKey = "", composing = false;
    var hist = { u: [], r: [], last: null, lastSel: null, t: 0, kind: "" };
    var api = {};

    stack.setAttribute("contenteditable", "true");
    stack.setAttribute("spellcheck", "true");
    stack.setAttribute("role", "textbox");
    stack.setAttribute("aria-multiline", "true");
    stack.setAttribute("aria-label", "Document");

    // ---------- rendering ----------
    function indexBlocks() { index = {}; doc.blocks.forEach(function (b, i) { index[b.id] = i; }); }
    function itemKey(it, b) { return b.id + ":" + it.k + ":" + (it.from == null ? "" : it.from + "-" + it.to) + (it.r0 == null ? "" : ":" + it.r0 + "-" + it.r1); }
    function itemEl(it, used) {
      var b = doc.blocks[it.bi], key = itemKey(it, b), g = plan.g, sig, el;
      if (it.k === "p") sig = JSON.stringify(b) + "|" + plan.labels[it.bi] + "|" + (doc.blocks.length === 1 ? "solo" : "");
      else if (it.k === "table") sig = JSON.stringify(b) + it.rep;
      else if (it.k === "img") sig = JSON.stringify(b) + (opts.blobUrl(b.blob) ? "1" : "0") + g.cw;
      else if (it.k === "toc") sig = JSON.stringify(plan.entries);
      else sig = b.id;
      var c = els[key];
      if (c && c.sig === sig) { used[key] = c; return c.el; }
      if (it.k === "p") {
        el = R.paraEl(doc, b, plan.labels[it.bi], it.from, it.to, { placeholder: doc.blocks.length === 1 && it.from === 0 ? "Start typing, or open a .docx file…" : "" });
        el.setAttribute("data-l", String(it.to - it.from));
      }
      else if (it.k === "table") el = R.tableEl(doc, b, it.r0, it.r1, it.rep);
      else if (it.k === "img") { el = R.imgEl(doc, b, opts.blobUrl(b.blob), g.cw, g.ch); if (!opts.blobUrl(b.blob)) opts.loadBlob(b.blob).then(function (u) { if (u) { delete els[key]; render(); } }); }
      else if (it.k === "toc") el = R.tocEl(doc, b, plan.entries);
      else el = R.pbEl(b);
      used[key] = { sig: sig, el: el };
      return el;
    }
    function pageEl(k) {
      var p = pageEls[k];
      if (!p) {
        p = { el: document.createElement("div"), hdr: document.createElement("div"), body: document.createElement("div"), ftr: document.createElement("div") };
        p.el.className = "dc-page"; p.hdr.className = "dc-hf dc-hdr"; p.ftr.className = "dc-hf dc-ftr"; p.body.className = "dc-body dc-content";
        p.hdr.setAttribute("contenteditable", "false"); p.ftr.setAttribute("contenteditable", "false");
        p.hdr.title = p.ftr.title = "Double-click to edit the header and footer";
        p.el.appendChild(p.hdr); p.el.appendChild(p.body); p.el.appendChild(p.ftr);
        pageEls[k] = p;
      }
      return p;
    }
    function render() {
      if (!doc) return;
      indexBlocks();
      plan = layout.paginate(doc);
      var g = plan.g, used = {}, n = plan.pages.length, date = new Date().toLocaleDateString();
      plan.pages.forEach(function (pg, k) {
        var p = pageEl(k), s = p.el.style;
        s.width = g.w + "px"; s.height = g.h + "px";
        s.setProperty("contain-intrinsic-size", "auto " + g.h + "px");
        var bs = p.body.style; bs.top = g.mt + "px"; bs.left = g.ml + "px"; bs.width = g.cw + "px"; bs.height = g.ch + "px";
        bs.fontFamily = R.fontStack(doc.base.font); bs.fontSize = doc.base.size + "pt"; bs.lineHeight = String(doc.base.line);
        var blank = doc.hf.first && k === 0, ctx = { page: k + 1, pages: n, title: doc.title, date: date };
        p.hdr.textContent = blank ? "" : M.fillTokens(doc.hf.header, ctx);
        p.ftr.textContent = blank ? "" : M.fillTokens(doc.hf.footer, ctx);
        p.hdr.style.cssText = "top:" + g.hd + "px;left:" + g.ml + "px;width:" + g.cw + "px;text-align:" + doc.hf.ha;
        p.ftr.style.cssText = "bottom:" + g.fd + "px;left:" + g.ml + "px;width:" + g.cw + "px;text-align:" + doc.hf.fa;
        p.el.setAttribute("data-page", String(k + 1));
        var kids = pg.items.map(function (it) { return itemEl(it, used); });
        var same = kids.length === p.body.children.length && kids.every(function (el, j) { return p.body.children[j] === el; });
        if (!same) { p.body.textContent = ""; kids.forEach(function (el) { p.body.appendChild(el); }); }
        if (p.el.parentNode !== stack || stack.children[k] !== p.el) stack.insertBefore(p.el, stack.children[k] || null);
      });
      while (pageEls.length > n) { var gone = pageEls.pop(); gone.el.remove(); }
      els = used;
      // contents page numbers can change without the block's size changing
      Array.prototype.forEach.call(stack.querySelectorAll(".dc-toc-e"), function (row) {
        var id = row.getAttribute("data-go"), n2 = row.querySelector(".dc-toc-n"), pgn = plan.pageOf[id];
        if (n2 && n2.textContent !== String(pgn || "")) n2.textContent = pgn ? String(pgn) : "";
      });
      if (opts.onRender) opts.onRender(plan);
    }

    // ---------- DOM <-> model positions ----------
    function holderEl(n) {
      var el = n && (n.nodeType === 1 ? n : n.parentNode);
      var h = el && el.closest ? el.closest("[data-o]") : null;
      return h && stack.contains(h) ? h : null;
    }
    function holderPos(h, o) {
      var bi = index[h.getAttribute("data-b")];
      if (bi == null) return null;
      var r = h.hasAttribute("data-r") ? +h.getAttribute("data-r") : -1, c = h.hasAttribute("data-c") ? +h.getAttribute("data-c") : -1;
      var p = M.P(bi, +h.getAttribute("data-o") + o, r, c);
      p.o = Math.max(0, Math.min(p.o, M.holderLen(doc, p)));
      return p;
    }
    function holdersIn(el) { return el.matches && el.matches("[data-o]") ? [el] : Array.prototype.slice.call(el.querySelectorAll ? el.querySelectorAll("[data-o]") : []); }
    function posFromDom(node, off) {
      if (!node || !stack.contains(node)) return null;
      var h = holderEl(node);
      if (h) return holderPos(h, R.localOffset(h, node, off));
      var el = node.nodeType === 1 ? node : node.parentNode;
      var at = el.closest && el.closest("[data-b]");
      if (at && stack.contains(at) && !at.hasAttribute("data-o")) {
        var bi = index[at.getAttribute("data-b")];
        if (bi != null && doc.blocks[bi].type !== "table" && doc.blocks[bi].type !== "p") return M.P(bi, 0);
      }
      if (node.nodeType === 1) {
        var kids = node.childNodes;
        for (var k = Math.min(off, kids.length - 1); k >= 0 && k < kids.length; k++) { var hs = kids[k].nodeType === 1 ? holdersIn(kids[k]) : []; if (hs.length) return holderPos(hs[0], 0); }
        for (k = Math.min(off, kids.length) - 1; k >= 0; k--) { hs = kids[k].nodeType === 1 ? holdersIn(kids[k]) : []; if (hs.length) { var last = hs[hs.length - 1]; return holderPos(last, R.walkLen(last)); } }
      }
      var pg = el.closest && el.closest(".dc-page");
      if (pg) { var all = holdersIn(pg); if (all.length) return holderPos(all[all.length - 1], R.walkLen(all[all.length - 1])); }
      return null;
    }
    function holderFor(p) {
      var b = doc.blocks[p.b];
      if (!b) return null;
      if (b.type === "table") return { el: stack.querySelector('.dc-cell[data-b="' + b.id + '"][data-r="' + p.r + '"][data-c="' + p.c + '"]'), base: 0 };
      if (b.type !== "p") return null;
      var frags = stack.querySelectorAll('p[data-b="' + b.id + '"]'), pick = null;
      for (var k = 0; k < frags.length; k++) { if (+frags[k].getAttribute("data-o") <= p.o) pick = frags[k]; }
      if (!pick) pick = frags[0];
      return pick ? { el: pick, base: +pick.getAttribute("data-o") } : null;
    }
    function domFromPos(p) {
      var hf = holderFor(p);
      if (hf && hf.el) return R.domPoint(hf.el, p.o - hf.base);
      var b = doc.blocks[p.b], el = b && stack.querySelector('[data-b="' + b.id + '"]');
      if (el && el.parentNode) { var i = Array.prototype.indexOf.call(el.parentNode.childNodes, el); return { node: el.parentNode, off: i + (p.o > 0 ? 1 : 0) }; }
      return null;
    }
    function readSel() {
      var s = window.getSelection();
      if (!s || !s.rangeCount || !stack.contains(s.anchorNode) || !stack.contains(s.focusNode)) return null;
      var a = posFromDom(s.anchorNode, s.anchorOffset), f = posFromDom(s.focusNode, s.focusOffset);
      if (!a || !f) return null;
      return mk(a, f);
    }
    function mk(a, f) { var o = M.order(a, f); return { a: a, f: f, start: o[0], end: o[1], collapsed: !M.cmp(a, f) }; }
    function sel() { var s = readSel(); if (s) lastSel = s; return s || lastSel || mk(M.startPos(doc), M.startPos(doc)); }
    function setSel(a, f, scroll) {
      f = f || a;
      lastSel = mk(a, f);
      var A = domFromPos(a), F = domFromPos(f), s = window.getSelection();
      if (!A || !F || !s) return;
      try { s.setBaseAndExtent(A.node, A.off, F.node, F.off); } catch (e) { /* node gone */ }
      if (scroll !== false) ensureVisible();
    }
    function ensureVisible() {
      var s = window.getSelection(), sc = opts.scroller;
      if (!s || !s.rangeCount || !sc) return;
      var r = s.getRangeAt(0).getBoundingClientRect();
      if (!r || (!r.height && !r.top)) { var n = s.focusNode && (s.focusNode.nodeType === 1 ? s.focusNode : s.focusNode.parentNode); if (n && n.getBoundingClientRect) r = n.getBoundingClientRect(); else return; }
      var box = sc.getBoundingClientRect();
      if (r.bottom > box.bottom - 30) sc.scrollTop += r.bottom - box.bottom + 60;
      else if (r.top < box.top + 10) sc.scrollTop -= box.top - r.top + 60;
    }

    // ---------- history ----------
    function snap() { return JSON.stringify({ title: doc.title, page: doc.page, hf: doc.hf, base: doc.base, blocks: doc.blocks }); }
    function commit(kind, s) {
      var now = Date.now(), cur = snap();
      if (cur === hist.last) return false;
      var merge = kind === "type" && hist.kind === "type" && now - hist.t < 1200;
      if (!merge) { hist.u.push({ s: hist.last, sel: hist.lastSel }); if (hist.u.length > 300) hist.u.shift(); }
      hist.r = []; hist.last = cur; hist.t = now; hist.kind = kind;
      hist.lastSel = s || lastSel;
      doc.updated = Date.now();
      if (opts.onChange) opts.onChange(kind);
      return true;
    }
    function restore(entry, into) {
      into.push({ s: snap(), sel: lastSel });
      var x = JSON.parse(entry.s);
      doc.title = x.title; doc.page = x.page; doc.hf = x.hf; doc.base = x.base; doc.blocks = x.blocks;
      hist.last = entry.s; hist.kind = ""; hist.lastSel = entry.sel;
      atom = -1;
      if (opts.onDocMeta) opts.onDocMeta();
      render();
      var p = entry.sel && entry.sel.start && doc.blocks[entry.sel.start.b] ? entry.sel.start : M.startPos(doc);
      stack.focus({ preventScroll: true });
      setSel(clampPos(p));
      doc.updated = Date.now();
      if (opts.onChange) opts.onChange("history");
    }
    function clampPos(p) {
      if (!doc.blocks[p.b]) return M.endPos(doc);
      var b = doc.blocks[p.b];
      if (b.type === "table" && (p.r < 0 || !b.rows[p.r] || !b.rows[p.r][p.c] || b.rows[p.r][p.c].m)) return M.firstIn(doc, p.b);
      if (b.type === "p" && p.r >= 0) return M.P(p.b, 0);
      return M.P(p.b, Math.min(p.o, M.holderLen(doc, p)), p.r, p.c);
    }
    function undo() { if (!hist.u.length || hist.u[hist.u.length - 1].s == null) { S3.toast("Nothing to undo"); return; } restore(hist.u.pop(), hist.r); }
    function redo() { if (!hist.r.length) { S3.toast("Nothing to redo"); return; } var e = hist.r.pop(); hist.u.push({ s: snap(), sel: lastSel }); var into = []; restore(e, into); }

    // one edit = mutate the model, record it, lay out again, put the caret back
    function edit(kind, fn) {
      var s = sel();
      var res = fn(s);
      if (res === false) return;
      var p = res && res.b != null ? res : res && res.caret ? res.caret : null;
      commit(kind, s);
      render();
      if (p) { stack.focus({ preventScroll: true }); setSel(clampPos(p), res.focus ? clampPos(res.focus) : null); }
      else if (res && res.keep) { stack.focus({ preventScroll: true }); setSel(clampPos(s.a), clampPos(s.f), false); }
      emitSel();
    }

    // ---------- typing ----------
    function typeMarks(p) {
      if (pending && pendingKey === JSON.stringify(p)) return pending;
      pending = null;
      return null;
    }
    function typeText(text) {
      edit("type", function (s) {
        var p = s.start, mk = typeMarks(s.start);
        if (atom >= 0) { var nb = M.para("normal"); doc.blocks.splice(atom + 1, 0, nb); p = M.P(atom + 1, 0); atom = -1; clearAtom(); }
        else if (!s.collapsed) p = M.deleteRange(doc, s.start, s.end);
        if (!M.holder(doc, p)) return false;
        pending = null;
        return M.insertTextAt(doc, p, text, mk);
      });
    }
    function del(dir, unit, targetRange) {
      edit("del", function (s) {
        if (atom >= 0) { var i = atom; atom = -1; clearAtom(); doc.blocks.splice(i, 1); M.normalize(doc); return M.firstIn(doc, Math.min(i, doc.blocks.length - 1)); }
        if (!s.collapsed) return M.deleteRange(doc, s.start, s.end);
        var p = s.start, h = M.holder(doc, p);
        if (!h) { var bi = p.b; doc.blocks.splice(bi, 1); M.normalize(doc); return M.firstIn(doc, Math.min(bi, doc.blocks.length - 1)); }
        var t = M.plain(h.text), L = t.length;
        if (dir < 0) {
          if (p.o === 0) return M.backAtStart(doc, p);
          var o2 = unit === "word" ? M.wordBack(t, p.o) : unit === "line" ? lineStart(p) : M.prevStop(t, p.o);
          if (targetRange && unit !== "char") { var tr = rangeFrom(targetRange); if (tr && tr.start.b === p.b && tr.start.r === p.r && tr.start.c === p.c && M.cmp(tr.end, p) <= 0) o2 = Math.min(o2, tr.start.o) === tr.start.o ? tr.start.o : o2; }
          return M.deleteRange(doc, M.P(p.b, o2, p.r, p.c), p);
        }
        if (p.o >= L) return M.deleteAtEnd(doc, p);
        var o3 = unit === "word" ? M.wordFwd(t, p.o) : unit === "line" ? L : M.nextStop(t, p.o);
        M.deleteRange(doc, p, M.P(p.b, o3, p.r, p.c));
        return p;
      });
    }
    function lineStart(p) {
      // start of the visual line the caret is on (for Ctrl/Cmd+Backspace "delete to line start")
      var s = window.getSelection();
      try { s.modify("extend", "backward", "lineboundary"); var q = readSel(); if (q && q.start.b === p.b && q.start.r === p.r) return q.start.o; } catch (e) { /* no modify */ }
      return 0;
    }
    function rangeFrom(sr) {
      var a = posFromDom(sr.startContainer, sr.startOffset), f = posFromDom(sr.endContainer, sr.endOffset);
      return a && f ? mk(a, f) : null;
    }
    function enter(soft) {
      edit("enter", function (s) {
        var p = s.start;
        if (atom >= 0) { var nb = M.para("normal"); doc.blocks.splice(atom + 1, 0, nb); var at = atom + 1; atom = -1; clearAtom(); return M.P(at, 0); }
        if (!s.collapsed) p = M.deleteRange(doc, s.start, s.end);
        if (soft) return M.insertTextAt(doc, p, "\n", typeMarks(s.start));
        return M.splitAt(doc, p);
      });
    }

    // read back what the browser did on its own (IME, uncancelable input) as a text diff
    function reconcile(h) {
      if (!h || !stack.contains(h)) return false;
      var p0 = holderPos(h, 0);
      if (!p0) return false;
      var holder = M.holder(doc, p0);
      if (!holder) return false;
      var base = +h.getAttribute("data-o"), oldLen = h.hasAttribute("data-l") ? +h.getAttribute("data-l") : M.len(holder.text) - base;
      var oldT = M.plain(M.slice(holder.text, base, base + oldLen)), newT = R.textOf(h);
      if (oldT === newT) return false;
      var a = 0, maxA = Math.min(oldT.length, newT.length);
      while (a < maxA && oldT[a] === newT[a]) a++;
      var z = 0;
      while (z < maxA - a && oldT[oldT.length - 1 - z] === newT[newT.length - 1 - z]) z++;
      var ins = newT.slice(a, newT.length - z), mk2 = M.marksAt(holder.text, base + a);
      if (!M.len(holder.text) && holder.m) mk2 = M.marksOf(holder.m);
      holder.text = M.cut(holder.text, base + a, base + oldT.length - z);
      if (ins) holder.text = M.insertText(holder.text, base + a, ins, mk2);
      if (holder.m && M.len(holder.text)) delete holder.m;
      return M.P(p0.b, base + a + ins.length, p0.r, p0.c);
    }
    var compHolder = null;
    stack.addEventListener("compositionstart", function () {
      composing = true;
      var s = window.getSelection();
      if (s && s.rangeCount && !s.isCollapsed) {
        var cur = readSel();
        if (cur) { M.deleteRange(doc, cur.start, cur.end); commit("type"); render(); setSel(cur.start); }
      }
      compHolder = holderEl(s && s.anchorNode);
    });
    stack.addEventListener("compositionend", function () {
      composing = false;
      var h = compHolder || holderEl(window.getSelection().anchorNode);
      compHolder = null;
      var p = reconcile(h);
      if (p) { commit("type"); render(); setSel(p); emitSel(); }
      else render();
    });
    stack.addEventListener("input", function (e) {
      if (composing || e.isComposing) return;
      var s = window.getSelection(), h = holderEl(s && s.anchorNode);
      var p = reconcile(h);
      if (p) { commit("type"); render(); setSel(p); emitSel(); } else render();
    });

    stack.addEventListener("beforeinput", function (e) {
      var t = e.inputType;
      if (t === "insertCompositionText" || e.isComposing || composing) return; // the IME owns these; read back on compositionend
      e.preventDefault();
      if (t === "insertText" || t === "insertReplacementText") {
        var data = e.data != null ? e.data : e.dataTransfer ? e.dataTransfer.getData("text/plain") : "";
        if (t === "insertReplacementText" && e.getTargetRanges && e.getTargetRanges().length) {
          var tr = rangeFrom(e.getTargetRanges()[0]);
          if (tr) { lastSel = tr; setSel(tr.start, tr.end, false); }
        }
        if (data) typeText(data);
        return;
      }
      if (t === "insertParagraph") return enter(false);
      if (t === "insertLineBreak") return enter(shiftEnter);
      var trs = e.getTargetRanges ? e.getTargetRanges() : [];
      if (t === "deleteContentBackward") return del(-1, "char");
      if (t === "deleteContentForward") return del(1, "char");
      if (t === "deleteWordBackward") return del(-1, "word", trs[0]);
      if (t === "deleteWordForward") return del(1, "word", trs[0]);
      if (t === "deleteSoftLineBackward" || t === "deleteHardLineBackward") return del(-1, "line");
      if (t === "deleteSoftLineForward" || t === "deleteHardLineForward") return del(1, "line");
      if (t === "deleteContent" || t === "deleteByCut") return del(-1, "char");
      if (t === "historyUndo") return undo();
      if (t === "historyRedo") return redo();
      if (t === "formatBold") return api.mark("b");
      if (t === "formatItalic") return api.mark("i");
      if (t === "formatUnderline") return api.mark("u");
      if (t === "formatStrikeThrough") return api.mark("s");
      if (t === "formatSuperscript") return api.mark("sup");
      if (t === "formatSubscript") return api.mark("sub");
      // insertFromPaste / insertFromDrop are handled by the paste / drop events; everything else is refused
    });

    // ---------- clipboard ----------
    function sliceBlocks(s) {
      var out = [];
      M.segments(doc, s.start, s.end).forEach(function (sg) {
        var b = doc.blocks[sg.b];
        if (sg.atom) { if (b.type === "img" || b.type === "pb") out.push(M.clone(b)); return; }
        if (sg.r >= 0) {
          var whole = (sg.b > s.start.b || s.start.r < 0) && (sg.b < s.end.b || s.end.r < 0);
          if (whole) { if (!out.length || out[out.length - 1].id !== b.id) out.push(M.clone(b)); return; }
          if (sg.to > sg.from) out.push(M.para("normal", M.slice(sg.h.text, sg.from, sg.to)));
          return;
        }
        var c = M.clone(b); c.text = M.slice(b.text, sg.from, sg.to); out.push(c);
      });
      return out.map(function (b) { var c = M.cleanBlock(b); c.id = M.uid(); return c; });
    }
    function onCopy(e, cut) {
      var s = readSel();
      var blocks = atom >= 0 ? [M.clone(doc.blocks[atom])] : s && !s.collapsed ? sliceBlocks(s) : null;
      if (!blocks || !blocks.length) return;
      e.preventDefault();
      var text = atom >= 0 ? "" : M.textBetween(doc, s.start, s.end);
      e.clipboardData.setData("text/plain", text);
      try { e.clipboardData.setData("text/html", window.DIO.blocksToHtml(blocks, { images: {} })); } catch (x) { /* html optional */ }
      try { e.clipboardData.setData("application/x-s3docs", JSON.stringify({ v: 1, blocks: blocks })); } catch (x) { /* custom type optional */ }
      if (cut) del(-1, "char");
    }
    stack.addEventListener("copy", function (e) { onCopy(e, false); });
    stack.addEventListener("cut", function (e) { onCopy(e, true); });
    stack.addEventListener("paste", function (e) {
      e.preventDefault();
      var dt = e.clipboardData;
      if (!dt) return;
      var files = Array.prototype.slice.call(dt.files || []).filter(function (f) { return /^image\//.test(f.type); });
      var own = "", html = "", text = "";
      try { own = dt.getData("application/x-s3docs"); } catch (x) { own = ""; }
      try { html = dt.getData("text/html"); } catch (x) { html = ""; }
      text = dt.getData("text/plain") || "";
      if (own) {
        try { var j = JSON.parse(own); if (j && Array.isArray(j.blocks)) return pasteBlocks(j.blocks.slice(0, 5000).map(function (b) { var c = M.cleanBlock(b); c.id = M.uid(); return c; })); } catch (x) { /* fall through */ }
      }
      if (files.length && !html && !text) return api.insertImages(files);
      if (e.shiftKey === true || !html) return pasteBlocks(window.DIO.textToBlocks(text));
      var res = window.DIO.htmlToBlocks(html);
      pasteBlocks(res.blocks);
      if (res.images.length) opts.adoptImages(res.images, function () { render(); });
    });
    function pasteBlocks(blocks) {
      if (!blocks || !blocks.length) return;
      edit("paste", function (s) {
        var p = s.start;
        if (atom >= 0) { p = M.P(atom, 1); atom = -1; clearAtom(); }
        else if (!s.collapsed) p = M.deleteRange(doc, s.start, s.end);
        if (doc.blocks[p.b] && doc.blocks[p.b].type !== "p" && doc.blocks[p.b].type !== "table") { var nb = M.para("normal"); doc.blocks.splice(p.b + 1, 0, nb); p = M.P(p.b + 1, 0); }
        return M.insertBlocks(doc, p, blocks);
      });
    }
    stack.addEventListener("dragstart", function (e) { if (!e.target.closest || !e.target.closest(".dc-handle")) e.preventDefault(); });
    stack.addEventListener("dragover", function (e) { if (e.dataTransfer && Array.prototype.indexOf.call(e.dataTransfer.types || [], "Files") >= 0) { e.preventDefault(); e.dataTransfer.dropEffect = "copy"; } });
    stack.addEventListener("drop", function (e) {
      e.preventDefault();
      var files = Array.prototype.slice.call((e.dataTransfer && e.dataTransfer.files) || []);
      var p = null;
      if (document.caretRangeFromPoint) { var r = document.caretRangeFromPoint(e.clientX, e.clientY); if (r) p = posFromDom(r.startContainer, r.startOffset); }
      else if (document.caretPositionFromPoint) { var cp = document.caretPositionFromPoint(e.clientX, e.clientY); if (cp) p = posFromDom(cp.offsetNode, cp.offset); }
      if (p) { lastSel = mk(p, p); }
      var imgs = files.filter(function (f) { return /^image\//.test(f.type); });
      if (imgs.length) api.insertImages(imgs);
      else if (files.length && opts.onFileDrop) opts.onFileDrop(files);
    });

    // ---------- keyboard ----------
    var isMac = /Mac|iPhone|iPad/.test(navigator.platform || "");
    var shiftEnter = false;
    stack.addEventListener("keydown", function (e) {
      shiftEnter = e.key === "Enter" && e.shiftKey;
      if (e.isComposing || composing) return;
      var mod = isMac ? e.metaKey : e.ctrlKey, k = e.key.toLowerCase(), code = e.code;
      if (atom >= 0) {
        if (e.key === "Backspace" || e.key === "Delete") { e.preventDefault(); return del(-1, "char"); }
        if (e.key === "Enter") { e.preventDefault(); return enter(false); }
        if (e.key === "Escape" || /^Arrow/.test(e.key)) {
          e.preventDefault();
          var i = atom; atom = -1; clearAtom();
          var back = e.key === "ArrowUp" || e.key === "ArrowLeft";
          var target = back ? (i > 0 ? M.lastIn(doc, i - 1) : M.P(i, 0)) : M.firstIn(doc, Math.min(i + 1, doc.blocks.length - 1));
          setSel(clampPos(target)); emitSel(); return;
        }
        if (mod && (k === "c" || k === "x" || k === "v" || k === "z" || k === "y")) { /* let clipboard / history through */ }
        else if (e.key.length === 1 && !mod) { /* typing replaces nothing; beforeinput adds a paragraph after */ }
      }
      if (mod && !e.altKey) {
        if (k === "z" && !e.shiftKey) { e.preventDefault(); return undo(); }
        if ((k === "z" && e.shiftKey) || k === "y") { e.preventDefault(); return redo(); }
        if (k === "b") { e.preventDefault(); return api.mark("b"); }
        if (k === "i") { e.preventDefault(); return api.mark("i"); }
        if (k === "u") { e.preventDefault(); return api.mark("u"); }
        if (k === "x" && e.shiftKey) { e.preventDefault(); return api.mark("s"); }
        if (code === "Period" && !e.shiftKey) { e.preventDefault(); return api.mark("sup"); }
        if (code === "Comma" && !e.shiftKey) { e.preventDefault(); return api.mark("sub"); }
        if (code === "Period" && e.shiftKey) { e.preventDefault(); return api.bumpSize(1); }
        if (code === "Comma" && e.shiftKey) { e.preventDefault(); return api.bumpSize(-1); }
        if (k === "a" && !e.shiftKey) { e.preventDefault(); return api.selectAll(); }
        if (e.key === "Enter") { e.preventDefault(); return api.pageBreak(); }
        if (e.shiftKey && (k === "l" || k === "e" || k === "r" || k === "j")) { e.preventDefault(); return api.align({ l: "left", e: "center", r: "right", j: "justify" }[k]); }
        if (!e.shiftKey && (k === "e" || k === "j")) { e.preventDefault(); return api.align(k === "e" ? "center" : "justify"); }
        if (e.shiftKey && code === "Digit7") { e.preventDefault(); return api.list("ol"); }
        if (e.shiftKey && code === "Digit8") { e.preventDefault(); return api.list("ul"); }
        if (e.shiftKey && code === "Digit9") { e.preventDefault(); return api.list("todo"); }
        if (k === "\\") { e.preventDefault(); return api.clearFormat(); }
        if (k === "]" ) { e.preventDefault(); return api.indent(1); }
        if (k === "[" ) { e.preventDefault(); return api.indent(-1); }
      }
      if (mod && e.altKey && /^Digit[0-6]$/.test(code)) {
        e.preventDefault();
        return api.style(["normal", "h1", "h2", "h3", "title", "subtitle", "quote"][+code.slice(5)]);
      }
      if (e.key === "Tab" && !mod && !e.altKey) {
        var s = sel();
        e.preventDefault();
        if (s.start.r >= 0) return api.tableTab(e.shiftKey);
        var b = doc.blocks[s.start.b];
        if (b && b.type === "p" && (b.list || e.shiftKey || (!s.collapsed && s.start.b !== s.end.b))) return api.indent(e.shiftKey ? -1 : 1);
        if (b && b.type === "p" && s.collapsed && s.start.o === 0 && M.len(b.text) && !b.fi) return edit("style", function () { b.fi = "first"; return s.start; });
        return typeText("\t");
      }
      if (e.key === "Backspace" && !mod && !e.altKey) {
        // Chrome sends no beforeinput when the caret sits at the start of an empty host, so handle the start-of-paragraph cases here
        var s2 = readSel();
        if (s2 && s2.collapsed && s2.start.o === 0) { e.preventDefault(); return del(-1, "char"); }
      }
      if (e.key === "Delete" && !mod) {
        var s3 = readSel();
        if (s3 && s3.collapsed && s3.start.o >= M.holderLen(doc, s3.start)) { e.preventDefault(); return del(1, "char"); }
      }
    });

    // ---------- pointer: atoms, checkboxes, links, contents, image + column resize ----------
    function clearAtom() { Array.prototype.forEach.call(stack.querySelectorAll(".dc-sel"), function (x) { x.classList.remove("dc-sel"); }); }
    function selectAtom(i) {
      atom = i; clearAtom();
      var el = stack.querySelector('[data-b="' + doc.blocks[i].id + '"]');
      if (el) el.classList.add("dc-sel");
      var s = window.getSelection(); if (s) s.removeAllRanges();
      stack.focus({ preventScroll: true });
      emitSel();
    }
    var drag = null;
    stack.addEventListener("pointerdown", function (e) {
      var t = e.target;
      if (t.closest(".dc-handle")) {
        var fig = t.closest(".dc-fig"), bi = index[fig.getAttribute("data-b")], box = fig.querySelector(".dc-imgbox");
        e.preventDefault();
        selectAtom(bi);
        drag = { kind: "img", bi: bi, x: e.clientX, w: box.offsetWidth, ratio: doc.blocks[bi].h / doc.blocks[bi].w, box: box, z: zoomNow(), rtl: getComputedStyle(fig).direction === "rtl" };
        stack.setPointerCapture(e.pointerId);
        return;
      }
      var td = t.closest("td,th");
      if (td && td.hasAttribute("data-c") && nearEdge(td, e)) {
        var tbl = td.closest("table"), tb = index[tbl.getAttribute("data-b")], c = +td.getAttribute("data-c") + (td.colSpan || 1) - 1;
        if (c < doc.blocks[tb].cw.length - 1) {
          e.preventDefault();
          drag = { kind: "col", bi: tb, c: c, x: e.clientX, w: tbl.getBoundingClientRect().width, cw: doc.blocks[tb].cw.slice(), tbl: tbl, rtl: getComputedStyle(tbl).direction === "rtl" };
          stack.setPointerCapture(e.pointerId);
          return;
        }
      }
      var fig2 = t.closest(".dc-fig,.dc-pbk");
      if (fig2) { e.preventDefault(); selectAtom(index[fig2.getAttribute("data-b")]); return; }
      var go = t.closest(".dc-toc-e");
      if (go) { e.preventDefault(); api.goTo(go.getAttribute("data-go")); return; }
      if (t.closest(".dc-toc")) { e.preventDefault(); selectAtom(index[t.closest(".dc-toc").getAttribute("data-b")]); return; }
      if (t.closest(".dc-hf")) { e.preventDefault(); return; }
      if (atom >= 0) { atom = -1; clearAtom(); }
      var p = t.closest && t.closest("p.dc-li-todo");
      if (p && t === p && onMarker(p, e)) {
        e.preventDefault();
        var bi2 = index[p.getAttribute("data-b")];
        edit("check", function () { var b = doc.blocks[bi2]; b.list.c = !b.list.c; return { keep: true }; });
        return;
      }
      var a = t.closest("a.dc-a");
      if (a && (e.ctrlKey || e.metaKey)) { e.preventDefault(); var u = M.safeUrl(a.getAttribute("href") || ""); if (u) window.open(u, "_blank", "noopener,noreferrer"); }
    });
    stack.addEventListener("click", function (e) { if (e.target.closest("a.dc-a")) e.preventDefault(); });
    stack.addEventListener("dblclick", function (e) { if (e.target.closest(".dc-hf") && opts.onHeaderFooter) opts.onHeaderFooter(); });
    function zoomNow() { return parseFloat(stack.style.zoom) || 1; }
    function nearEdge(td, e) {
      var r = td.getBoundingClientRect(), rtl = getComputedStyle(td).direction === "rtl";
      return rtl ? e.clientX - r.left < 5 : r.right - e.clientX < 5;
    }
    function onMarker(p, e) {
      var r = p.getBoundingClientRect(), z = zoomNow(), rtl = p.dir === "rtl";
      var mkx = parseFloat(p.style.getPropertyValue("--mk")) || 0, x = rtl ? (r.right - e.clientX) / z : (e.clientX - r.left) / z;
      return x >= mkx - 4 && x <= mkx + 24;
    }
    stack.addEventListener("pointermove", function (e) {
      if (!drag) {
        var td = e.target.closest && e.target.closest("td,th");
        stack.style.cursor = td && td.hasAttribute("data-c") && nearEdge(td, e) ? "col-resize" : "";
        return;
      }
      var z = drag.kind === "img" ? drag.z : zoomNow(), dx = (e.clientX - drag.x) / z;
      if (drag.kind === "img") {
        if (drag.rtl) dx = -dx;
        var w = Math.max(24, Math.min(plan.g.cw, drag.w + dx));
        drag.nw = Math.round(w);
        drag.box.style.width = drag.nw + "px"; drag.box.style.height = Math.round(w * drag.ratio) + "px";
        var im = drag.box.querySelector("img"); if (im) { im.width = drag.nw; im.height = Math.round(w * drag.ratio); }
      } else {
        if (drag.rtl) dx = -dx;
        var f = dx * zoomNow() / drag.w, cw = drag.cw.slice(), c = drag.c, tot = cw[c] + cw[c + 1];
        cw[c] = Math.max(0.04, Math.min(tot - 0.04, cw[c] + f)); cw[c + 1] = tot - cw[c];
        drag.ncw = cw;
        Array.prototype.forEach.call(drag.tbl.querySelectorAll("col"), function (col, k) { col.style.width = (cw[k] * 100).toFixed(3) + "%"; });
      }
    });
    stack.addEventListener("pointerup", function () {
      if (!drag) return;
      var d = drag; drag = null;
      if (d.kind === "img" && d.nw) {
        edit("resize", function () { var b = doc.blocks[d.bi]; b.w = d.nw; b.h = Math.max(4, Math.round(d.nw * d.ratio)); return { keep: true }; });
        selectAtom(d.bi);
      } else if (d.kind === "col" && d.ncw) edit("resize", function () { doc.blocks[d.bi].cw = d.ncw; return { keep: true }; });
    });

    // ---------- selection state for the toolbar ----------
    var selTimer = 0;
    function emitSel() {
      if (!opts.onSel || !doc) return;
      cancelAnimationFrame(selTimer);
      selTimer = requestAnimationFrame(function () { opts.onSel(api.state()); });
    }
    document.addEventListener("selectionchange", function () {
      if (!doc || composing) return;
      var s = readSel();
      if (s) {
        lastSel = s;
        if (pending && pendingKey !== JSON.stringify(s.start)) pending = null;
        // keep a collapsed caret inside a text holder (Ctrl+End, clicks in the margins can park it on the page itself)
        var ds = window.getSelection();
        if (s.collapsed && ds.isCollapsed && !holderEl(ds.anchorNode) && doc.blocks[s.start.b] && (doc.blocks[s.start.b].type === "p" || s.start.r >= 0)) setSel(s.start, null, false);
        emitSel();
      }
    });

    // ---------- commands ----------
    api.state = function () {
      var s = lastSel || mk(M.startPos(doc), M.startPos(doc));
      var b = doc.blocks[s.start.b] || doc.blocks[0], h = M.holder(doc, s.start);
      var marks = pending && pendingKey === JSON.stringify(s.start) ? pending : M.marksInRange(doc, s.start, s.end);
      var st = b.type === "p" ? M.styleOf(doc, b) : null;
      return {
        style: b.type === "p" ? b.style : "", marks: marks, list: b.type === "p" && b.list ? b.list.k : "",
        align: b.type === "p" ? toPhysical(b.align, R.dirOf(b) === "rtl") : h ? toPhysical(h.al || "", R.dirOf({ dir: "", text: h.text }) === "rtl") : "left",
        rtl: b.type === "p" ? R.dirOf(b) === "rtl" : false,
        dir: b.type === "p" ? b.dir : "", lh: st ? st.lh : doc.base.line, inTable: s.start.r >= 0 ? { b: s.start.b, r: s.start.r, c: s.start.c } : null,
        atom: atom >= 0 ? doc.blocks[atom].type : "", collapsed: s.collapsed, sel: s, size: marks.z || (st ? st.size : doc.base.size), font: marks.f || (st && st.font) || doc.base.font,
        page: pageOfPos(s.start)
      };
    };
    function pageOfPos(p) {
      var b = doc.blocks[p.b];
      if (!b || !plan) return 1;
      var el = (holderFor(p) || {}).el || stack.querySelector('[data-b="' + b.id + '"]'), pg = el && el.closest(".dc-page");
      return pg ? +pg.getAttribute("data-page") : 1;
    }
    api.mark = function (key, value) {
      edit("mark", function (s) {
        if (atom >= 0) return false;
        if (s.collapsed) {
          var cur = M.marksInRange(doc, s.start, s.start), on = value !== undefined ? value : !cur[key];
          pending = Object.assign({}, typeMarks(s.start) || cur);
          if (on != null && on !== false && on !== "") pending[key] = on === true && M.BOOL.indexOf(key) < 0 ? pending[key] : on; else delete pending[key];
          if (key === "sup" && on) delete pending.sub;
          if (key === "sub" && on) delete pending.sup;
          pendingKey = JSON.stringify(s.start);
          var h = M.holder(doc, s.start);
          if (h && !M.len(h.text)) { var mm = Object.assign({}, pending); delete mm.a; if (Object.keys(mm).length) h.m = mm; else delete h.m; }
          var keepPending = pending, keepKey = pendingKey;
          setTimeout(function () { pending = keepPending; pendingKey = keepKey; emitSel(); }, 0);
          return { keep: true };
        }
        if (value === undefined) M.toggleMarkRange(doc, s.start, s.end, key);
        else M.setMarkRange(doc, s.start, s.end, key, value);
        return { keep: true };
      });
    };
    api.bumpSize = function (dir) {
      var st = api.state(), cur = st.size || doc.base.size, steps = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 60, 72, 96];
      var nx = dir > 0 ? steps.filter(function (v) { return v > cur; })[0] || cur + 6 : steps.filter(function (v) { return v < cur; }).pop() || Math.max(1, cur - 1);
      api.mark("z", nx);
    };
    api.clearFormat = function () { edit("mark", function (s) { M.clearFormat(doc, s.start, s.end); M.setStyle(doc, s.start, s.end, "normal"); M.parasIn(doc, s.start, s.end).forEach(function (bi) { var b = doc.blocks[bi]; b.align = ""; b.lh = 0; b.sb = b.sa = null; delete b.m; }); return { keep: true }; }); };
    api.style = function (style) { edit("style", function (s) { M.setStyle(doc, s.start, s.end, style); return { keep: true }; }); };
    // the toolbar speaks physical left / right; the model stores logical start ("") / end ("right") like .docx does
    function toLogical(al, rtl) { if (al === "center" || al === "justify") return al; if (!rtl) return al === "left" ? "" : "right"; return al === "right" ? "" : "right"; }
    function toPhysical(al, rtl) { if (al === "center" || al === "justify") return al; if (!rtl) return al === "right" ? "right" : "left"; return al === "right" ? "left" : "right"; }
    api.align = function (al) {
      edit("style", function (s) {
        M.parasIn(doc, s.start, s.end).forEach(function (bi) { var b = doc.blocks[bi]; b.align = toLogical(al, R.dirOf(b) === "rtl"); });
        M.segments(doc, s.start, s.end).forEach(function (sg) { if (sg.r >= 0) sg.h.al = toLogical(al, R.dirOf({ dir: "", text: sg.h.text }) === "rtl"); });
        return { keep: true };
      });
    };
    api.list = function (kind) { edit("style", function (s) { M.toggleList(doc, s.start, s.end, kind); return { keep: true }; }); };
    api.indent = function (d) { edit("style", function (s) { M.indent(doc, s.start, s.end, d); return { keep: true }; }); };
    api.dir = function (dir) { edit("style", function (s) { M.parasIn(doc, s.start, s.end).forEach(function (bi) { doc.blocks[bi].dir = dir; if (dir && !doc.blocks[bi].align) doc.blocks[bi].align = ""; }); return { keep: true }; }); };
    api.paraProp = function (key, value) { edit("style", function (s) { M.parasIn(doc, s.start, s.end).forEach(function (bi) { doc.blocks[bi][key] = value; }); return { keep: true }; }); };
    api.link = function (url) {
      edit("link", function (s) {
        if (s.collapsed) {
          var h = M.holder(doc, s.start);
          if (!h) return false;
          // caret inside an existing link: change or remove the whole link
          var at = s.start.o, spans = h.text, pos = 0, hit = null;
          spans.forEach(function (sp) { if (!hit && sp.a && at >= pos && at <= pos + sp.t.length) hit = [pos, pos + sp.t.length]; pos += sp.t.length; });
          if (hit) { h.text = M.setMark(h.text, hit[0], hit[1], "a", url || null); return { keep: true }; }
          if (!url) return false;
          var label = url.replace(/^mailto:/, "");
          h.text = M.insertText(h.text, at, label, Object.assign(M.marksAt(h.text, at), { a: url }));
          return M.P(s.start.b, at + label.length, s.start.r, s.start.c);
        }
        M.setMarkRange(doc, s.start, s.end, "a", url || null);
        return { keep: true };
      });
    };
    api.linkAt = function () {
      var s = sel(), h = M.holder(doc, s.start);
      if (!h) return "";
      var part = M.slice(h.text, Math.max(0, s.start.o - (s.collapsed ? 1 : 0)), s.collapsed ? s.start.o + 1 : s.end.r === s.start.r && s.end.b === s.start.b ? s.end.o : M.len(h.text));
      var l = part.filter(function (x) { return x.a; })[0];
      return l ? l.a : "";
    };
    api.selectedText = function () { var s = sel(); return s.collapsed ? "" : M.textBetween(doc, s.start, s.end); };
    api.insertBlock = function (block, into) {
      edit("insert", function (s) {
        var p = s.start;
        if (atom >= 0) { p = M.P(atom, 1); atom = -1; clearAtom(); }
        else if (!s.collapsed) p = M.deleteRange(doc, s.start, s.end);
        var caret = M.insertBlockAt(doc, p, block);
        indexBlocks();
        if (into && block.type === "table") return M.P(index[block.id], 0, 0, 0);
        return caret;
      });
    };
    api.pageBreak = function () { api.insertBlock(M.pagebreak()); };
    api.insertTable = function (r, c) { api.insertBlock(M.table(r, c), true); };
    api.insertToc = function () {
      if (doc.blocks.some(function (b) { return b.type === "toc"; })) { S3.toast("This document already has a table of contents. It updates by itself."); return; }
      api.insertBlock(M.toc());
    };
    api.insertImages = function (files) {
      var list = Array.prototype.slice.call(files || []).slice(0, 20), at = sel();
      return list.reduce(function (pr, f) {
        return pr.then(function () {
          return opts.addImage(f).then(function (img) {
            if (!img) return;
            var g = M.geometry(doc), w = Math.min(img.w, g.cw), h = Math.round(img.h * (w / img.w));
            lastSel = at;
            api.insertBlock(M.image(img.id, Math.round(w), Math.max(4, h), { alt: (f.name || "").replace(/\.[a-z0-9]+$/i, "").slice(0, 120) }));
            at = lastSel;
          });
        });
      }, Promise.resolve()).catch(function (e) { S3.toast(e && e.message ? e.message : "Could not add that image", "error"); });
    };
    api.imageProp = function (key, value) {
      if (atom < 0 || doc.blocks[atom].type !== "img") return;
      var i = atom;
      edit("style", function () { var b = doc.blocks[i]; if (key === "scale") { var g = M.geometry(doc), w = Math.round(g.cw * value); b.h = Math.round(b.h * w / b.w); b.w = w; } else b[key] = value; return { keep: true }; });
      selectAtom(i);
    };
    api.atomBlock = function () { return atom >= 0 ? doc.blocks[atom] : null; };
    api.deleteAtom = function () { if (atom >= 0) del(-1, "char"); };
    api.table = function (op, arg) {
      edit("table", function (s) {
        var c = s.start;
        if (c.r < 0 || !doc.blocks[c.b] || doc.blocks[c.b].type !== "table") return false;
        var t = doc.blocks[c.b], cl = t.rows[c.r][c.c], R2 = c.r, C2 = c.c;
        if (op === "rowAbove") { M.addRow(t, c.r); return M.P(c.b, 0, c.r, c.c); }
        if (op === "rowBelow") { M.addRow(t, c.r + cl.rs); return M.P(c.b, 0, c.r + cl.rs, firstCol(t, c.r + cl.rs, c.c)); }
        if (op === "colLeft") { M.addCol(t, c.c); return M.P(c.b, 0, c.r, c.c); }
        if (op === "colRight") { M.addCol(t, c.c + cl.cs); return M.P(c.b, 0, c.r, c.c + cl.cs); }
        if (op === "delRow") { if (t.rows.length <= 1) return delTable(c.b); M.delRow(t, c.r); return M.firstIn(doc, c.b); }
        if (op === "delCol") { if (t.rows[0].length <= 1) return delTable(c.b); M.delCol(t, c.c); return M.firstIn(doc, c.b); }
        if (op === "delTable") return delTable(c.b);
        if (op === "header") { t.hdr = !t.hdr; return { keep: true }; }
        if (op === "shade") { var cells = M.segments(doc, s.start, s.end).filter(function (x) { return x.r >= 0 && x.b === c.b; }); (cells.length ? cells : [{ h: cl }]).forEach(function (x) { x.h.bg = arg || ""; }); return { keep: true }; }
        if (op === "merge") {
          if (s.end.b === c.b && s.end.r >= 0 && (s.end.r !== c.r || s.end.c !== c.c)) { R2 = s.end.r; C2 = s.end.c; }
          else if (c.c + cl.cs < t.rows[0].length) C2 = c.c + cl.cs;
          else return false;
          var r0 = Math.min(c.r, R2), c0 = Math.min(c.c, C2);
          M.mergeCells(t, c.r, c.c, R2 + (t.rows[R2][C2].rs - 1), C2 + (t.rows[R2][C2].cs - 1));
          return M.P(c.b, 0, r0, c0);
        }
        if (op === "split") { M.splitCell(t, c.r, c.c); return M.P(c.b, 0, c.r, c.c); }
        if (op === "even") { t.cw = t.cw.map(function () { return 1 / t.cw.length; }); return { keep: true }; }
        return false;
      });
    };
    function firstCol(t, r, c) { if (!t.rows[r]) return 0; while (c > 0 && t.rows[r][c].m) c--; return c; }
    function delTable(bi) { doc.blocks.splice(bi, 1); M.normalize(doc); return M.firstIn(doc, Math.min(bi, doc.blocks.length - 1)); }
    api.tableTab = function (back) {
      var s = sel(), c = s.start, t = doc.blocks[c.b], cells = [];
      t.rows.forEach(function (row, r) { row.forEach(function (cl, ci) { if (!cl.m) cells.push([r, ci]); }); });
      var k = cells.findIndex(function (x) { return x[0] === c.r && x[1] === c.c; }), n = k + (back ? -1 : 1);
      if (n < 0) return;
      if (n >= cells.length) { edit("table", function () { M.addRow(t, t.rows.length); return M.P(c.b, 0, t.rows.length - 1, 0); }); return; }
      var tgt = cells[n], L = M.len(t.rows[tgt[0]][tgt[1]].text);
      setSel(M.P(c.b, 0, tgt[0], tgt[1]), M.P(c.b, L, tgt[0], tgt[1])); emitSel();
    };
    api.selectAll = function () { atom = -1; clearAtom(); stack.focus({ preventScroll: true }); setSel(M.startPos(doc), M.endPos(doc), false); emitSel(); };
    api.goTo = function (id, focusIt) {
      indexBlocks();
      var bi = index[id];
      if (bi == null) return;
      var el = stack.querySelector('[data-b="' + id + '"]');
      if (el && opts.scroller) { var sc = opts.scroller, r = el.getBoundingClientRect(), box = sc.getBoundingClientRect(); sc.scrollTop += r.top - box.top - 40; }
      stack.focus({ preventScroll: true });
      setSel(M.firstIn(doc, bi), null, false); emitSel();
    };
    api.select = function (a, f) { atom = -1; clearAtom(); stack.focus({ preventScroll: true }); setSel(a, f); emitSel(); };
    api.rangeOf = function (a, f) {
      var A = domFromPos(a), F = domFromPos(f);
      if (!A || !F) return null;
      var r = document.createRange();
      try { r.setStart(A.node, A.off); r.setEnd(F.node, F.off); } catch (e) { return null; }
      return r;
    };
    api.change = function (kind, fn) { fn(doc); M.normalize(doc); commit(kind); render(); if (lastSel) { lastSel = mk(clampPos(lastSel.a), clampPos(lastSel.f)); } emitSel(); };
    api.relayout = function () { layout.reset(); els = {}; render(); };
    api.render = render;
    api.undo = undo; api.redo = redo;
    api.canUndo = function () { return hist.u.length > 0 && hist.u[hist.u.length - 1].s != null; };
    api.load = function (d) {
      doc = d; atom = -1; pending = null; lastSel = null;
      hist = { u: [], r: [], last: snap(), lastSel: null, t: 0, kind: "" };
      layout.reset(); els = {};
      pageEls.forEach(function (p) { p.el.remove(); }); pageEls = [];
      stack.textContent = "";
      render();
      lastSel = mk(M.startPos(doc), M.startPos(doc));
      emitSel();
    };
    api.focusStart = function () { stack.focus({ preventScroll: true }); setSel(M.startPos(doc), null, false); emitSel(); };
    api.focusEnd = function () { stack.focus({ preventScroll: true }); setSel(M.endPos(doc)); emitSel(); };
    api.doc = function () { return doc; };
    api.plan = function () { return plan; };
    api.sel = sel;
    api.posFromDom = posFromDom; api.domFromPos = domFromPos;
    // for the self-test: drive the editor the way the browser would
    api._type = typeText; api._enter = enter; api._del = del; api._setSel = function (a, f) { stack.focus({ preventScroll: true }); setSel(a, f, false); };
    return api;
  }
  window.DE = { Editor: Editor };
})();
