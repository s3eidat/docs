/* Docs: the app shell. Home (templates, recent documents, open), the editor chrome (menus, toolbar, contextual
 * table / image bar, outline, find & replace, status bar, zoom), dialogs, storage, import / export and the demo. */
(function () {
  "use strict";
  var M = window.DM, R = window.DR, DB = window.DDB, UI = window.DUI, h = S3.h;
  function $(id) { return document.getElementById(id); }
  var stack = $("stack"), canvas = $("canvas"), tools = $("tools"), ctxBar = $("ctx");
  var doc = null, persistent = true, saveTimer = 0, dirty = false, lastState = null;
  var zoom = { mode: S3.store("dc-zoom-mode") || "auto", value: +(S3.store("dc-zoom") || 1) };
  var urls = {}, IMG_MAX = 25e6;

  $("dc-dyn").textContent = R.css() + "\n::highlight(dc-find){background:#ffe58a;color:inherit}::highlight(dc-find-cur){background:#ff9e3d;color:#000}";

  // ======================= images (Blobs in IndexedDB) =======================
  function blobUrl(id) { return urls[id] || null; }
  function loadBlob(id) {
    if (!id) return Promise.resolve(null);
    if (urls[id]) return Promise.resolve(urls[id]);
    return DB.getBlob(id).then(function (rec) { if (!rec || !rec.data) return null; urls[id] = URL.createObjectURL(rec.data); return urls[id]; }).catch(function () { return null; });
  }
  function bitmapSize(blob) {
    return new Promise(function (res, rej) {
      var u = URL.createObjectURL(blob), im = new Image();
      im.onload = function () { res({ w: im.naturalWidth, h: im.naturalHeight, img: im, url: u }); };
      im.onerror = function () { URL.revokeObjectURL(u); rej(new Error("That image could not be read")); };
      im.src = u;
    });
  }
  function toPng(info) {
    var c = document.createElement("canvas"), k = Math.min(1, 4000 / info.w, 4000 / info.h);
    c.width = Math.max(1, Math.round(info.w * k)); c.height = Math.max(1, Math.round(info.h * k));
    c.getContext("2d").drawImage(info.img, 0, 0, c.width, c.height);
    return new Promise(function (res) { c.toBlob(function (b) { res(b); }, "image/png"); });
  }
  // store image bytes (any browser-readable format; anything but PNG/JPEG/GIF becomes PNG so .docx can carry it)
  function storeImage(blob, id) {
    if (blob.size > IMG_MAX) return Promise.reject(new Error("Images up to 25 MB, please"));
    return blob.slice(0, 16).arrayBuffer().then(function (head) {
      var u8 = new Uint8Array(head), type = window.DX.sniff(new Uint8Array(Array.prototype.slice.call(u8).concat(new Array(12).fill(0))));
      var avif = u8.length >= 12 && String.fromCharCode(u8[4], u8[5], u8[6], u8[7]) === "ftyp";
      if (!type && !avif) throw new Error("That file is not an image this app can use (PNG, JPEG, GIF, WebP, BMP, AVIF)");
      return bitmapSize(blob).then(function (info) {
        var keep = type === "png" || type === "jpg" || type === "gif";
        return (keep ? Promise.resolve(blob.type ? blob : new Blob([blob], { type: type === "jpg" ? "image/jpeg" : "image/" + type })) : toPng(info)).then(function (out) {
          URL.revokeObjectURL(info.url);
          id = id || M.uid();
          var rec = { id: id, data: out, w: info.w, h: info.h };
          return DB.putBlob(rec).then(function () { urls[id] = URL.createObjectURL(out); return { id: id, w: info.w, h: info.h }; });
        });
      });
    });
  }
  function addImage(file) { return storeImage(file); }
  function bytesOf(id) {
    return DB.getBlob(id).then(function (rec) {
      if (!rec || !rec.data) return null;
      return rec.data.arrayBuffer().then(function (ab) { var u8 = new Uint8Array(ab), t = window.DX.sniff(u8); return t === "png" || t === "jpg" || t === "gif" ? { data: u8, type: t } : null; });
    });
  }
  function dataUrlOf(id) { return DB.getBlob(id).then(function (rec) { return rec && rec.data ? S3.readFile(rec.data, "dataurl") : null; }).catch(function () { return null; }); }
  function adoptImages(list, done) {
    list.reduce(function (pr, x) {
      return pr.then(function () {
        return fetchDataUrl(x.src).then(function (blob) { return storeImage(blob); }).then(function (img) {
          var g = M.geometry(doc), w = Math.min(img.w, g.cw);
          x.block.blob = img.id; x.block.w = Math.round(w); x.block.h = Math.max(4, Math.round(img.h * w / img.w));
        }).catch(function () { x.block.blob = ""; });
      });
    }, Promise.resolve()).then(function () {
      doc.blocks = doc.blocks.filter(function (b) { return b.type !== "img" || b.blob; });
      M.normalize(doc); ed.change("paste-images", function () {}); if (done) done();
    });
  }
  function fetchDataUrl(src) {
    // decode a data: URL ourselves (no fetch, so the CSP's connect-src stays out of it)
    var m = /^data:([^;,]+);base64,(.*)$/i.exec(src);
    if (!m) return Promise.reject(new Error("bad image"));
    var bin = atob(m[2]), u8 = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
    return Promise.resolve(new Blob([u8], { type: m[1] }));
  }

  // ======================= the editor =======================
  var ed = window.DE.Editor({
    stack: stack, measurer: $("measure"), scroller: canvas,
    blobUrl: blobUrl, loadBlob: loadBlob, addImage: addImage, adoptImages: adoptImages,
    onChange: function () { scheduleSave(); },
    onSel: function (st) { lastState = st; syncTools(st); },
    onRender: function (plan) { onRendered(plan); },
    onHeaderFooter: function () { headerFooterDialog(); },
    onFileDrop: function (files) { openFile(files[0]); },
    onDocMeta: function () { $("doc-title").value = doc.title || ""; }
  });

  // ======================= saving =======================
  function setSaved(t) { $("saved").textContent = t; }
  function scheduleSave() {
    dirty = true; setSaved("Saving…");
    clearTimeout(saveTimer);
    saveTimer = setTimeout(saveNow, 450);
  }
  function saveNow() {
    clearTimeout(saveTimer);
    if (!doc || !dirty) return Promise.resolve();
    dirty = false;
    doc.updated = Date.now();
    return DB.putDocs([doc]).then(function () { setSaved(persistent ? "Saved on this device" : "Not saved: browser storage is blocked"); S3.store("dc-last", doc.id); })
      .catch(function () { setSaved("Could not save. Download a .docx to keep your work."); });
  }
  document.addEventListener("visibilitychange", function () { if (document.visibilityState === "hidden") saveNow(); });
  window.addEventListener("pagehide", function () { saveNow(); });

  // ======================= toolbar =======================
  var ICON = {
    undo: '<path d="M9 14H5V10"/><path d="M5 14a7 7 0 1 1 2 4.9"/>', redo: '<path d="M15 14h4v-4"/><path d="M19 14a7 7 0 1 0-2 4.9"/>',
    print: '<path d="M7 9V4h10v5"/><rect x="4" y="9" width="16" height="7" rx="2"/><path d="M7 14h10v6H7z"/>',
    link: '<path d="M10 14a4 4 0 0 0 5.7 0l3-3a4 4 0 0 0-5.7-5.7l-1 1"/><path d="M14 10a4 4 0 0 0-5.7 0l-3 3a4 4 0 0 0 5.7 5.7l1-1"/>',
    image: '<rect x="4" y="5" width="16" height="14" rx="2"/><circle cx="9.5" cy="10" r="1.6"/><path d="M5 17l4.5-4.5 3 3 2.5-2.5L20 17"/>',
    table: '<rect x="4" y="5" width="16" height="14" rx="1.5"/><path d="M4 10h16M4 14.5h16M10 5v14"/>',
    left: '<path d="M4 6h16M4 10h10M4 14h16M4 18h10"/>', center: '<path d="M4 6h16M7 10h10M4 14h16M7 18h10"/>', right: '<path d="M4 6h16M10 10h10M4 14h16M10 18h10"/>', justify: '<path d="M4 6h16M4 10h16M4 14h16M4 18h16"/>',
    ul: '<circle cx="5.5" cy="7" r="1.3" fill="currentColor"/><circle cx="5.5" cy="12" r="1.3" fill="currentColor"/><circle cx="5.5" cy="17" r="1.3" fill="currentColor"/><path d="M10 7h10M10 12h10M10 17h10"/>',
    ol: '<path d="M4.5 5.5h1.5v4M4.5 9.5h3M4.3 14.2c.4-.7 2.6-.9 2.6.5 0 1-2.8 1.8-2.8 3.3h3"/><path d="M11 7h9M11 12h9M11 17h9"/>',
    todo: '<rect x="4" y="5" width="5" height="5" rx="1"/><path d="M5 15.5l1.4 1.4L9 14"/><path d="M12 7.5h8M12 15.5h8"/>',
    outdent: '<path d="M4 6h16M11 10h9M11 14h9M4 18h16"/><path d="M8 9.5L5 12l3 2.5"/>', indent: '<path d="M4 6h16M11 10h9M11 14h9M4 18h16"/><path d="M5 9.5L8 12l-3 2.5"/>',
    spacing: '<path d="M11 6h9M11 12h9M11 18h9"/><path d="M6 4v16M3.5 6.5L6 4l2.5 2.5M3.5 17.5L6 20l2.5-2.5"/>',
    clear: '<path d="M6 6h11M11 6l-3 12"/><path d="M14 14l5 5M19 14l-5 5"/>', pb: '<path d="M6 4v5h12V4M6 20v-5h12v5"/><path d="M4 12h3M10 12h4M17 12h3"/>',
    marker: '<path d="M14.5 4.5l5 5-8 8H7l-1.5 1.5H3l2.5-2.5V12z"/>', rtl: '<path d="M10 5h8M14 5v14M10 5a3 3 0 0 0 0 6h4"/><path d="M8 17H3m2-2l-2 2 2 2"/>', ltr: '<path d="M8 5h8M12 5v14M8 5a3 3 0 0 0 0 6h4"/><path d="M16 17h5m-2-2l2 2-2 2"/>',
    find: '<circle cx="11" cy="11" r="6"/><path d="M20 20l-4.5-4.5"/>', outline: '<path d="M4 6h3M4 12h3M4 18h3M10 6h10M12 12h8M12 18h8"/>'
  };
  function ic(name) { return '<svg class="dc-svg" viewBox="0 0 24 24" aria-hidden="true">' + ICON[name] + "</svg>"; }
  function tb(id, label, inner, extra) {
    var b = h("button", Object.assign({ type: "button", "class": "dc-tb", "data-cmd": id, title: label, "aria-label": label }, extra || {}));
    if (inner && inner.charAt(0) === "<") b.innerHTML = inner; else b.textContent = inner;
    return b;
  }
  function sep() { return h("span", { "class": "dc-tsep", "aria-hidden": "true" }); }
  var FONT_LIST = ["Calibri", "Cambria", "Aptos", "Arial", "Helvetica", "Inter", "Segoe UI", "Verdana", "Tahoma", "Trebuchet MS", "Century Gothic", "Gill Sans", "Times New Roman", "Georgia", "Garamond", "Palatino Linotype", "Book Antiqua", "Baskerville", "Constantia", "Courier New", "Consolas", "Menlo", "Roboto", "Open Sans", "Noto Sans", "Noto Serif", "Liberation Sans", "Liberation Serif", "DejaVu Sans", "Carlito", "Caladea", "Traditional Arabic", "Simplified Arabic", "Arabic Typesetting", "Sakkal Majalla", "Noto Naskh Arabic", "Noto Sans Arabic", "Amiri", "Cairo", "Dubai", "Geeza Pro", "Al Nile", "Baghdad", "Damascus"];
  var fontsOk = null;
  function deviceFonts() {
    if (fontsOk) return fontsOk;
    var c = document.createElement("canvas").getContext("2d"), probe = "mmmmmmmmmmlli WwQq@ ابجد هوز", out = [];
    function wd(f) { c.font = "48px " + f; return c.measureText(probe).width; }
    var bases = ["monospace", "serif", "sans-serif"].map(function (b) { return [b, wd(b)]; });
    FONT_LIST.forEach(function (f) {
      if (f === "Inter") { out.push(f); return; }
      if (bases.some(function (b) { return Math.abs(wd('"' + f + '", ' + b[0]) - b[1]) > 0.5; })) out.push(f);
    });
    fontsOk = out;
    return out;
  }
  var styleSel, fontSel, sizeIn;
  function buildTools() {
    styleSel = h("select", { "class": "dc-sel dc-sel-style", "aria-label": "Paragraph style", title: "Paragraph style (Ctrl+Alt+0–3)" });
    M.STYLE_KEYS.forEach(function (k) { styleSel.appendChild(h("option", { value: k, text: M.STYLES[k].name })); });
    fontSel = h("select", { "class": "dc-sel dc-sel-font", "aria-label": "Font", title: "Font (fonts installed on this device)" });
    sizeIn = h("input", { "class": "dc-size", "aria-label": "Font size (pt)", title: "Font size in points", inputmode: "decimal", maxlength: "5", autocomplete: "off" });
    var colorBtn = tb("color", "Text colour", '<span class="dc-a-glyph">A</span><span class="dc-bar-c" id="c-bar"></span>');
    var hiBtn = tb("hilite", "Highlight colour", ic("marker") + '<span class="dc-bar-c" id="h-bar"></span>');
    [
      tb("undo", "Undo (Ctrl+Z)", ic("undo")), tb("redo", "Redo (Ctrl+Y)", ic("redo")), sep(),
      styleSel, sep(), fontSel, sep(),
      tb("size-", "Smaller (Ctrl+Shift+,)", "−", { "class": "dc-tb dc-tb-sm" }), sizeIn, tb("size+", "Bigger (Ctrl+Shift+.)", "+", { "class": "dc-tb dc-tb-sm" }), sep(),
      tb("b", "Bold (Ctrl+B)", "B", { "class": "dc-tb dc-tb-b" }), tb("i", "Italic (Ctrl+I)", "I", { "class": "dc-tb dc-tb-i" }), tb("u", "Underline (Ctrl+U)", "U", { "class": "dc-tb dc-tb-u" }), tb("s", "Strikethrough (Ctrl+Shift+X)", "S", { "class": "dc-tb dc-tb-s" }),
      colorBtn, hiBtn, sep(),
      tb("link", "Link (Ctrl+K)", ic("link")), tb("image", "Insert image", ic("image")), tb("table", "Insert table", ic("table")), sep(),
      tb("al-left", "Align left (Ctrl+Shift+L)", ic("left")), tb("al-center", "Centre (Ctrl+Shift+E)", ic("center")), tb("al-right", "Align right (Ctrl+Shift+R)", ic("right")), tb("al-justify", "Justify (Ctrl+Shift+J)", ic("justify")),
      tb("spacing", "Line and paragraph spacing", ic("spacing")), sep(),
      tb("ul", "Bulleted list (Ctrl+Shift+8)", ic("ul")), tb("ol", "Numbered list (Ctrl+Shift+7)", ic("ol")), tb("todo", "Checklist (Ctrl+Shift+9)", ic("todo")),
      tb("outdent", "Decrease indent (Shift+Tab)", ic("outdent")), tb("indent", "Increase indent (Tab)", ic("indent")), sep(),
      tb("ltr", "Left-to-right paragraph", ic("ltr")), tb("rtl", "Right-to-left paragraph", ic("rtl")), sep(),
      tb("pb", "Page break (Ctrl+Enter)", ic("pb")), tb("clear", "Clear formatting (Ctrl+\\)", ic("clear")), sep(),
      tb("find", "Find and replace (Ctrl+F)", ic("find")), tb("outline", "Outline", ic("outline"))
    ].forEach(function (x) { tools.appendChild(x); });
    styleSel.addEventListener("change", function () { ed.style(styleSel.value); });
    fontSel.addEventListener("change", function () {
      if (fontSel.value === "__more") { syncTools(lastState); return otherFont(fontSel); }
      ed.mark("f", fontSel.value === doc.base.font ? null : fontSel.value);
    });
    sizeIn.addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); applySize(); } if (e.key === "Escape") { syncTools(lastState); ed.select(lastState.sel.a, lastState.sel.f); } });
    sizeIn.addEventListener("change", applySize);
    tools.addEventListener("pointerdown", function (e) { if (e.target.closest("button")) e.preventDefault(); });
    tools.addEventListener("click", function (e) { var b = e.target.closest("[data-cmd]"); if (b) cmd(b.getAttribute("data-cmd"), b); });
  }
  function applySize() {
    var v = parseFloat(String(sizeIn.value).replace(",", "."));
    if (!isFinite(v) || v < 1 || v > 400) { S3.toast("Font size from 1 to 400 pt", "error"); syncTools(lastState); return; }
    v = Math.round(v * 2) / 2;
    ed.mark("z", v);
  }
  function fillFonts(cur) {
    var list = deviceFonts().slice(), base = doc ? doc.base.font : "Calibri";
    [base, cur].forEach(function (f) { if (f && list.indexOf(f) < 0) list.unshift(f); });
    var key = list.join("|");
    if (fontSel.getAttribute("data-k") !== key) {
      fontSel.textContent = "";
      list.forEach(function (f) { var o = h("option", { value: f, text: f }); o.style.fontFamily = R.fontStack(f); fontSel.appendChild(o); });
      fontSel.appendChild(h("option", { value: "__more", text: "Other font…" }));
      fontSel.setAttribute("data-k", key);
    }
    fontSel.value = cur || base;
  }
  function otherFont(anchor) {
    var inp = h("input", { "class": "input", placeholder: "Font name, e.g. Gill Sans", "aria-label": "Font name", maxlength: "48" });
    var go = h("button", { type: "button", "class": "btn small primary", text: "Use font" });
    var box = h("div", { "class": "dc-pop stack" }, [h("label", { "class": "field" }, [h("span", { text: "A font installed on this device" }), inp]), h("div", { "class": "row" }, [go])]);
    function apply() {
      var v = inp.value.trim();
      if (!M.FONT_RE.test(v)) { S3.toast("Letters, numbers, spaces, dots and dashes only", "error"); return; }
      UI.close(); ed.mark("f", v);
    }
    go.addEventListener("click", apply);
    inp.addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); apply(); } });
    UI.popover(anchor, box); inp.focus();
  }
  function syncTools(st) {
    if (!st || !doc) return;
    var mk = st.marks || {};
    Array.prototype.forEach.call(tools.querySelectorAll("[data-cmd]"), function (b) {
      var c = b.getAttribute("data-cmd"), on = null;
      if (c === "b" || c === "i" || c === "u" || c === "s") on = !!mk[c];
      else if (/^al-/.test(c)) on = st.align === c.slice(3) || (!st.align && c === "al-left");
      else if (c === "ul" || c === "ol" || c === "todo") on = st.list === c;
      else if (c === "rtl") on = st.rtl; else if (c === "ltr") on = !st.rtl;
      else if (c === "outline") on = !$("outline").hidden;
      else if (c === "find") on = !$("find").hidden;
      if (on != null) b.setAttribute("aria-pressed", String(on));
    });
    if (document.activeElement !== styleSel) styleSel.value = st.style || "normal";
    styleSel.disabled = !st.style;
    if (document.activeElement !== fontSel) fillFonts(st.font);
    if (document.activeElement !== sizeIn) sizeIn.value = String(st.size || doc.base.size);
    $("c-bar").style.background = mk.c || "var(--text)";
    $("h-bar").style.background = mk.h || "transparent";
    $("st-page").textContent = "Page " + (st.page || 1) + " of " + ((ed.plan() && ed.plan().pages.length) || 1);
    syncCtx(st);
    markOutline(st);
    if (!st.collapsed) { var t = ed.selectedText(); if (t) { var c2 = M.countText(t); $("st-words").textContent = c2.words.toLocaleString() + " of " + wordsTotal.toLocaleString() + " words"; } }
    else $("st-words").textContent = wordsTotal.toLocaleString() + " word" + (wordsTotal === 1 ? "" : "s");
  }
  function cmd(c, anchor) {
    switch (c) {
      case "undo": return ed.undo();
      case "redo": return ed.redo();
      case "print": return printDoc();
      case "size-": return ed.bumpSize(-1);
      case "size+": return ed.bumpSize(1);
      case "b": case "i": case "u": case "s": return ed.mark(c);
      case "color": return colorPop(anchor, "c");
      case "hilite": return colorPop(anchor, "h");
      case "link": return linkPop(anchor);
      case "image": return $("img-in").click();
      case "table": return tablePop(anchor);
      case "al-left": case "al-center": case "al-right": case "al-justify": return ed.align(c.slice(3));
      case "spacing": return spacingMenu(anchor);
      case "ul": case "ol": case "todo": return ed.list(c);
      case "outdent": return ed.indent(-1);
      case "indent": return ed.indent(1);
      case "rtl": return ed.dir("rtl");
      case "ltr": return ed.dir("ltr");
      case "pb": return ed.pageBreak();
      case "clear": return ed.clearFormat();
      case "find": return toggleFind();
      case "outline": return toggleOutline();
    }
  }
  var TEXT_COLORS = ["#000000", "#434343", "#666666", "#999999", "#cccccc", "#ffffff", "#c0392b", "#e67e22", "#d4ac0d", "#27ae60", "#16a085", "#2471a3", "#1d3557", "#7d3c98", "#c2185b",
    "#7b241c", "#a04000", "#7d6608", "#145a32", "#0e6251", "#154360", "#17233a", "#4a235a", "#880e4f", "#f5b7b1", "#fad7a0", "#f9e79f", "#abebc6", "#a3e4d7", "#aed6f1"];
  var HILITES = ["yellow", "green", "cyan", "magenta", "red", "blue", "lightGray", "darkYellow", "darkGreen", "darkCyan", "darkMagenta", "darkRed", "darkBlue", "darkGray", "black"];
  function colorPop(anchor, key, onPick) {
    var grid = h("div", { "class": "dc-swatches", role: "listbox", "aria-label": key === "h" ? "Highlight colours" : "Colours" });
    var list = key === "h" ? HILITES.map(function (n) { return M.HL[n]; }) : key === "bg" ? ["#f1f4f8", "#eef1f5", "#dbe7f3", "#e3f2e8", "#fdf2d0", "#fbe1dc", "#ece4f6", "#d9d9d9", "#fff2cc", "#cfe2f3", "#d9ead3", "#f4cccc", "#1d3557", "#2471a3", "#27ae60"] : TEXT_COLORS;
    function pick(v) { UI.close(); if (onPick) onPick(v); else ed.mark(key, v); }
    list.forEach(function (c) {
      var b = h("button", { type: "button", "class": "dc-sw", title: c, "aria-label": c });
      b.style.background = c; b.addEventListener("click", function () { pick(c); }); grid.appendChild(b);
    });
    var none = h("button", { type: "button", "class": "btn small", text: key === "c" ? "Automatic" : "None" });
    none.addEventListener("click", function () { pick(null); });
    var kids = [grid, h("div", { "class": "row" }, [none])];
    if (key !== "h") {
      var custom = h("input", { type: "color", "class": "dc-colorin", "aria-label": "Custom colour", value: "#2471a3" });
      custom.addEventListener("change", function () { pick(custom.value.toLowerCase()); });
      kids[1].appendChild(h("label", { "class": "row dc-custom" }, [h("span", { "class": "faint", text: "Custom" }), custom]));
    }
    UI.popover(anchor, h("div", { "class": "dc-pop" }, kids));
  }
  function linkPop(anchor) {
    var cur = ed.linkAt();
    var inp = h("input", { "class": "input", placeholder: "Web address or email", "aria-label": "Link address", value: cur, spellcheck: "false", autocomplete: "off" });
    var msg = h("p", { "class": "faint dc-small", text: "Web (http, https) and email links only." });
    var ok = h("button", { type: "button", "class": "btn small primary", text: cur ? "Update" : "Apply" });
    var rm = h("button", { type: "button", "class": "btn small", text: "Remove link" });
    function apply() { var u = M.safeUrl(inp.value); if (!u) { msg.textContent = "That address isn't allowed. Use a web address (starting with https) or an email address."; msg.classList.add("dc-err"); return; } UI.close(); ed.link(u); }
    ok.addEventListener("click", apply); rm.addEventListener("click", function () { UI.close(); ed.link(null); });
    inp.addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); apply(); } });
    var row = h("div", { "class": "row" }, [ok]); if (cur) row.appendChild(rm);
    var at = anchor || caretRect();
    UI.popover(at, h("div", { "class": "dc-pop stack dc-linkpop" }, [inp, msg, row]));
    inp.focus(); inp.select();
  }
  function caretRect() {
    var s = window.getSelection();
    if (s && s.rangeCount) { var r = s.getRangeAt(0).getBoundingClientRect(); if (r && (r.width || r.height || r.top)) return r; }
    return tools.getBoundingClientRect();
  }
  function tablePop(anchor) {
    var R0 = 8, C0 = 10, grid = h("div", { "class": "dc-tgrid" }), lab = h("div", { "class": "faint dc-small", text: "Choose a size" }), cells = [];
    for (var r = 0; r < R0; r++) for (var c = 0; c < C0; c++) {
      (function (r, c) {
        var b = h("button", { type: "button", "class": "dc-tcell", "aria-label": (r + 1) + " by " + (c + 1) + " table" });
        b.addEventListener("mouseenter", function () { hl(r, c); }); b.addEventListener("focus", function () { hl(r, c); });
        b.addEventListener("click", function () { UI.close(); ed.insertTable(r + 1, c + 1); });
        cells.push([b, r, c]); grid.appendChild(b);
      })(r, c);
    }
    function hl(r, c) { cells.forEach(function (x) { x[0].classList.toggle("on", x[1] <= r && x[2] <= c); }); lab.textContent = (r + 1) + " × " + (c + 1); }
    UI.popover(anchor, h("div", { "class": "dc-pop" }, [grid, lab]));
  }
  function spacingMenu(anchor) {
    var st = lastState || {}, items = [{ section: "Line spacing" }];
    [[1, "Single"], [1.15, "1.15"], [1.5, "1.5"], [2, "Double"], [2.5, "2.5"], [3, "Triple"]].forEach(function (x) {
      items.push({ label: x[1], checked: Math.abs((st.lh || 0) - x[0]) < 0.01, onClick: function () { ed.paraProp("lh", x[0]); } });
    });
    items.push({ sep: true });
    items.push({ label: "Add space before paragraph", onClick: function () { ed.paraProp("sb", 12); } });
    items.push({ label: "Remove space before", onClick: function () { ed.paraProp("sb", 0); } });
    items.push({ label: "Add space after paragraph", onClick: function () { ed.paraProp("sa", 12); } });
    items.push({ label: "Remove space after", onClick: function () { ed.paraProp("sa", 0); } });
    items.push({ sep: true });
    items.push({ label: "First-line indent", checked: false, onClick: function () { ed.paraProp("fi", "first"); } });
    items.push({ label: "Hanging indent", onClick: function () { ed.paraProp("fi", "hang"); } });
    items.push({ label: "No special indent", onClick: function () { ed.paraProp("fi", ""); } });
    items.push({ sep: true });
    items.push({ label: "Custom spacing…", onClick: customSpacing });
    UI.menu(anchor, items);
  }
  function numField(label, val, attrs) {
    var inp = h("input", Object.assign({ "class": "input", inputmode: "decimal", value: String(val), autocomplete: "off" }, attrs || {}));
    return { el: h("label", { "class": "field" }, [h("span", { text: label }), inp]), inp: inp, val: function (lo, hi, def) { var v = parseFloat(String(inp.value).replace(",", ".")); return isFinite(v) ? Math.max(lo, Math.min(hi, v)) : def; } };
  }
  function customSpacing() {
    var b = doc.blocks[(lastState && lastState.sel.start.b) || 0], st = b && b.type === "p" ? M.styleOf(doc, b) : { sb: 0, sa: doc.base.after, lh: doc.base.line };
    var fb = numField("Space before (pt)", st.sb), fa = numField("Space after (pt)", st.sa), fl = numField("Line spacing (multiple)", st.lh);
    var go = h("button", { type: "button", "class": "btn primary", text: "Apply" });
    var dlg = UI.modal("Paragraph spacing", h("div", { "class": "stack" }, [h("div", { "class": "grid-3" }, [fb.el, fa.el, fl.el]), h("div", { "class": "row" }, [go])]));
    go.addEventListener("click", function () { dlg.close(); ed.paraProp("sb", fb.val(0, 200, 0)); ed.paraProp("sa", fa.val(0, 200, 8)); ed.paraProp("lh", fl.val(0.5, 5, 1.15)); });
  }

  // ---------- contextual bar: table / image ----------
  function syncCtx(st) {
    var mode = st.atom === "img" ? "img" : st.inTable ? "table" : "";
    if (ctxBar.getAttribute("data-mode") === mode) { if (mode === "table") syncTableCtx(st); return; }
    ctxBar.setAttribute("data-mode", mode);
    ctxBar.textContent = "";
    ctxBar.hidden = !mode;
    function b(label, fn, title) { var x = h("button", { type: "button", "class": "btn small", text: label, title: title || label }); x.addEventListener("pointerdown", function (e) { e.preventDefault(); }); x.addEventListener("click", function () { fn(x); }); ctxBar.appendChild(x); return x; }
    if (mode === "table") {
      ctxBar.appendChild(h("span", { "class": "dc-ctx-l", text: "Table" }));
      b("+ Row above", function () { ed.table("rowAbove"); }); b("+ Row below", function () { ed.table("rowBelow"); });
      b("+ Column left", function () { ed.table("colLeft"); }); b("+ Column right", function () { ed.table("colRight"); });
      b("− Row", function () { ed.table("delRow"); }, "Delete row"); b("− Column", function () { ed.table("delCol"); }, "Delete column");
      b("Merge", function () { ed.table("merge"); }, "Merge the selected cells (or with the cell to the right)");
      b("Split", function () { ed.table("split"); }, "Split merged cell");
      b("Shading", function (x) { colorPop(x, "bg", function (v) { ed.table("shade", v); }); }, "Cell shading");
      var hd = b("Header row", function () { ed.table("header"); }); hd.setAttribute("data-k", "hdr");
      b("Even columns", function () { ed.table("even"); });
      b("Delete table", function () { ed.table("delTable"); });
      syncTableCtx(st);
    } else if (mode === "img") {
      ctxBar.appendChild(h("span", { "class": "dc-ctx-l", text: "Image" }));
      b("Left", function () { ed.imageProp("align", "left"); }); b("Centre", function () { ed.imageProp("align", "center"); }); b("Right", function () { ed.imageProp("align", "right"); });
      [[0.25, "25%"], [0.5, "50%"], [0.75, "75%"], [1, "Full width"]].forEach(function (x) { b(x[1], function () { ed.imageProp("scale", x[0]); }, "Width: " + x[1] + " of the text column"); });
      b("Alt text…", altText);
      b("Download", function () { var blk = ed.atomBlock(); if (blk) DB.getBlob(blk.blob).then(function (rec) { if (rec) S3.download(rec.data, (blk.alt || "image").replace(/[^\w\- ]+/g, "").slice(0, 60) + "." + (rec.data.type.split("/")[1] || "png").replace("jpeg", "jpg")); }); });
      b("Delete", function () { ed.deleteAtom(); });
    }
  }
  function syncTableCtx(st) {
    var x = ctxBar.querySelector('[data-k="hdr"]');
    if (x && st.inTable) { var t = doc.blocks[st.inTable.b]; x.setAttribute("aria-pressed", String(!!(t && t.hdr))); }
  }
  function altText() {
    var blk = ed.atomBlock(); if (!blk) return;
    var inp = h("textarea", { "class": "textarea prose", rows: "3", dir: "auto", maxlength: "500", "aria-label": "Alt text" }); inp.value = blk.alt || "";
    var go = h("button", { type: "button", "class": "btn primary", text: "Save" });
    var dlg = UI.modal("Alt text", h("div", { "class": "stack" }, [h("p", { "class": "faint", text: "Describe the image for people who can't see it. Saved into the .docx too." }), inp, h("div", { "class": "row" }, [go])]));
    go.addEventListener("click", function () { dlg.close(); ed.imageProp("alt", inp.value.slice(0, 500)); });
  }

  // ======================= menus =======================
  function menuItems(which) {
    var st = lastState || {};
    if (which === "file") return [
      { label: "New document", icon: "＋", onClick: function () { showHome(); } },
      { label: "Open file…", icon: "📂", kbd: "Ctrl+O", onClick: function () { $("file-in").click(); } },
      { label: "Make a copy", icon: "⧉", onClick: duplicateCurrent },
      { sep: true },
      { section: "Download" },
      { label: "Word document (.docx)", icon: "W", onClick: exportDocx },
      { label: "PDF (print → Save as PDF)", icon: "⎙", kbd: "Ctrl+P", onClick: printDoc },
      { label: "Markdown (.md)", icon: "M↓", onClick: exportMd },
      { label: "Web page (.html)", icon: "<>", onClick: exportHtml },
      { label: "Plain text (.txt)", icon: "T", onClick: exportTxt },
      { sep: true },
      { label: "Page setup…", icon: "▭", onClick: pageSetup },
      { label: "Header & footer…", icon: "▤", onClick: headerFooterDialog },
      { label: "Word count", icon: "#", onClick: wordCount },
      { sep: true },
      { label: "Delete this document", icon: "🗑", danger: true, onClick: function () { confirmDelete(doc.id, doc.title); } }
    ];
    if (which === "edit") return [
      { label: "Undo", icon: "↶", kbd: "Ctrl+Z", onClick: ed.undo }, { label: "Redo", icon: "↷", kbd: "Ctrl+Y", onClick: ed.redo },
      { sep: true },
      { label: "Cut, copy, paste", icon: "✂", hint: "Use Ctrl+X, Ctrl+C, Ctrl+V (Shift for plain text)", onClick: function () { S3.toast("Use Ctrl+X, Ctrl+C and Ctrl+V. Ctrl+Shift+V pastes plain text."); } },
      { label: "Select all", icon: "▣", kbd: "Ctrl+A", onClick: ed.selectAll },
      { label: "Find and replace", icon: "⌕", kbd: "Ctrl+H", onClick: function () { openFind(true); } }
    ];
    if (which === "insert") return [
      { label: "Image…", icon: "🖼", onClick: function () { $("img-in").click(); } },
      { label: "Table…", icon: "▦", onClick: function () { tablePop($("menus")); } },
      { label: "Link…", icon: "🔗", kbd: "Ctrl+K", onClick: function () { linkPop(null); } },
      { label: "Checklist", icon: "☑", onClick: function () { ed.list("todo"); } },
      { label: "Page break", icon: "⤓", kbd: "Ctrl+Enter", onClick: ed.pageBreak },
      { label: "Table of contents", icon: "☰", onClick: ed.insertToc },
      { sep: true },
      { section: "Page numbers" },
      { label: "Bottom centre: 1", onClick: function () { pageNumbers("footer", "{page}", "center"); } },
      { label: "Bottom right: Page 1 of N", onClick: function () { pageNumbers("footer", "Page {page} of {pages}", "right"); } },
      { label: "Top right: 1", onClick: function () { pageNumbers("header", "{page}", "right"); } },
      { label: "Header & footer…", onClick: headerFooterDialog },
      { sep: true },
      { label: "Today's date", icon: "📅", onClick: function () { insertPlain(new Date().toLocaleDateString(undefined, { year: "numeric", month: "long", day: "numeric" })); } }
    ];
    if (which === "format") return [
      { section: "Text" },
      { label: "Bold", kbd: "Ctrl+B", checked: !!(st.marks && st.marks.b), onClick: function () { ed.mark("b"); } },
      { label: "Italic", kbd: "Ctrl+I", checked: !!(st.marks && st.marks.i), onClick: function () { ed.mark("i"); } },
      { label: "Underline", kbd: "Ctrl+U", checked: !!(st.marks && st.marks.u), onClick: function () { ed.mark("u"); } },
      { label: "Strikethrough", kbd: "Ctrl+Shift+X", checked: !!(st.marks && st.marks.s), onClick: function () { ed.mark("s"); } },
      { label: "Superscript", kbd: "Ctrl+.", checked: !!(st.marks && st.marks.sup), onClick: function () { ed.mark("sup"); } },
      { label: "Subscript", kbd: "Ctrl+,", checked: !!(st.marks && st.marks.sub), onClick: function () { ed.mark("sub"); } },
      { sep: true },
      { section: "Paragraph" },
      { label: "Right-to-left", checked: !!st.rtl, onClick: function () { ed.dir("rtl"); } },
      { label: "Left-to-right", checked: !st.rtl, onClick: function () { ed.dir("ltr"); } },
      { label: "Direction: follow the text", onClick: function () { ed.dir(""); } },
      { label: "Line & paragraph spacing…", onClick: customSpacing },
      { label: "Clear formatting", kbd: "Ctrl+\\", onClick: ed.clearFormat },
      { sep: true },
      { label: "Page setup…", onClick: pageSetup }
    ];
    if (which === "view") {
      var items = [
        { label: "Outline", checked: !$("outline").hidden, onClick: toggleOutline },
        { label: "Find and replace", checked: !$("find").hidden, onClick: toggleFind },
        { sep: true }, { section: "Zoom" },
        { label: "Fit to width", checked: zoom.mode === "fit", onClick: function () { setZoom("fit"); } }
      ];
      [0.5, 0.75, 0.9, 1, 1.25, 1.5, 2].forEach(function (z) { items.push({ label: Math.round(z * 100) + "%", checked: zoom.mode === "fixed" && Math.abs(zoom.value - z) < 0.001, onClick: function () { setZoom(z); } }); });
      return items;
    }
    return [];
  }
  function insertPlain(text) { if (lastState) ed.select(lastState.sel.a, lastState.sel.f); ed._type(text); }
  $("menus").addEventListener("pointerdown", function (e) { if (e.target.closest(".dc-mb")) e.preventDefault(); });
  $("menus").addEventListener("click", function (e) {
    var b = e.target.closest("[data-menu]"); if (!b) return;
    UI.menu(b, menuItems(b.getAttribute("data-menu")), { cls: "dc-menu-wide" });
  });
  function pageNumbers(where, text, align) {
    ed.change("hf", function (d) { if (where === "footer") { d.hf.footer = text; d.hf.fa = align; } else { d.hf.header = text; d.hf.ha = align; } });
    S3.toast("Page numbers added to the " + where);
  }

  // ======================= dialogs =======================
  function seg(name, options, value) {
    var wrap = h("div", { "class": "seg", role: "group", "aria-label": name });
    options.forEach(function (o) {
      var b = h("button", { type: "button", text: o[1], "aria-pressed": String(o[0] === value), "data-v": o[0] });
      b.addEventListener("click", function () { Array.prototype.forEach.call(wrap.children, function (x) { x.setAttribute("aria-pressed", String(x === b)); }); wrap.value = o[0]; if (wrap.onchange2) wrap.onchange2(o[0]); });
      wrap.appendChild(b);
    });
    wrap.value = value;
    return wrap;
  }
  function pageSetup() {
    var p = doc.page, unit = S3.store("dc-unit") || "mm", k = function () { return unit === "in" ? 25.4 : 1; };
    var size = h("select", { "class": "select", "aria-label": "Paper size" });
    [["a4", "A4 (210 × 297 mm)"], ["letter", "Letter (8.5 × 11 in)"], ["legal", "Legal (8.5 × 14 in)"], ["a5", "A5 (148 × 210 mm)"], ["b5", "B5 (176 × 250 mm)"], ["custom", "Custom size"]].forEach(function (o) { size.appendChild(h("option", { value: o[0], text: o[1] })); });
    size.value = p.size;
    var orient = seg("Orientation", [["portrait", "Portrait"], ["landscape", "Landscape"]], p.orient);
    var units = seg("Units", [["mm", "mm"], ["in", "inches"]], unit);
    function fmt(v) { return String(Math.round(v / k() * 100) / 100); }
    var fw = numField("Width", fmt(p.w)), fh = numField("Height", fmt(p.h));
    var mt = numField("Top", fmt(p.mt)), mb = numField("Bottom", fmt(p.mb)), ml = numField("Left", fmt(p.ml)), mr = numField("Right", fmt(p.mr));
    var customRow = h("div", { "class": "grid-2" }, [fw.el, fh.el]);
    customRow.hidden = p.size !== "custom";
    size.addEventListener("change", function () { customRow.hidden = size.value !== "custom"; });
    units.onchange2 = function (u) {
      var vals = [fw, fh, mt, mb, ml, mr].map(function (f) { return f.val(0, 2000, 0) * k(); });
      unit = u; S3.store("dc-unit", u);
      [fw, fh, mt, mb, ml, mr].forEach(function (f, i) { f.inp.value = fmt(vals[i]); });
    };
    var presets = h("div", { "class": "row dc-presets" });
    [["Normal", 25.4, 25.4], ["Narrow", 12.7, 12.7], ["Moderate", 25.4, 19.05], ["Wide", 25.4, 50.8]].forEach(function (x) {
      var b = h("button", { type: "button", "class": "btn small", text: x[0] });
      b.addEventListener("click", function () { mt.inp.value = mb.inp.value = fmt(x[1]); ml.inp.value = mr.inp.value = fmt(x[2]); });
      presets.appendChild(b);
    });
    var fontSel2 = h("select", { "class": "select", "aria-label": "Default font" });
    deviceFonts().concat(deviceFonts().indexOf(doc.base.font) < 0 ? [doc.base.font] : []).forEach(function (f) { fontSel2.appendChild(h("option", { value: f, text: f })); });
    fontSel2.value = doc.base.font;
    var bs = numField("Size (pt)", doc.base.size), bl = numField("Line spacing", doc.base.line), ba = numField("Space after (pt)", doc.base.after);
    var go = h("button", { type: "button", "class": "btn primary", text: "Apply" });
    var body = h("div", { "class": "stack" }, [
      h("div", { "class": "grid-2" }, [h("label", { "class": "field" }, [h("span", { text: "Paper size" }), size]), h("div", { "class": "field" }, [h("span", { text: "Orientation" }), orient])]),
      customRow,
      h("div", { "class": "row between" }, [h("b", { text: "Margins" }), units]),
      h("div", { "class": "grid-2" }, [mt.el, mb.el, ml.el, mr.el]),
      presets,
      h("b", { text: "Normal text for this document" }),
      h("div", { "class": "grid-2" }, [h("label", { "class": "field" }, [h("span", { text: "Font" }), fontSel2]), bs.el, bl.el, ba.el]),
      h("div", { "class": "row" }, [go])
    ]);
    var dlg = UI.modal("Page setup", body);
    go.addEventListener("click", function () {
      var np = M.cleanPage({ size: size.value, orient: orient.value, w: fw.val(50, 1200 / k(), 210) * k(), h: fh.val(50, 1200 / k(), 297) * k(), mt: mt.val(0, 150, 25.4) * k(), mb: mb.val(0, 150, 25.4) * k(), ml: ml.val(0, 150, 25.4) * k(), mr: mr.val(0, 150, 25.4) * k(), hd: doc.page.hd, fd: doc.page.fd });
      if (size.value === "custom" && orient.value === "landscape" && np.w < np.h) { var t = np.w; np.w = np.h; np.h = t; np.orient = "landscape"; }
      var nb = M.cleanBase({ font: fontSel2.value, size: bs.val(4, 72, 11), line: bl.val(0.8, 4, 1.15), after: ba.val(0, 72, 8) });
      dlg.close();
      ed.change("page", function (d) { d.page = np; d.base = nb; });
      ed.relayout(); applyZoom();
    });
  }
  function headerFooterDialog() {
    var hf = doc.hf;
    var hi = h("input", { "class": "input", dir: "auto", value: hf.header, maxlength: "500", "aria-label": "Header text", placeholder: "e.g. {title}" });
    var fi = h("input", { "class": "input", dir: "auto", value: hf.footer, maxlength: "500", "aria-label": "Footer text", placeholder: "e.g. Page {page} of {pages}" });
    var ha = seg("Header alignment", [["left", "Left"], ["center", "Centre"], ["right", "Right"]], hf.ha), fa = seg("Footer alignment", [["left", "Left"], ["center", "Centre"], ["right", "Right"]], hf.fa);
    var first = h("input", { type: "checkbox", id: "hf-first" }); first.checked = hf.first;
    var last = hi;
    [hi, fi].forEach(function (x) { x.addEventListener("focus", function () { last = x; }); });
    var tokens = h("div", { "class": "row dc-tokens" });
    [["{page}", "Page number"], ["{pages}", "Page count"], ["{title}", "Title"], ["{date}", "Date"]].forEach(function (t) {
      var b = h("button", { type: "button", "class": "chip", text: "+ " + t[1] });
      b.addEventListener("pointerdown", function (e) { e.preventDefault(); });
      b.addEventListener("click", function () { var s = last.selectionStart != null ? last.selectionStart : last.value.length, e2 = last.selectionEnd != null ? last.selectionEnd : s; last.value = last.value.slice(0, s) + t[0] + last.value.slice(e2); last.focus(); last.selectionStart = last.selectionEnd = s + t[0].length; });
      tokens.appendChild(b);
    });
    var go = h("button", { type: "button", "class": "btn primary", text: "Apply" });
    var body = h("div", { "class": "stack" }, [
      h("label", { "class": "field" }, [h("span", { text: "Header" }), hi]), ha,
      h("label", { "class": "field" }, [h("span", { text: "Footer" }), fi]), fa,
      tokens,
      h("label", { "class": "check" }, [first, h("span", { text: " Different first page (no header or footer on page 1, e.g. a title page)" })]),
      h("div", { "class": "row" }, [go])
    ]);
    var dlg = UI.modal("Header & footer", body);
    go.addEventListener("click", function () {
      var nh = M.cleanHf({ header: hi.value, footer: fi.value, ha: ha.value, fa: fa.value, first: first.checked });
      dlg.close(); ed.change("hf", function (d) { d.hf = nh; });
    });
  }
  var wordsTotal = 0;
  function wordCount() {
    var c = M.count(doc), sel = ed.selectedText(), cs = sel ? M.countText(sel) : null, plan = ed.plan();
    var rows = [["Pages", plan ? plan.pages.length : 1, ""], ["Words", c.words, cs && cs.words], ["Characters", c.chars, cs && cs.chars], ["Characters without spaces", c.charsNoSpaces, cs && cs.charsNoSpaces], ["Paragraphs", c.paragraphs, ""]];
    var tbl = h("table", { "class": "tbl" }, [h("thead", {}, [h("tr", {}, [h("th", { text: "" }), h("th", { text: "Document" }), h("th", { text: cs ? "Selection" : "" })])]),
      h("tbody", {}, rows.map(function (r) { return h("tr", {}, [h("td", { text: r[0] }), h("td", { "class": "mono", text: Number(r[1]).toLocaleString() }), h("td", { "class": "mono", text: r[2] === "" || r[2] == null ? "" : Number(r[2]).toLocaleString() })]); }))]);
    UI.modal("Word count", h("div", { "class": "tbl-wrap" }, [tbl]));
  }

  // ======================= outline, find, zoom, status =======================
  function toggleOutline() { var o = $("outline"); o.hidden = !o.hidden; S3.store("dc-outline", !o.hidden); renderOutline(); applyZoom(); syncTools(lastState); }
  $("outline-x").addEventListener("click", toggleOutline);
  function renderOutline() {
    var ol = $("ol"), plan = ed.plan();
    if ($("outline").hidden || !doc) return;
    ol.textContent = "";
    var hs = M.headings(doc);
    if (!hs.length) { ol.appendChild(h("li", { "class": "dc-ol-empty faint", text: "Headings you add (Heading 1–3) appear here." })); return; }
    hs.forEach(function (x) {
      var b = h("button", { type: "button", "class": "dc-ol-it dc-ol-l" + x.level, "data-id": x.id, dir: "auto" }, [h("span", { "class": "dc-ol-t", text: x.text }), h("span", { "class": "dc-ol-p faint", text: plan && plan.pageOf[x.id] ? String(plan.pageOf[x.id]) : "" })]);
      b.addEventListener("click", function () { ed.goTo(x.id); if (innerWidth < 900) toggleOutline(); });
      ol.appendChild(h("li", {}, [b]));
    });
  }
  function markOutline(st) {
    if ($("outline").hidden || !st || !st.sel) return;
    var bi = st.sel.start.b, cur = null;
    M.headings(doc).forEach(function (x) { if (x.b <= bi) cur = x.id; });
    Array.prototype.forEach.call($("ol").querySelectorAll(".dc-ol-it"), function (b) { b.classList.toggle("on", b.getAttribute("data-id") === cur); });
  }
  var find = { ms: [], i: -1 };
  function toggleFind() { if ($("find").hidden) openFind(false); else closeFind(); }
  function openFind(replace) {
    $("find").hidden = false;
    var t = ed.selectedText(); if (t && t.length < 100 && t.indexOf("\n") < 0) $("f-q").value = t;
    (replace && $("f-q").value ? $("f-r") : $("f-q")).focus();
    $("f-q").select && !replace && $("f-q").select();
    runFind(); applyZoom(); syncTools(lastState);
  }
  function closeFind() { $("find").hidden = true; if (window.CSS && CSS.highlights) { CSS.highlights.delete("dc-find"); CSS.highlights.delete("dc-find-cur"); } applyZoom(); syncTools(lastState); }
  $("find-x").addEventListener("click", closeFind);
  function fopts() { return { matchCase: $("f-case").checked, whole: $("f-word").checked }; }
  function runFind() {
    var q = $("f-q").value;
    find.ms = q ? M.findAll(doc, q, fopts()) : [];
    if (find.i >= find.ms.length) find.i = find.ms.length - 1;
    paintFind();
  }
  function paintFind() {
    var n = find.ms.length;
    $("f-n").textContent = !$("f-q").value ? "Type to search" : n ? (find.i >= 0 ? (find.i + 1) + " of " + n : n + " match" + (n === 1 ? "" : "es")) : "No matches";
    if (!window.CSS || !CSS.highlights || typeof Highlight === "undefined") return;
    var all = new Highlight();
    find.ms.slice(0, 2000).forEach(function (m, k) { if (k === find.i) return; var r = ed.rangeOf(m.a, m.f); if (r) all.add(r); });
    CSS.highlights.set("dc-find", all);
    var cur = find.ms[find.i] && ed.rangeOf(find.ms[find.i].a, find.ms[find.i].f);
    if (cur) CSS.highlights.set("dc-find-cur", new Highlight(cur)); else CSS.highlights.delete("dc-find-cur");
  }
  function step(d) {
    if (!find.ms.length) runFind();
    if (!find.ms.length) return;
    find.i = (find.i + d + find.ms.length) % find.ms.length;
    var m = find.ms[find.i];
    ed.select(m.a, m.f);
    paintFind();
    $(d > 0 ? "f-next" : "f-prev").focus();
  }
  $("f-q").addEventListener("input", function () { find.i = -1; runFind(); });
  $("f-q").addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); step(e.shiftKey ? -1 : 1); } if (e.key === "Escape") closeFind(); });
  $("f-r").addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); $("f-one").click(); } if (e.key === "Escape") closeFind(); });
  [$("f-case"), $("f-word")].forEach(function (x) { x.addEventListener("change", function () { find.i = -1; runFind(); }); });
  $("f-next").addEventListener("click", function () { step(1); });
  $("f-prev").addEventListener("click", function () { step(-1); });
  $("f-one").addEventListener("click", function () {
    if (find.i < 0 || !find.ms[find.i]) return step(1);
    var m = find.ms[find.i], rep = $("f-r").value;
    ed.change("replace", function (d) { M.replaceAt(d, m, rep); });
    var keep = find.i; runFind();
    if (find.ms.length) { find.i = keep - 1; step(1); } else paintFind();
  });
  $("f-all").addEventListener("click", function () {
    var q = $("f-q").value; if (!q) return;
    var n = 0, rep = $("f-r").value;
    ed.change("replace", function (d) { n = M.replaceAll(d, q, rep, fopts()); });
    find.i = -1; runFind();
    S3.toast(n ? "Replaced " + n + " match" + (n === 1 ? "" : "es") : "No matches");
  });

  function pageWidth() { var g = M.geometry(doc); return g.w; }
  function setZoom(z) {
    if (z === "fit") zoom.mode = "fit"; else { zoom.mode = "fixed"; zoom.value = Math.max(0.25, Math.min(3, z)); }
    S3.store("dc-zoom-mode", zoom.mode); S3.store("dc-zoom", zoom.value);
    applyZoom();
  }
  function applyZoom() {
    if (!doc) return;
    var avail = canvas.clientWidth - (innerWidth < 600 ? 16 : 48), pw = pageWidth(), z;
    if (zoom.mode === "fit") z = Math.min(2, avail / pw);
    else if (zoom.mode === "auto") z = Math.min(1, avail / pw);
    else z = zoom.value;
    z = Math.max(0.2, Math.round(z * 100) / 100);
    stack.style.zoom = String(z);
    $("z-val").textContent = Math.round(z * 100) + "%";
  }
  $("z-in").addEventListener("click", function () { var z = parseFloat(stack.style.zoom) || 1; setZoom([0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3].filter(function (v) { return v > z + 0.01; })[0] || 3); });
  $("z-out").addEventListener("click", function () { var z = parseFloat(stack.style.zoom) || 1; setZoom([0.25, 0.33, 0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5].filter(function (v) { return v < z - 0.01; }).pop() || 0.25); });
  $("z-val").addEventListener("click", function () { setZoom(zoom.mode === "fit" ? 1 : "fit"); });
  window.addEventListener("resize", S3.debounce(applyZoom, 120));
  $("st-words").addEventListener("click", wordCount);
  // status-bar buttons must not take focus from the page (keyboard shortcuts keep working)
  document.querySelector(".dc-status").addEventListener("pointerdown", function (e) { if (e.target.closest("button")) e.preventDefault(); });

  var countTimer = 0;
  function onRendered(plan) {
    $("st-page").textContent = "Page " + ((lastState && lastState.page) || 1) + " of " + plan.pages.length;
    var p = doc.page;
    $("st-size").textContent = (M.SIZES[p.size] ? p.size.toUpperCase().replace("LETTER", "Letter").replace("LEGAL", "Legal") : Math.round(p.w) + "×" + Math.round(p.h) + " mm") + (p.orient === "landscape" ? " landscape" : "");
    clearTimeout(countTimer);
    countTimer = setTimeout(function () { wordsTotal = M.count(doc).words; renderOutline(); if (!$("find").hidden) runFind(); if (lastState) syncTools(lastState); }, 250);
  }

  // ======================= print / export =======================
  function baseName() { return (doc.title || "Document").replace(/[\\/:*?"<>|\u0000-\u001F]+/g, " ").trim().slice(0, 80) || "Document"; }
  function printDoc() {
    var p = doc.page;
    $("dc-print").textContent = "@page{size:" + p.w + "mm " + p.h + "mm;margin:0}";
    window.print();
    S3.done();
  }
  var zoomBeforePrint = "";
  window.addEventListener("beforeprint", function () { if (doc) { var p = doc.page; $("dc-print").textContent = "@page{size:" + p.w + "mm " + p.h + "mm;margin:0}"; } zoomBeforePrint = stack.style.zoom; stack.style.zoom = "1"; document.body.classList.add("dc-printing"); });
  window.addEventListener("afterprint", function () { stack.style.zoom = zoomBeforePrint; document.body.classList.remove("dc-printing"); });
  function exportDocx() {
    saveNow();
    setSaved("Making .docx…");
    return window.DX.exportDocx(doc, bytesOf, ed.plan() ? ed.plan().pageOf : null).then(function (blob) {
      S3.download(blob, baseName() + ".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
      setSaved("Saved on this device"); S3.done();
    }).catch(function (e) { setSaved(""); S3.toast("Could not make the .docx: " + (e && e.message ? e.message : e), "error"); });
  }
  function imagesMap() {
    var ids = doc.blocks.filter(function (b) { return b.type === "img" && b.blob; }).map(function (b) { return b.blob; }), map = {};
    return Promise.all(ids.map(function (id) { return dataUrlOf(id).then(function (u) { if (u) map[id] = u; }); })).then(function () { return map; });
  }
  function exportHtml() { imagesMap().then(function (map) { S3.download(window.DIO.docToHtml(doc, map, ed.plan() && ed.plan().pageOf), baseName() + ".html", "text/html"); S3.done(); }); }
  function exportTxt() { S3.download(window.DIO.docToText(doc), baseName() + ".txt", "text/plain"); S3.done(); }
  function exportMd() {
    var imgs = doc.blocks.filter(function (b) { return b.type === "img" && b.blob; });
    if (!imgs.length) { S3.download(window.DIO.docToMarkdown(doc), baseName() + ".md", "text/markdown"); S3.done(); return; }
    var names = {}, files = [];
    Promise.all(imgs.map(function (b, k) {
      return DB.getBlob(b.blob).then(function (rec) {
        if (!rec) return;
        var ext = (rec.data.type.split("/")[1] || "png").replace("jpeg", "jpg"), name = "images/image-" + (k + 1) + "." + ext;
        names[b.id] = name;
        return rec.data.arrayBuffer().then(function (ab) { files.push({ name: name, data: new Uint8Array(ab) }); });
      });
    })).then(function () {
      files.unshift({ name: baseName() + ".md", data: window.DIO.docToMarkdown(doc, function (b) { return names[b.id] || "image"; }) });
      return S3.zip(files);
    }).then(function (z) { S3.download(z, baseName() + " (markdown).zip", "application/zip"); S3.done(); });
  }

  // ======================= open / import =======================
  function openFile(file) {
    if (!file) return;
    var name = file.name || "Document", ext = (name.split(".").pop() || "").toLowerCase();
    if (file.size > 60e6) { S3.toast("That file is over 60 MB.", "error"); return; }
    if (["docx", "md", "markdown", "txt", "html", "htm"].indexOf(ext) < 0) {
      // no useful extension: a zip is treated as .docx, anything else as text
      return file.slice(0, 4).arrayBuffer().then(function (ab) { var u = new Uint8Array(ab); openFileAs(file, u[0] === 0x50 && u[1] === 0x4b ? "docx" : /wordprocessingml/.test(file.type) ? "docx" : "txt"); });
    }
    return openFileAs(file, ext);
  }
  function openFileAs(file, ext) {
    var name = file.name || "Document", title = name.replace(/\.(docx|md|markdown|txt|html?)$/i, "").slice(0, 200) || "Document";
    if (ext === "docx") {
      setSaved("Opening…");
      return S3.readFile(file).then(function (buf) { return window.DX.importDocx(buf); }).then(function (res) {
        return res.images.reduce(function (pr, im) {
          return pr.then(function () { return storeImage(new Blob([im.data], { type: im.type === "jpg" ? "image/jpeg" : "image/" + im.type }), im.id).catch(function () { res.doc.blocks = res.doc.blocks.filter(function (b) { return b.blob !== im.id; }); }); });
        }, Promise.resolve()).then(function () {
          res.doc.title = title;
          M.normalize(res.doc);
          return adopt(res.doc).then(function () {
            S3.toast(res.warnings.length ? "Opened. Note: " + res.warnings.join("; ") + "." : "Opened " + name);
            S3.done();
          });
        });
      }).catch(function (e) { setSaved(""); S3.toast(e && e.message ? e.message : "Could not open that file", "error"); });
    }
    if (ext === "md" || ext === "markdown" || ext === "txt" || ext === "html" || ext === "htm") {
      return S3.readFile(file, "text").then(function (text) {
        var d = M.newDoc({ title: title }), blocks;
        if (ext === "txt") blocks = window.DIO.textToBlocks(text);
        else if (ext === "html" || ext === "htm") { var res = window.DIO.htmlToBlocks(text); blocks = res.blocks; d.blocks = blocks.length ? blocks : d.blocks; M.normalize(d); return adopt(d).then(function () { if (res.images.length) adoptImages(res.images); S3.done(); }); }
        else blocks = window.DIO.mdToBlocks(text);
        d.blocks = blocks.length ? blocks : d.blocks;
        M.normalize(d);
        return adopt(d).then(function () { S3.done(); });
      });
    }
    S3.toast("Open a .docx, .md, .txt or .html file", "error");
  }
  function adopt(d) { d.created = d.updated = Date.now(); return DB.putDocs([d]).then(function () { openDoc(d); }); }
  $("file-in").addEventListener("change", function () { var f = this.files[0]; this.value = ""; openFile(f); });
  $("img-in").addEventListener("change", function () { var fs = Array.prototype.slice.call(this.files); this.value = ""; if (fs.length) ed.insertImages(fs); });
  $("open-btn").addEventListener("click", function () { $("file-in").click(); });
  S3.drop($("drop"), function (files) { openFile(files[0]); }, { accept: ".docx,.md,.markdown,.txt,.html,.htm" });

  // ======================= home =======================
  function showHome() {
    saveNow();
    $("ed").hidden = true; $("home").hidden = false;
    document.body.classList.add("dc-at-home");
    renderHome();
  }
  function rel(t) {
    var s = (Date.now() - t) / 1000;
    if (s < 60) return "just now"; if (s < 3600) return Math.round(s / 60) + " min ago"; if (s < 86400) return Math.round(s / 3600) + " h ago";
    return new Date(t).toLocaleDateString();
  }
  function renderHome() {
    var tp = $("tpls");
    if (!tp.children.length) window.DT.list.forEach(function (t) {
      var lines = h("span", { "class": "dc-mini-lines" + (t.id === "arletter" ? " rtl" : "") });
      for (var k = 0; k < (t.id === "blank" ? 0 : 7); k++) lines.appendChild(h("i", {}));
      var card = h("button", { type: "button", "class": "dc-tpl dc-tpl-" + t.id }, [h("span", { "class": "dc-mini" }, [h("b", { text: t.id === "blank" ? "+" : t.name, dir: "auto" }), lines]), h("span", { "class": "dc-tpl-n", text: t.name, dir: "auto" }), h("span", { "class": "faint dc-small", text: t.desc })]);
      card.addEventListener("click", function () { var d = window.DT.make(t.id); if (t.id === "blank") d.title = ""; adopt(d); });
      tp.appendChild(card);
    });
    DB.allDocs().then(function (list) {
      var ul = $("recent"); ul.textContent = "";
      list.sort(function (a, b) { return b.updated - a.updated; });
      if (!list.length) { ul.appendChild(h("li", { "class": "empty", text: "No documents yet. Start one above or open a .docx." })); return; }
      list.forEach(function (d) {
        var words = 0; try { words = M.count(M.cleanDoc(d)).words; } catch (e) { words = 0; }
        var open = h("button", { type: "button", "class": "dc-rec-open" }, [h("span", { "class": "dc-rec-ic", text: "▤", "aria-hidden": "true" }), h("span", { "class": "dc-rec-t", text: d.title || "Untitled document", dir: "auto" }), h("span", { "class": "faint dc-small", text: rel(d.updated) + " · " + words.toLocaleString() + " words" })]);
        open.addEventListener("click", function () { DB.getDoc(d.id).then(function (x) { if (x) openDoc(M.cleanDoc(x)); }); });
        var dup = h("button", { type: "button", "class": "btn small ghost", text: "Copy", title: "Make a copy" });
        dup.addEventListener("click", function () { var c = M.cleanDoc(M.clone(d)); c.id = M.uid(); c.title = (d.title || "Untitled") + " (copy)"; c.created = c.updated = Date.now(); DB.putDocs([c]).then(renderHome); });
        var del = h("button", { type: "button", "class": "btn small ghost", text: "Delete", title: "Delete this document" });
        del.addEventListener("click", function () { confirmDelete(d.id, d.title); });
        ul.appendChild(h("li", { "class": "dc-rec" }, [open, h("span", { "class": "row" }, [dup, del])]));
      });
    });
  }
  function confirmDelete(id, title) {
    var yes = h("button", { type: "button", "class": "btn primary", text: "Delete" }), no = h("button", { type: "button", "class": "btn", text: "Keep it" });
    var dlg = UI.modal("Delete document?", h("div", { "class": "stack" }, [h("p", { text: "“" + (title || "Untitled document") + "” will be removed from this browser. This can't be undone; download a .docx first if you may need it." }), h("div", { "class": "row" }, [yes, no])]));
    no.addEventListener("click", dlg.close);
    yes.addEventListener("click", function () {
      dlg.close();
      DB.deleteDocs([id]).then(gcBlobs).then(function () {
        if (doc && doc.id === id) { doc = null; S3.store("dc-last", undefined); showHome(); } else renderHome();
        S3.toast("Deleted");
      });
    });
  }
  function gcBlobs() {
    return Promise.all([DB.allDocs(), DB.blobKeys()]).then(function (r) {
      var used = {};
      r[0].forEach(function (d) { (d.blocks || []).forEach(function (b) { if (b && b.type === "img") used[b.blob] = 1; }); });
      var gone = r[1].filter(function (k) { return !used[k]; });
      gone.forEach(function (k) { if (urls[k]) { URL.revokeObjectURL(urls[k]); delete urls[k]; } });
      return gone.length ? DB.deleteBlobs(gone) : null;
    });
  }
  function duplicateCurrent() {
    saveNow();
    var c = M.cleanDoc(M.clone(doc)); c.id = M.uid(); c.title = (doc.title || "Untitled") + " (copy)";
    adopt(c).then(function () { S3.toast("Copy made: you are now editing the copy"); });
  }
  $("new-blank").addEventListener("click", function () { var d = window.DT.make("blank"); d.title = ""; adopt(d); });
  $("home-btn").addEventListener("click", showHome);

  function openDoc(d) {
    if (doc && doc.id !== d.id) saveNow();
    doc = d;
    $("home").hidden = true; $("ed").hidden = false;
    document.body.classList.remove("dc-at-home");
    $("doc-title").value = doc.title || "";
    S3.store("dc-last", doc.id);
    find.ms = []; find.i = -1;
    ed.load(doc);
    applyZoom();
    canvas.scrollTop = 0;
    setSaved(persistent ? "Saved on this device" : "Storage blocked: download a .docx to keep your work");
    renderOutline();
  }
  $("doc-title").addEventListener("input", function () { doc.title = this.value.slice(0, 200); scheduleSave(); if (/\{title\}/.test(doc.hf.header + doc.hf.footer)) ed.render(); });
  $("doc-title").addEventListener("keydown", function (e) { if (e.key === "Enter") { e.preventDefault(); ed.focusStart(); } });
  $("print-btn").addEventListener("click", printDoc);
  $("docx-btn").addEventListener("click", exportDocx);
  canvas.addEventListener("pointerdown", function (e) {
    // a click on the grey desk below the last page puts the caret at the end
    if (e.target === canvas || (e.target === stack && stack.lastElementChild && e.clientY > stack.lastElementChild.getBoundingClientRect().bottom)) { e.preventDefault(); ed.focusEnd(); }
  });

  // app-level shortcuts
  document.addEventListener("keydown", function (e) {
    if ($("ed").hidden || !doc) return;
    var mod = e.ctrlKey || e.metaKey, k = e.key.toLowerCase();
    if (!mod || e.altKey) return;
    if (k === "p") { e.preventDefault(); printDoc(); }
    else if (k === "f" && !e.shiftKey) { e.preventDefault(); openFind(false); }
    else if (k === "h") { e.preventDefault(); openFind(true); }
    else if (k === "k") { e.preventDefault(); linkPop(null); }
    else if (k === "s") { e.preventDefault(); saveNow().then(function () { S3.toast("Saved on this device. Use .docx to download a copy."); }); }
    else if (k === "o") { e.preventDefault(); $("file-in").click(); }
  });


  // ======================= start =======================
  function topOffset() { var t = document.querySelector(".s3-top"); document.documentElement.style.setProperty("--dc-top", (t ? t.offsetHeight : 0) + "px"); }
  topOffset(); setTimeout(topOffset, 0); window.addEventListener("resize", S3.debounce(topOffset, 100));
  buildTools();
  var ready = DB.open().then(function (ok) {
    persistent = ok;
    $("store-note").textContent = ok ? "Stored in this browser only" : "Browser storage is blocked here: documents last until you close the tab";
    var last = S3.store("dc-last");
    return (last ? DB.getDoc(last) : Promise.resolve(null)).then(function (d) {
      if (d) { openDoc(M.cleanDoc(d)); return; }
      return DB.allDocs().then(function (list) { if (list.length) showHome(); else adopt(welcomeDoc()); });
    });
  }).catch(function (e) { S3.toast("Could not start: " + (e && e.message), "error"); });

  function welcomeDoc() {
    var d = M.newDoc({ title: "Welcome" });
    d.blocks = [
      M.para("title", "Welcome to Docs"),
      M.para("subtitle", "A private word processor that works on real pages"),
      M.para("normal", "Start typing anywhere. Pick a style on the left of the toolbar, format with Ctrl+B / I / U, and add lists, tables, images and links from the toolbar or the Insert menu."),
      M.para("normal", [M.withText({ b: true }, "Open a .docx: "), M.withText({}, "File → Open file, or drop it on the page. "), M.withText({ b: true }, "Save one: "), M.withText({}, "the .docx button at the top. "), M.withText({ b: true }, "PDF: "), M.withText({}, "print and choose Save as PDF; the pages come out exactly as you see them.")]),
      M.para("normal", "Your documents autosave in this browser and never leave this device.", { list: { k: "todo", l: 0, c: true } }),
      M.para("normal", "Try a template: ☰ (top left) shows all documents and templates.", { list: { k: "todo", l: 0, c: false } }),
      M.para("normal")
    ];
    d.hf = M.cleanHf({ footer: "Page {page} of {pages}", fa: "center" });
    return d;
  }

  // ======================= demo (README screenshot) =======================
  function chartBlob() {
    var c = document.createElement("canvas"), W = 1200, H = 520; c.width = W; c.height = H;
    var x = c.getContext("2d");
    x.fillStyle = "#ffffff"; x.fillRect(0, 0, W, H);
    x.fillStyle = "#17233a"; x.font = "600 30px Arial, sans-serif"; x.fillText("Visits per quarter (thousands)", 40, 56);
    var data = [["North", 42, 51], ["South", 35, 47], ["East", 28, 39], ["West", 50, 58]], max = 60, bw = 70, gx = 110, top = 100, bottom = H - 70;
    x.strokeStyle = "#e3e7ee"; x.lineWidth = 2; x.font = "22px Arial, sans-serif"; x.fillStyle = "#6b7280";
    for (var g = 0; g <= 60; g += 20) { var yy = bottom - (bottom - top) * g / max; x.beginPath(); x.moveTo(80, yy); x.lineTo(W - 40, yy); x.stroke(); x.fillText(String(g), 36, yy + 8); }
    data.forEach(function (d, k) {
      var x0 = gx + k * 270;
      [[d[1], "#9fb8d9"], [d[2], "#1d3557"]].forEach(function (v, j) {
        var hgt = (bottom - top) * v[0] / max; x.fillStyle = v[1]; x.fillRect(x0 + j * (bw + 12), bottom - hgt, bw, hgt);
        x.fillStyle = "#17233a"; x.font = "600 22px Arial, sans-serif"; x.fillText(String(v[0]), x0 + j * (bw + 12) + 18, bottom - hgt - 10);
      });
      x.fillStyle = "#374151"; x.font = "24px Arial, sans-serif"; x.fillText(d[0], x0 + 40, bottom + 38);
    });
    x.fillStyle = "#9fb8d9"; x.fillRect(W - 300, 34, 22, 22); x.fillStyle = "#1d3557"; x.fillRect(W - 170, 34, 22, 22);
    x.fillStyle = "#374151"; x.font = "22px Arial, sans-serif"; x.fillText("Q2", W - 268, 53); x.fillText("Q3", W - 138, 53);
    return new Promise(function (res) { c.toBlob(res, "image/png"); });
  }
  function demoDoc(img) {
    var r = function (t, m) { return M.withText(m || {}, t); }, P = M.para, L = function (t, k, l, c) { return P("normal", t, { list: { k: k, l: l || 0, c: c } }); };
    var t = M.table(5, 4, [["Region", "Q2 visits", "Q3 visits", "Change"], ["North", "42,100", "51,300", "+21.9%"], ["South", "35,400", "47,200", "+33.3%"], ["East", "28,900", "39,100", "+35.3%"], ["West", "50,200", "58,000", "+15.5%"]]);
    t.cw = [0.34, 0.22, 0.22, 0.22];
    t.rows.slice(1).forEach(function (row) { row[3].text = M.setMark(row[3].text, 0, M.len(row[3].text), "c", "#1e7a46"); row[3].al = "right"; row[1].al = row[2].al = "right"; });
    t.rows[0].forEach(function (cl, k) { if (k) cl.al = "right"; });
    var d = M.newDoc({ title: "Quarterly Field Report", hf: { header: "{title}", footer: "Page {page} of {pages}", ha: "right", fa: "center" } });
    d.blocks = [
      P("title", "Quarterly Field Report"),
      P("subtitle", "Clinic outreach programme · Q3 2026"),
      P("normal", [r("Prepared for ", { c: "#5a6372" }), r("the regional team", { b: true, c: "#5a6372" }), r("  ·  " + new Date().toLocaleDateString(undefined, { year: "numeric", month: "long", day: "numeric" }), { c: "#5a6372" })], { sa: 14 }),
      P("h1", "Summary"),
      P("normal", [r("Visits grew in every region this quarter, led by the "), r("East (+35%)", { b: true }), r(" after the mobile unit started weekly stops. Waiting times fell for the second quarter in a row, and the new booking line now handles "), r("one in four", { i: true }), r(" appointments.")], { align: "justify" }),
      L([r("Total visits: ", { b: true }), r("195,600 (+25.6% on Q2)")], "ul"),
      L([r("Average wait: ", { b: true }), r("18 minutes, down from 26")], "ul"),
      L([r("Satisfaction: ", { b: true }), r("4.6 out of 5 across 3,210 surveys")], "ul"),
      P("h1", "Results by region"),
      P("normal", "The table compares visits with the previous quarter. Every region improved; the smallest gain was in the West, which was already the busiest.", { align: "justify" }),
      t,
      img,
      P("normal", [r("Figure 1. ", { b: true, z: 9.5, c: "#5a6372" }), r("Visits per region, Q2 against Q3.", { i: true, z: 9.5, c: "#5a6372" })], { align: "center", sa: 12 }),
      P("h2", "What worked"),
      L("Weekly mobile-unit stops in the East.", "ol"),
      L("Text-message reminders the day before each visit.", "ol"),
      L("Ten minutes of slack built into each morning session.", "ol"),
      L("Fewer first-visit no-shows as a result.", "ol", 1),
      P("quote", "“The reminder message meant I didn’t lose my slot. I was seen in fifteen minutes.” — patient survey, East region"),
      P("h2", "Next steps"),
      L("Extend the reminder messages to the South", "todo", 0, true),
      L("Pilot Saturday sessions in the West", "todo", 0, false),
      L("Review staffing for the winter peak", "todo", 0, false),
      P("h1", "Notes"),
      P("normal", "All figures come from the booking system and the monthly survey export. Numbers are rounded to the nearest hundred; percentage changes are calculated on the unrounded totals.", { align: "justify" }),
      P("normal", [r("ملاحظة: ", { b: true }), r("يدعم المحرر الفقرات العربية من اليمين إلى اليسار، ويمكن تبديل الاتجاه لكل فقرة.")], { dir: "rtl" }),
      P("normal")
    ].filter(Boolean);
    return d;
  }
  window.S3_DEMO = function () {
    return ready.then(function () { return chartBlob(); }).then(function (blob) { return storeImage(blob); }).then(function (img) {
      var g = M.geometry(M.newDoc({})), w = Math.round(g.cw * 0.86), ib = M.image(img.id, w, Math.round(img.h * w / img.w), { alt: "Bar chart of visits per region, Q2 against Q3" });
      var d = demoDoc(ib);
      if (S3.store("dc-outline") == null && innerWidth >= 1100) S3.store("dc-outline", true);
      return adopt(d);
    }).then(function () {
      if (S3.store("dc-outline") && $("outline").hidden) toggleOutline();
      if (document.activeElement) document.activeElement.blur();
      var s = window.getSelection(); if (s) s.removeAllRanges();
      canvas.scrollTop = 0;
    });
  };
  if (S3.store("dc-outline") && innerWidth >= 900) $("outline").hidden = false;
  window.DOCS_APP = { ed: ed, doc: function () { return doc; } };
})();
