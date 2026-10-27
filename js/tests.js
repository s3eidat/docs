/* Docs: self-test. Model operations, pagination against known page counts, .docx export -> import round trip,
 * a hand-made hostile .docx, contents, counting and XSS payloads. Returns true, or the list of what failed. */
(function () {
  "use strict";
  var M = window.DM, R = window.DR;
  var HTTP = "http" + "://", HTTPS = "https" + "://";
  var NS = { w: HTTP + "schemas.openxmlformats.org/wordprocessingml/2006/main", r: HTTP + "schemas.openxmlformats.org/officeDocument/2006/relationships",
    a: HTTP + "schemas.openxmlformats.org/drawingml/2006/main", wp: HTTP + "schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", pr: HTTP + "schemas.openxmlformats.org/package/2006/relationships" };

  function S(t, m) { return M.withText(m || {}, t); }
  function txt(b) { return M.plain(b.text); }
  function span(b, t) { return b.text.filter(function (s) { return s.t === t; })[0] || {}; }

  function modelTests(ok) {
    var d = M.newDoc({});
    d.blocks = [M.para("normal", "Hello world"), M.para("normal", "abcdef")];
    var p = M.splitAt(d, M.P(0, 5));
    ok("split", d.blocks.length === 3 && txt(d.blocks[0]) === "Hello" && txt(d.blocks[1]) === " world" && p.b === 1 && p.o === 0);
    var q = M.backAtStart(d, M.P(1, 0));
    ok("merge", txt(d.blocks[0]) === "Hello world" && q.b === 0 && q.o === 5);
    M.setMarkRange(d, M.P(0, 6), M.P(1, 3), "b", true);
    ok("marks across paragraphs", M.plain(d.blocks[0].text.filter(function (s) { return s.b; })) === "world" && M.plain(d.blocks[1].text.filter(function (s) { return s.b; })) === "abc");
    var c = M.insertTextAt(d, M.P(0, 11), "!");
    ok("typing inherits marks", span(d.blocks[0], "world!").b === true && c.o === 12);
    ok("toggle off", M.toggleMarkRange(d, M.P(0, 6), M.P(0, 12), "b") === false && !d.blocks[0].text.some(function (s) { return s.b; }));
    M.deleteRange(d, M.P(0, 3), M.P(1, 3));
    ok("delete across paragraphs", txt(d.blocks[0]) === "Heldef" && d.blocks.length === 1);
    // lists
    var L = M.newDoc({});
    L.blocks = [M.para("normal", "a"), M.para("normal", "b"), M.para("normal", "c"), M.para("normal", "d")];
    M.toggleList(L, M.P(0, 0), M.P(3, 0), "ol");
    M.indent(L, M.P(2, 0), M.P(2, 0), 1);
    var lab = M.listLabels(L);
    ok("numbering", lab.join(",") === "1.,2.,a.,3.");
    ok("list off", M.toggleList(L, M.P(0, 0), M.P(3, 0), "ol") === false && !L.blocks.some(function (b) { return b.list; }));
    L.blocks[0].list = { k: "ul", l: 0 }; L.blocks[0].text = [];
    M.splitAt(L, M.P(0, 0));
    ok("enter on empty item ends list", !L.blocks[0].list);
    // tables
    var t = M.table(3, 3);
    M.addRow(t, 1); M.addCol(t, 3); M.delCol(t, 0);
    ok("table add/remove", t.rows.length === 4 && t.rows[0].length === 3 && Math.abs(t.cw.reduce(function (a, b) { return a + b; }, 0) - 1) < 1e-9);
    t.rows[0][0].text = M.spansOf("x"); t.rows[0][1].text = M.spansOf("y");
    M.mergeCells(t, 0, 0, 1, 1);
    ok("merge cells", t.rows[0][0].cs === 2 && t.rows[0][0].rs === 2 && t.rows[1][1].m && txt(t.rows[0][0]) === "x\ny");
    M.addRow(t, 1);
    ok("insert row through a merge unmerges", t.rows[0][0].rs === 1 && !t.rows[1][1].m);
    // find / replace / count
    var F = M.newDoc({}); F.blocks = [M.para("normal", "The cat and the Cat."), M.para("normal", "category")];
    ok("find", M.findAll(F, "cat").length === 3 && M.findAll(F, "cat", { whole: true }).length === 2 && M.findAll(F, "Cat", { matchCase: true }).length === 1);
    ok("replace all", M.replaceAll(F, "cat", "dog", { whole: true }) === 2 && txt(F.blocks[0]) === "The dog and the dog.");
    ok("word count", M.countText("Hello world, this is a test.").words === 6 && M.countText("مرحبا بالعالم").words === 2 && M.countText("it's e-mail 3.5").words === 3);
    // links
    ok("safe links", M.safeUrl("javascript:alert(1)") === null && M.safeUrl(" JaVaScRiPt:alert(1)") === null && M.safeUrl("data:text/html,x") === null && M.safeUrl("vbscript:x") === null &&
      M.safeUrl(HTTPS + "ok.example/a b") === HTTPS + "ok.example/a%20b" && M.safeUrl("me@mail.example") === "mailto:me@mail.example" && M.safeUrl("example.org") === HTTPS + "example.org" && M.safeUrl("java\nscript:alert(1)") === null);
    ok("clean spans refuse bad values", M.cleanSpans([{ t: "x", c: "red;x", f: "a;}b{", z: 9999, a: "javascript:1", h: "#ffff00" }]).every(function (s) { return !s.c && !s.f && !s.z && !s.a && s.h === "#ffff00"; }));
  }

  function xssTests(ok) {
    var b = M.para("normal", '<img src=x onerror="window.__dcPwn=1"><script>window.__dcPwn=2</script>');
    var el = R.paraEl(M.newDoc({}), b, "", 0);
    ok("xss text inert", !el.querySelector("img,script") && R.textOf(el) === txt(b));
    var res = window.DIO.htmlToBlocks('<p onclick="x()">hi <a href="javascript:alert(1)">there</a><script>window.__dcPwn=3</script><img src="' + HTTPS + 'evil.example/x.png" onerror="y()"></p><iframe src="x"></iframe><svg><script>z()</script></svg>');
    ok("paste html sanitized", res.blocks.length === 1 && txt(res.blocks[0]) === "hi there" && !res.blocks[0].text.some(function (s) { return s.a; }) && !res.images.length);
    var html = window.DIO.blocksToHtml([M.para("normal", '"><script>alert(1)</script>')]);
    ok("html export escapes", html.indexOf("<script>") < 0 && html.indexOf("&lt;script&gt;") >= 0);
    ok("no payload ran", typeof window.__dcPwn === "undefined");
  }

  function measurer() {
    var m = document.createElement("div");
    m.className = "dc-measure dc-body";
    document.body.appendChild(m);
    return m;
  }
  function paginationTests(ok) {
    var m = measurer(), lay = window.DL.Layout(m);
    try {
      var d = M.newDoc({ base: { font: "Arial", size: 11, line: 1.15, after: 8 } });
      d.blocks = []; for (var i = 1; i <= 100; i++) d.blocks.push(M.para("normal", "Line " + i));
      // one line = 11pt x 1.15 = 16.87 px, + 8 pt after = 27.53 px; A4 text area 246.2 mm = 930.5 px -> 34 lines per page
      var a = lay.paginate(d);
      ok("A4 page count", a.pages.length === 3 && a.pages[0].items.length === 34);
      d.blocks.splice(10, 0, M.pagebreak());
      ok("page break", lay.paginate(d).pages.length === 4);
      d.blocks.splice(10, 1);
      d.page = M.cleanPage({ size: "letter" });
      ok("Letter page count", lay.paginate(d).pages.length === 4);
      d.page = M.cleanPage({ size: "a4", orient: "landscape" });
      ok("landscape", lay.paginate(d).pages.length === 5);
      // a long paragraph splits at line boundaries and every fragment fits its page
      var words = []; for (i = 0; i < 1800; i++) words.push(["alpha", "beta", "gamma", "delta", "epsilon"][i % 5] + (i % 7 === 0 ? "," : ""));
      var big = M.newDoc({ base: { font: "Arial", size: 11, line: 1.15, after: 8 } });
      big.blocks = [M.para("normal", "Intro"), M.para("normal", words.join(" "), { align: "justify" }), M.para("normal", "End")];
      var b = lay.paginate(big), frags = [], fits = true, g = M.geometry(big);
      b.pages.forEach(function (pg) { var sum = 0; pg.items.forEach(function (it) { sum += it.h; if (it.bi === 1) frags.push(it); }); if (sum - 11 > g.ch + 1) fits = false; });
      var L = M.len(big.blocks[1].text), chain = frags.length >= 2 && frags[0].from === 0 && frags[frags.length - 1].to === L && frags.every(function (f, k) { return !k || f.from === frags[k - 1].to; });
      ok("paragraph splits across pages", b.pages.length >= 3 && chain && fits);
      // the split never cuts inside a word
      var text = M.plain(big.blocks[1].text);
      ok("split at word boundary", frags.slice(1).every(function (f) { return /\s/.test(text[f.from - 1]) && !/\s/.test(text[f.from]); }));
      // tables split by rows and repeat the header row
      var rows = []; for (i = 0; i < 70; i++) rows.push(["Row " + i, "x", "y"]);
      var td = M.newDoc({ base: { font: "Arial", size: 11, line: 1.15, after: 8 } });
      td.blocks = [M.table(70, 3, rows), M.para("normal")];
      var tp = lay.paginate(td), tf = []; tp.pages.forEach(function (pg) { pg.items.forEach(function (it) { if (it.k === "table") tf.push(it); }); });
      ok("table splits by rows", tf.length >= 2 && tf[0].r0 === 0 && tf[tf.length - 1].r1 === 70 && tf[1].rep === true && tf.every(function (f, k) { return !k || f.r0 === tf[k - 1].r1; }));
      // contents: headings and their pages
      var cd = M.newDoc({});
      cd.blocks = [M.toc(), M.para("h1", "Alpha"), M.para("normal", "x"), M.pagebreak(), M.para("h1", "Beta"), M.para("h2", "Gamma"), M.para("normal")];
      var cp = lay.paginate(cd), en = cp.entries;
      ok("table of contents", en.length === 3 && en.map(function (e) { return e.level; }).join() === "1,1,2" && en[0].page === 1 && en[1].page === 2 && en[2].page === 2);
      var tocEl = R.tocEl(cd, cd.blocks[0], en);
      ok("contents render", tocEl.querySelectorAll(".dc-toc-e").length === 3 && tocEl.querySelector(".dc-toc-n").textContent === "1");
    } finally { m.remove(); }
  }

  function pngBytes() {
    var c = document.createElement("canvas"); c.width = 40; c.height = 30;
    var x = c.getContext("2d"); x.fillStyle = "#c0392b"; x.fillRect(0, 0, 40, 30); x.fillStyle = "#fff"; x.fillRect(5, 5, 10, 10);
    return new Promise(function (res) { c.toBlob(function (b) { b.arrayBuffer().then(function (ab) { res(new Uint8Array(ab)); }); }, "image/png"); });
  }
  function roundTrip(ok) {
    return pngBytes().then(function (png) {
      var d = M.newDoc({ title: "RT", page: { size: "letter", orient: "landscape", mt: 20, mr: 20, mb: 20, ml: 20 }, hf: { header: "Draft", ha: "right", footer: "Page {page} of {pages}", fa: "center" } });
      var t = M.table(3, 3, [["H1", "H2", "H3"], ["merged", "", "c"], ["d", "e", "shaded"]]);
      t.rows[1][0].cs = 2; M.fixMerges(t); t.rows[2][2].bg = "#fdf2d0";
      d.blocks = [
        M.para("title", "Round Trip"), M.para("h1", "Heading One"),
        M.para("normal", [S("Bold", { b: true }), S(" "), S("italic", { i: true }), S(" "), S("under", { u: true }), S(" "), S("strike", { s: true }), S(" "), S("red", { c: "#c0392b" }), S(" "), S("hi", { h: "#ffff00" }), S(" "), S("big", { z: 16 }), S(" "), S("mono", { f: "Courier New" }), S(" "), S("link", { a: HTTPS + "example.org/x" })], { align: "justify" }),
        M.para("normal", "bullet one", { list: { k: "ul", l: 0 } }), M.para("normal", "bullet two", { list: { k: "ul", l: 1 } }),
        M.para("normal", "first", { list: { k: "ol", l: 0 } }), M.para("normal", "second", { list: { k: "ol", l: 0 } }),
        M.para("normal", "done item", { list: { k: "todo", l: 0, c: true } }),
        t, M.image("img1abc", 120, 90, { alt: "red box" }), M.pagebreak(),
        M.para("normal", "مرحبا بالعالم", { dir: "rtl" }), M.para("h2", "Second"), M.para("normal")
      ];
      return window.DX.exportDocx(d, function () { return Promise.resolve({ data: png, type: "png" }); }, {}).then(function (blob) { return blob.arrayBuffer(); })
        .then(function (buf) { return window.DX.importDocx(buf); }).then(function (res) {
          var B = res.doc.blocks;
          ok("rt: headings", B[0].style === "title" && txt(B[0]) === "Round Trip" && B[1].style === "h1" && B[12] && B[12].style === "h2");
          var p = B[2];
          ok("rt: marks", span(p, "Bold").b && span(p, "italic").i && span(p, "under").u && span(p, "strike").s && span(p, "red").c === "#c0392b" && span(p, "hi").h === "#ffff00" && span(p, "big").z === 16 && span(p, "mono").f === "Courier New");
          ok("rt: link", span(p, "link").a === HTTPS + "example.org/x" && !span(p, "link").c && p.align === "justify");
          ok("rt: lists", B[3].list && B[3].list.k === "ul" && B[3].list.l === 0 && B[4].list.l === 1 && B[5].list.k === "ol" && B[6].list.k === "ol" && B[7].list && B[7].list.k === "todo" && B[7].list.c === true && txt(B[7]) === "done item");
          var tb = B[8];
          ok("rt: table", tb && tb.type === "table" && tb.rows.length === 3 && tb.rows[0].length === 3 && tb.hdr === true && tb.rows[1][0].cs === 2 && tb.rows[1][1].m === true && tb.rows[2][2].bg === "#fdf2d0" && txt(tb.rows[2][2]) === "shaded" && !tb.rows[0][0].text.some(function (s) { return s.b; }));
          ok("rt: image", B[9] && B[9].type === "img" && B[9].w === 120 && B[9].h === 90 && res.images.length === 1 && res.images[0].type === "png" && B[9].alt === "red box");
          ok("rt: page break", B[10] && B[10].type === "pb");
          ok("rt: rtl", B[11] && B[11].dir === "rtl" && txt(B[11]) === "مرحبا بالعالم");
          var pg = res.doc.page;
          ok("rt: page size", pg.size === "letter" && pg.orient === "landscape" && Math.abs(pg.w - 279.4) < 0.5 && Math.abs(pg.mt - 20) < 0.1 && Math.abs(pg.ml - 20) < 0.1);
          ok("rt: header/footer", res.doc.hf.header === "Draft" && res.doc.hf.ha === "right" && res.doc.hf.footer === "Page {page} of {pages}" && res.doc.hf.fa === "center");
          ok("rt: base style", res.doc.base.font === "Calibri" && res.doc.base.size === 11 && Math.abs(res.doc.base.line - 1.15) < 0.01 && res.doc.base.after === 8);
          ok("rt: no extra blocks", B.length === 14);
        });
    });
  }

  function zipOf(files) {
    var z = new window.JSZip();
    Object.keys(files).forEach(function (k) { z.file(k, files[k]); });
    return z.generateAsync({ type: "arraybuffer" });
  }
  function hostile(ok) {
    var doc = '<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="' + NS.w + '" xmlns:r="' + NS.r + '" xmlns:a="' + NS.a + '" xmlns:wp="' + NS.wp + '"><w:body>' +
      '<w:p><w:r><w:t>&lt;img src=x onerror=window.__dcPwn=1&gt;</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>&lt;script&gt;window.__dcPwn=2&lt;/script&gt;</w:t></w:r></w:p>' +
      '<w:p><w:hyperlink r:id="rJs"><w:r><w:t>click</w:t></w:r></w:hyperlink></w:p>' +
      '<w:p><w:r><w:rPr><w:sz w:val="99999"/><w:color w:val="red;x"/><w:rFonts w:ascii="x;}body{display:none"/></w:rPr><w:t>big</w:t></w:r></w:p>' +
      '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="999999999999" cy="5"/><a:graphic><a:graphicData><a:blip r:link="rExt"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>' +
      '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="100" cy="100"/><a:graphic><a:graphicData><a:blip r:embed="rExtEmbed"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>' +
      '<w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val="9999"/><w:vMerge/></w:tcPr><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
      '<w:sectPr><w:pgSz w:w="99999999" w:h="5"/><w:pgMar w:top="-99999" w:left="999999" w:right="5" w:bottom="5"/><w:headerReference w:type="default" r:id="rH"/></w:sectPr></w:body></w:document>';
    var rels = '<?xml version="1.0"?><Relationships xmlns="' + NS.pr + '">' +
      '<Relationship Id="rJs" Type="' + NS.r + '/hyperlink" Target="javascript:alert(1)" TargetMode="External"/>' +
      '<Relationship Id="rExt" Type="' + NS.r + '/image" Target="' + HTTPS + 'evil.example/track.png" TargetMode="External"/>' +
      '<Relationship Id="rExtEmbed" Type="' + NS.r + '/image" Target="' + HTTPS + 'evil.example/x.png" TargetMode="External"/>' +
      '<Relationship Id="rH" Type="' + NS.r + '/header" Target="header1.xml"/></Relationships>';
    var bomb = '<?xml version="1.0"?><!DOCTYPE lolz [<!ENTITY lol "lol"><!ENTITY lol2 "&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;">]><w:hdr xmlns:w="' + NS.w + '"><w:p><w:r><w:t>&lol2;</w:t></w:r></w:p></w:hdr>';
    return zipOf({ "word/document.xml": doc, "word/_rels/document.xml.rels": rels, "word/header1.xml": bomb }).then(function (buf) { return window.DX.importDocx(buf); }).then(function (res) {
      var B = res.doc.blocks, all = B.filter(function (b) { return b.type === "p"; });
      ok("hostile: text kept literally", txt(B[0]) === "<img src=x onerror=window.__dcPwn=1>" && txt(B[1]) === "<script>window.__dcPwn=2</script>");
      var host = document.createElement("div");
      all.forEach(function (b) { host.appendChild(R.paraEl(res.doc, b, "", 0)); });
      ok("hostile: renders inert", !host.querySelector("img,script,iframe") && host.textContent.indexOf("<script>") >= 0);
      ok("hostile: javascript link dropped", !all.some(function (b) { return b.text.some(function (s) { return s.a; }); }) && !host.querySelector("a"));
      var big = all.filter(function (b) { return txt(b) === "big"; })[0];
      ok("hostile: values clamped", big && big.text[0].z <= 400 && !big.text[0].c && !big.text[0].f);
      ok("hostile: external images ignored", !B.some(function (b) { return b.type === "img"; }) && !res.images.length);
      var tb = B.filter(function (b) { return b.type === "table"; })[0];
      ok("hostile: table clamped", tb && tb.rows[0].length <= 63);
      var pg = res.doc.page;
      ok("hostile: page clamped", pg.w <= 1200 && pg.h >= 50 && pg.mt >= 0 && pg.mt <= 150 && pg.w - pg.ml - pg.mr >= 29.9 && pg.h - pg.mt - pg.mb >= 29.9);
      ok("hostile: entity header refused", res.doc.hf.header === "" && res.warnings.some(function (w) { return /header/.test(w); }));
      ok("hostile: nothing ran", typeof window.__dcPwn === "undefined");
    }).then(function () {
      var long = new Array(30000).join("<w:p><w:r><w:t>x</w:t></w:r></w:p>");
      return zipOf({ "word/document.xml": '<w:document xmlns:w="' + NS.w + '"><w:body>' + long + "</w:body></w:document>" });
    }).then(function (buf) {
      return window.DX.importDocx(buf, { maxPart: 100000 }).then(function () { ok("size cap", false); }, function (e) { ok("size cap", /too large/.test(e.message)); });
    }).then(function () {
      return window.DX.importDocx(new TextEncoder().encode("not a zip at all").buffer).then(function () { ok("not a zip", false); }, function (e) { ok("not a zip", /not a valid/.test(e.message)); });
    });
  }

  function ioTests(ok) {
    var d = M.newDoc({ title: "T" });
    d.blocks = [M.para("h1", "Head"), M.para("normal", [S("bold", { b: true }), S(" text")]), M.para("normal", "item", { list: { k: "ul", l: 0 } }), M.para("normal", "task", { list: { k: "todo", l: 0, c: true } }), M.table(2, 2, [["a", "b"], ["c", "d"]]), M.para("normal")];
    var md = window.DIO.docToMarkdown(d);
    ok("markdown export", md.indexOf("# Head") === 0 && md.indexOf("**bold** text") > 0 && md.indexOf("- item") > 0 && md.indexOf("- [x] task") > 0 && md.indexOf("| a | b |") > 0);
    var back = window.DIO.mdToBlocks(md);
    ok("markdown import", back[0].style === "h1" && back[1].text[0].b && back[2].list.k === "ul" && back[3].list.k === "todo" && back[3].list.c && back[4].type === "table");
    var tx = window.DIO.docToText(d);
    ok("text export", tx.indexOf("Head\nbold text\n• item\n☑ task\na\tb") === 0);
  }

  window.S3_SELFTEST = function () {
    var fails = [];
    function ok(name, cond) { if (!cond) fails.push(name); }
    try { modelTests(ok); } catch (e) { fails.push("model threw: " + e.message); }
    try { xssTests(ok); } catch (e) { fails.push("xss threw: " + e.message); }
    try { paginationTests(ok); } catch (e) { fails.push("pagination threw: " + e.message); }
    try { ioTests(ok); } catch (e) { fails.push("io threw: " + e.message); }
    return roundTrip(ok).catch(function (e) { fails.push("round trip threw: " + (e && e.message)); })
      .then(function () { return hostile(ok); }).catch(function (e) { fails.push("hostile threw: " + (e && e.message)); })
      .then(function () { return fails.length ? fails : true; });
  };
})();
