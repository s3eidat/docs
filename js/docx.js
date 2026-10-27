/* Docs: .docx export (docx 9, MIT) and .docx import (own reader: JSZip + DOMParser).
 * The reader treats every file as hostile: size caps before and after inflating, XML with a DOCTYPE/ENTITY is
 * refused, only the parts it understands are read, external relationships (linked images, remote templates) are
 * ignored, links go through DM.safeUrl, text becomes model spans (never HTML) and every value is clamped. */
(function () {
  "use strict";
  var M = window.DM;
  var TW = 56.6929; // twips per mm
  function twip(mm) { return Math.round(mm * TW); }
  var HL_NAME = {}; Object.keys(M.HL).forEach(function (k) { HL_NAME[M.HL[k]] = k; });
  var RTL = /[֐-ࣿיִ-﷿ﹰ-﻿]/;

  // ======================= export =======================
  function exportDocx(doc, bytesOf, pageOf) {
    var D = window.docx;
    if (!D) return Promise.reject(new Error("The .docx library did not load"));
    var imgBlocks = doc.blocks.filter(function (b) { return b.type === "img" && b.blob; });
    return Promise.all(imgBlocks.map(function (b) { return bytesOf(b.blob).catch(function () { return null; }); })).then(function (list) {
      var imgs = {};
      imgBlocks.forEach(function (b, k) { if (list[k]) imgs[b.blob] = list[k]; });
      return build(D, doc, imgs, pageOf);
    });
  }
  function build(D, doc, imgs, pageOf) {
    var p = doc.page, textW = twip(p.w - p.ml - p.mr), base = doc.base;
    var AL = { left: D.AlignmentType.LEFT, center: D.AlignmentType.CENTER, right: D.AlignmentType.RIGHT, justify: D.AlignmentType.JUSTIFIED };
    function runProps(s, rtl, extra) {
      var o = {};
      if (s.b) o.bold = true;
      if (s.i) o.italics = true;
      if (s.u) o.underline = { type: D.UnderlineType.SINGLE };
      if (s.s) o.strike = true;
      if (s.sup) o.superScript = true;
      if (s.sub) o.subScript = true;
      if (s.c) o.color = s.c.slice(1).toUpperCase();
      if (s.z) o.size = Math.round(s.z * 2);
      if (s.f) o.font = { ascii: s.f, hAnsi: s.f, cs: s.f, eastAsia: s.f };
      if (s.h) { if (HL_NAME[s.h]) o.highlight = HL_NAME[s.h]; else o.shading = { type: D.ShadingType.CLEAR, color: "auto", fill: s.h.slice(1).toUpperCase() }; }
      if (rtl && RTL.test(s.t)) o.rightToLeft = true;
      if (extra) Object.keys(extra).forEach(function (k) { if (o[k] == null) o[k] = extra[k]; });
      return o;
    }
    function textRuns(s, rtl, extra) {
      var out = [];
      s.t.split("\n").forEach(function (line, li) {
        var parts = line.split("\t");
        if (!line && li) { out.push(new D.TextRun(Object.assign(runProps(s, rtl, extra), { text: "", break: 1 }))); return; }
        parts.forEach(function (part, pi) {
          var o = runProps(s, rtl, extra);
          if (li && pi === 0) o.break = 1;
          if (pi) out.push(new D.TextRun(Object.assign(runProps(s, rtl, extra), { children: [new D.Tab()] })));
          if (part || o.break) { o.text = part; out.push(new D.TextRun(o)); }
        });
      });
      return out;
    }
    function runs(spans, rtl, extra) {
      var out = [], k = 0;
      while (k < spans.length) {
        var s = spans[k];
        if (s.a) {
          var group = [];
          while (k < spans.length && spans[k].a === s.a) { group = group.concat(textRuns(spans[k], rtl, Object.assign({ style: "Hyperlink" }, extra || {}))); k++; }
          out.push(new D.ExternalHyperlink({ link: s.a, children: group }));
          continue;
        }
        out = out.concat(textRuns(s, rtl, extra));
        k++;
      }
      return out;
    }
    var HEAD = { title: D.HeadingLevel.TITLE, h1: D.HeadingLevel.HEADING_1, h2: D.HeadingLevel.HEADING_2, h3: D.HeadingLevel.HEADING_3 };
    var PSTYLE = { subtitle: "Subtitle", quote: "Quote", code: "Code" };
    // each run of list paragraphs is its own list instance, so numbering restarts like it does on screen
    var inst = [], n = 0;
    doc.blocks.forEach(function (b, i) { if (b.type === "p" && b.list) { var prev = doc.blocks[i - 1]; if (!prev || prev.type !== "p" || !prev.list) n++; inst[i] = n; } });
    function para(b, i) {
      var rtl = window.DR.dirOf(b) === "rtl", o = { children: [] };
      if (HEAD[b.style] != null) o.heading = HEAD[b.style];
      if (PSTYLE[b.style]) o.style = PSTYLE[b.style];
      if (b.align) o.alignment = AL[b.align];
      if (rtl) o.bidirectional = true;
      var sp = {};
      if (b.sb != null) sp.before = Math.round(b.sb * 20);
      if (b.sa != null) sp.after = Math.round(b.sa * 20);
      if (b.lh) { sp.line = Math.round(b.lh * 240); sp.lineRule = D.LineRuleType.AUTO; }
      if (Object.keys(sp).length) o.spacing = sp;
      var kids = [];
      if (b.list && b.list.k !== "todo") o.numbering = { reference: b.list.k === "ol" ? "dc-num" : "dc-bul", level: b.list.l, instance: inst[i] || 0 };
      else if (b.list) { o.indent = { left: 720 * (b.list.l + 1), hanging: 360 }; kids.push(new D.TextRun({ text: (b.list.c ? "☒" : "☐") + " " })); }
      else {
        var ind = {};
        if (b.ind) ind.left = 720 * b.ind;
        if (b.fi === "first") ind.firstLine = 720;
        if (b.fi === "hang") { ind.left = (ind.left || 0) + 720; ind.hanging = 720; }
        if (Object.keys(ind).length) o.indent = ind;
      }
      if (M.HEADINGS[b.style] != null) o.keepNext = true;
      o.children = kids.concat(runs(b.text, rtl));
      return new D.Paragraph(o);
    }
    function cellParas(cl, bold) {
      var rtl = window.DR.dirOf({ dir: "", text: cl.text }) === "rtl";
      var o = { children: runs(cl.text, rtl, bold ? { bold: true } : null) };
      if (cl.al) o.alignment = AL[cl.al];
      if (rtl) o.bidirectional = true;
      o.spacing = { after: 0 };
      return [new D.Paragraph(o)];
    }
    function table(t) {
      var widths = t.cw.map(function (f) { return Math.max(100, Math.round(f * textW)); });
      return new D.Table({
        width: { size: textW, type: D.WidthType.DXA }, columnWidths: widths, layout: D.TableLayoutType.FIXED,
        rows: t.rows.map(function (row, r) {
          var hdr = t.hdr && r === 0;
          return new D.TableRow({
            tableHeader: hdr || undefined,
            children: row.map(function (cl, c) {
              if (cl.m) return null;
              var w = 0; for (var k = c; k < c + cl.cs; k++) w += widths[k] || 0;
              var o = { children: cellParas(cl, hdr), width: { size: w, type: D.WidthType.DXA } };
              if (cl.cs > 1) o.columnSpan = cl.cs;
              if (cl.rs > 1) o.rowSpan = cl.rs;
              var bg = cl.bg || (hdr ? "#eef1f5" : "");
              if (bg) o.shading = { type: D.ShadingType.CLEAR, color: "auto", fill: bg.slice(1).toUpperCase() };
              return new D.TableCell(o);
            }).filter(Boolean)
          });
        })
      });
    }
    function hfPara(text, align) {
      var parts = String(text || "").split(/(\{page\}|\{pages\}|\{title\}|\{date\})/), kids = [], st = { size: 18, color: "555555" };
      parts.forEach(function (x) {
        if (!x) return;
        if (x === "{page}") kids.push(new D.TextRun(Object.assign({ children: [D.PageNumber.CURRENT] }, st)));
        else if (x === "{pages}") kids.push(new D.TextRun(Object.assign({ children: [D.PageNumber.TOTAL_PAGES] }, st)));
        else if (x === "{title}") kids.push(new D.TextRun(Object.assign({ text: doc.title || "" }, st)));
        else if (x === "{date}") kids.push(new D.TextRun(Object.assign({ text: new Date().toLocaleDateString() }, st)));
        else kids.push(new D.TextRun(Object.assign({ text: x }, st)));
      });
      return new D.Paragraph({ alignment: AL[align] || D.AlignmentType.CENTER, children: kids });
    }
    var children = [];
    doc.blocks.forEach(function (b, i) {
      if (b.type === "p") children.push(para(b, i));
      else if (b.type === "table") { children.push(table(b)); if (doc.blocks[i + 1] && doc.blocks[i + 1].type === "table") children.push(new D.Paragraph({ spacing: { after: 0 } })); }
      else if (b.type === "pb") children.push(new D.Paragraph({ children: [new D.PageBreak()] }));
      else if (b.type === "img") {
        var im = imgs[b.blob];
        if (!im) return;
        children.push(new D.Paragraph({ alignment: AL[b.align] || D.AlignmentType.CENTER, children: [new D.ImageRun({ type: im.type, data: im.data, transformation: { width: b.w, height: b.h }, altText: { name: b.alt || "Image", description: b.alt || "", title: b.alt || "" } })] }));
      } else if (b.type === "toc") {
        children.push(new D.Paragraph({ style: "TOCHeading", children: [new D.TextRun({ text: "Contents" })] }));
        M.tocEntries(doc, pageOf).forEach(function (e) {
          children.push(new D.Paragraph({ style: "TOC" + e.level, tabStops: [{ type: D.TabStopType.RIGHT, position: textW, leader: D.LeaderType.DOT }],
            children: [new D.TextRun({ text: e.text }), new D.TextRun({ children: [new D.Tab()] }), new D.TextRun({ text: e.page ? String(e.page) : "" })] }));
        });
      }
    });
    function lvls(kind) {
      var out = [];
      for (var k = 0; k <= 8; k++) {
        out.push(kind === "ul" ? { level: k, format: D.LevelFormat.BULLET, text: ["•", "◦", "▪"][k % 3], alignment: D.AlignmentType.LEFT, style: { paragraph: { indent: { left: 720 * (k + 1), hanging: 360 } } } }
          : { level: k, format: [D.LevelFormat.DECIMAL, D.LevelFormat.LOWER_LETTER, D.LevelFormat.LOWER_ROMAN][k % 3], text: "%" + (k + 1) + ".", alignment: D.AlignmentType.LEFT, style: { paragraph: { indent: { left: 720 * (k + 1), hanging: 360 } } } });
      }
      return out;
    }
    var S = M.STYLES;
    function hs(s, extra) {
      var run = { size: Math.round((s.size || base.size) * 2) }, para = { spacing: {} };
      if (s.bold) run.bold = true;
      if (s.italic) run.italics = true;
      if (s.color) run.color = s.color.slice(1).toUpperCase();
      if (s.font) run.font = s.font;
      if (s.sb != null) para.spacing.before = Math.round(s.sb * 20);
      if (s.sa != null) para.spacing.after = Math.round(s.sa * 20);
      if (s.lh) { para.spacing.line = Math.round(s.lh * 240); para.spacing.lineRule = D.LineRuleType.AUTO; }
      return Object.assign({ run: run, paragraph: para }, extra || {});
    }
    var portrait = [p.w, p.h].sort(function (a, b) { return a - b; });
    var headers = { default: new D.Header({ children: [hfPara(doc.hf.header, doc.hf.ha)] }) };
    var footers = { default: new D.Footer({ children: [hfPara(doc.hf.footer, doc.hf.fa)] }) };
    if (doc.hf.first) { headers.first = new D.Header({ children: [new D.Paragraph({})] }); footers.first = new D.Footer({ children: [new D.Paragraph({})] }); }
    var d = new D.Document({
      creator: "S3eidat Docs", title: doc.title || "Document", description: "Made with S3eidat Docs",
      styles: {
        default: {
          document: { run: { font: base.font, size: Math.round(base.size * 2) }, paragraph: { spacing: { after: Math.round(base.after * 20), line: Math.round(base.line * 240), lineRule: D.LineRuleType.AUTO } } },
          title: hs(S.title), heading1: hs(S.h1), heading2: hs(S.h2), heading3: hs(S.h3),
          hyperlink: { run: { color: "1155CC", underline: { type: D.UnderlineType.SINGLE } } }
        },
        paragraphStyles: [
          Object.assign({ id: "Subtitle", name: "Subtitle", basedOn: "Normal", next: "Normal", quickFormat: true }, hs(S.subtitle)),
          Object.assign({ id: "Quote", name: "Quote", basedOn: "Normal", next: "Normal", quickFormat: true }, hs(S.quote, {})),
          Object.assign({ id: "Code", name: "Code", basedOn: "Normal", next: "Code", quickFormat: true }, hs(S.code)),
          { id: "TOCHeading", name: "TOC Heading", basedOn: "Normal", next: "Normal", run: { size: 28, bold: true, color: "1D3557" }, paragraph: { spacing: { before: 240, after: 120 } } },
          { id: "TOC1", name: "toc 1", basedOn: "Normal", next: "Normal", paragraph: { spacing: { after: 60 } } },
          { id: "TOC2", name: "toc 2", basedOn: "Normal", next: "Normal", paragraph: { spacing: { after: 60 }, indent: { left: 240 } } },
          { id: "TOC3", name: "toc 3", basedOn: "Normal", next: "Normal", paragraph: { spacing: { after: 60 }, indent: { left: 480 } } }
        ]
      },
      numbering: { config: [{ reference: "dc-bul", levels: lvls("ul") }, { reference: "dc-num", levels: lvls("ol") }] },
      sections: [{
        properties: {
          titlePage: doc.hf.first || undefined,
          page: { size: { width: twip(portrait[0]), height: twip(portrait[1]), orientation: p.orient === "landscape" ? D.PageOrientation.LANDSCAPE : D.PageOrientation.PORTRAIT },
            margin: { top: twip(p.mt), right: twip(p.mr), bottom: twip(p.mb), left: twip(p.ml), header: twip(p.hd), footer: twip(p.fd) } }
        },
        headers: headers, footers: footers, children: children
      }]
    });
    return D.Packer.toBlob(d);
  }

  // ======================= import =======================
  var LIMITS = { maxFile: 60e6, maxPart: 40e6, maxTotal: 400e6, maxMedia: 25e6, maxImages: 400, maxBlocks: 40000, maxChars: 5e6 };
  function kids(el, name) { var out = []; if (!el) return out; for (var c = el.firstElementChild; c; c = c.nextElementSibling) if (!name || c.localName === name) out.push(c); return out; }
  function kid(el, name) { if (!el) return null; for (var c = el.firstElementChild; c; c = c.nextElementSibling) if (c.localName === name) return c; return null; }
  function attr(el, name, ns) {
    if (!el || !el.attributes) return null;
    for (var k = 0; k < el.attributes.length; k++) {
      var a = el.attributes[k];
      if (a.localName === name && (!ns || (a.namespaceURI || "").indexOf(ns) >= 0)) return a.value;
    }
    return null;
  }
  function w(el, name) { return attr(el, name, "wordprocessingml"); }
  function rel(el, name) { return attr(el, name, "/relationships"); }
  function onOff(el) { if (!el) return null; var v = w(el, "val"); return v == null || !/^(0|false|off|none)$/i.test(v); }
  function int(v, lo, hi, def) { var n = parseInt(v, 10); return isFinite(n) ? Math.max(lo, Math.min(hi, n)) : def; }
  function hexOf(v) { return v && /^[0-9a-f]{6}$/i.test(v) ? "#" + v.toLowerCase() : null; }
  function sniff(u8) {
    if (u8.length < 12) return null;
    if (u8[0] === 0x89 && u8[1] === 0x50 && u8[2] === 0x4e && u8[3] === 0x47) return "png";
    if (u8[0] === 0xff && u8[1] === 0xd8 && u8[2] === 0xff) return "jpg";
    if (u8[0] === 0x47 && u8[1] === 0x49 && u8[2] === 0x46 && u8[3] === 0x38) return "gif";
    if (u8[0] === 0x42 && u8[1] === 0x4d) return "bmp";
    if (u8[0] === 0x52 && u8[1] === 0x49 && u8[2] === 0x46 && u8[3] === 0x46 && u8[8] === 0x57 && u8[9] === 0x45 && u8[10] === 0x42 && u8[11] === 0x50) return "webp";
    return null;
  }
  function resolve(from, target) {
    if (!target) return null;
    var t = target.replace(/\\/g, "/");
    if (/^[a-z][a-z0-9+.-]*:/i.test(t)) return null;
    var parts = t.charAt(0) === "/" ? [] : from.split("/").slice(0, -1);
    t.replace(/^\//, "").split("/").forEach(function (seg) { if (seg === "..") parts.pop(); else if (seg && seg !== ".") parts.push(seg); });
    return parts.join("/");
  }

  function importDocx(buf, limits) {
    var L = Object.assign({}, LIMITS, limits || {}), warnings = [];
    if (!buf || buf.byteLength > L.maxFile) return Promise.reject(new Error("This file is larger than " + Math.round(L.maxFile / 1e6) + " MB."));
    var zip;
    return window.JSZip.loadAsync(buf).catch(function () { throw new Error("This is not a valid .docx file (it is not a zip archive)."); }).then(function (z) {
      zip = z;
      var total = 0, bad = null;
      zip.forEach(function (path, f) {
        var u = f._data && f._data.uncompressedSize;
        if (typeof u === "number") { total += u; if (/\.xml$|\.rels$/i.test(path) && u > L.maxPart) bad = path; }
      });
      if (bad) throw new Error("A part of this file is too large to open safely (" + bad + ").");
      if (total > L.maxTotal) throw new Error("This file expands to more than " + Math.round(L.maxTotal / 1e6) + " MB, which is too large to open safely.");
      return text("_rels/.rels");
    }).then(function (rootRels) {
      var main = "word/document.xml";
      if (rootRels) {
        var r = xml(rootRels, "_rels/.rels");
        kids(r.documentElement, "Relationship").forEach(function (x) { if (/officeDocument$/.test(x.getAttribute("Type") || "") && !/^external$/i.test(x.getAttribute("TargetMode") || "")) { var t = resolve("", x.getAttribute("Target")); if (t) main = t; } });
      }
      var relPath = main.replace(/[^/]+$/, "_rels/") + main.split("/").pop() + ".rels";
      return Promise.all([text(main), text(relPath), text("word/styles.xml"), text("word/numbering.xml"), text("word/theme/theme1.xml")]).then(function (parts) {
        if (!parts[0]) throw new Error("This file has no document body (word/document.xml is missing).");
        return parse(main, parts);
      });
    });

    function text(path) {
      var f = zip.file(path);
      if (!f) return Promise.resolve(null);
      return f.async("string").then(function (s) { if (s.length > L.maxPart) throw new Error("A part of this file is too large to open safely."); return s; });
    }
    function xml(s, name) {
      if (/<!DOCTYPE|<!ENTITY/i.test(s.slice(0, 4096)) || /<!ENTITY/i.test(s)) throw new Error("This file contains XML declarations that are not allowed (" + name + ").");
      var d = new DOMParser().parseFromString(s, "application/xml");
      if (d.getElementsByTagName("parsererror").length) throw new Error("This file is damaged: " + name + " is not valid XML.");
      return d;
    }
    function relsMap(s, owner) {
      var map = {};
      if (!s) return map;
      kids(xml(s, owner + ".rels").documentElement, "Relationship").forEach(function (x) {
        var ext = /^external$/i.test(x.getAttribute("TargetMode") || "");
        map[x.getAttribute("Id")] = { type: (x.getAttribute("Type") || "").split("/").pop(), target: ext ? x.getAttribute("Target") : resolve(owner, x.getAttribute("Target")), external: ext };
      });
      return map;
    }

    function parse(main, parts) {
      var dx = xml(parts[0], main), rels = relsMap(parts[1], main);
      var styles = {}, num = { nums: {}, abs: {} }, base = { font: "Calibri", size: 11, line: 1.15, after: 8 }, themeFont = "";
      if (parts[4]) {
        try { var th = xml(parts[4], "theme"), minor = th.getElementsByTagNameNS("*", "minorFont")[0], lat = minor && kid(minor, "latin"); if (lat && M.FONT_RE.test(lat.getAttribute("typeface") || "")) themeFont = lat.getAttribute("typeface"); } catch (e) { warnings.push("theme ignored"); }
      }
      // ---- styles ----
      if (parts[2]) {
        var sx = xml(parts[2], "styles.xml"), dd = sx.getElementsByTagNameNS("*", "docDefaults")[0];
        if (dd) {
          var rpr = dd.getElementsByTagNameNS("*", "rPr")[0], ppr = dd.getElementsByTagNameNS("*", "pPr")[0];
          readBase(rpr, ppr);
        }
        kids(sx.documentElement, "style").forEach(function (st) {
          var id = w(st, "styleId"), nm = kid(st, "name"), bo = kid(st, "basedOn");
          if (!id) return;
          styles[id] = { id: id, type: w(st, "type") || "paragraph", name: (nm ? w(nm, "val") || "" : "").toLowerCase(), basedOn: bo ? w(bo, "val") : null, pPr: kid(st, "pPr"), rPr: kid(st, "rPr"), def: w(st, "default") === "1" };
        });
        Object.keys(styles).forEach(function (id) { var s = styles[id]; if (s.type === "paragraph" && (s.def || s.name === "normal")) readBase(s.rPr, s.pPr); });
      }
      function readBase(rpr, ppr) {
        if (rpr) {
          var sz = kid(rpr, "sz"); if (sz) base.size = int(w(sz, "val"), 8, 144, 22) / 2;
          var rf = kid(rpr, "rFonts"); if (rf) { var f = w(rf, "ascii") || w(rf, "hAnsi"); if (f && M.FONT_RE.test(f)) base.font = f; else if (w(rf, "asciiTheme") && themeFont) base.font = themeFont; }
        }
        if (ppr) {
          var sp = kid(ppr, "spacing");
          if (sp) {
            if (w(sp, "after") != null) base.after = int(w(sp, "after"), 0, 1440, 160) / 20;
            var ln = w(sp, "line"), rule = w(sp, "lineRule");
            if (ln != null && (!rule || rule === "auto")) base.line = Math.max(0.8, Math.min(4, int(ln, 120, 960, 276) / 240));
          }
        }
      }
      function chain(id, fn) { var seen = 0; while (id && styles[id] && seen++ < 20) { var v = fn(styles[id]); if (v != null) return v; id = styles[id].basedOn; } return null; }
      function mapStyle(id) {
        var nm = chain(id, function (s) {
          var n = s.name;
          if (n === "title") return "title";
          if (n === "subtitle") return "subtitle";
          var m = /^heading (\d)$/.exec(n);
          if (m) return +m[1] === 1 ? "h1" : +m[1] === 2 ? "h2" : "h3";
          if (n === "quote" || n === "intense quote" || n === "block text") return "quote";
          if (/code|html preformatted|plain text|source/.test(n)) return "code";
          if (/^toc \d|^toc heading$|^contents heading$/.test(n)) return "toc";
          return null;
        });
        return nm || "normal";
      }
      // ---- numbering ----
      if (parts[3]) {
        var nx = xml(parts[3], "numbering.xml");
        kids(nx.documentElement, "abstractNum").forEach(function (an) {
          var lv = {};
          kids(an, "lvl").forEach(function (l) {
            var f = kid(l, "numFmt"), k = int(w(l, "ilvl"), 0, 8, 0), pp = kid(l, "pPr"), ind = pp && kid(pp, "ind");
            lv[k] = f ? w(f, "val") : "decimal";
            lv["i" + k] = ind ? int(w(ind, "left") || w(ind, "start"), 0, 31680, 0) : 0;
          });
          var mlt = kid(an, "multiLevelType");
          lv.single = (mlt && w(mlt, "val") === "singleLevel") || kids(an, "lvl").length === 1;
          num.abs[w(an, "abstractNumId")] = lv;
        });
        kids(nx.documentElement, "num").forEach(function (n) { var a = kid(n, "abstractNumId"); if (a) num.nums[w(n, "numId")] = w(a, "val"); });
      }
      function listFor(numId, ilvl) {
        if (numId == null || numId === "0") return null;
        var abs = num.abs[num.nums[numId]];
        var fmt = abs ? abs[ilvl] || abs[0] : "bullet";
        if (fmt === "none") return null;
        // single-level lists ("List Bullet 2" and friends) show depth only through their indent, in 0.25 in steps
        var left = abs ? abs["i" + ilvl] || 0 : 0, byIndent = abs && abs.single ? Math.round(left / 360) - 1 : Math.round(left / 720) - 1;
        return { k: fmt === "bullet" ? "ul" : "ol", l: Math.max(ilvl, Math.min(8, byIndent)) };
      }

      // ---- runs ----
      var images = [], imgCount = 0, chars = 0;
      function runMarks(rpr, m) {
        if (!rpr) return m;
        var o = Object.assign({}, m), e;
        if ((e = kid(rpr, "rStyle"))) { var cs = styles[w(e, "val")]; if (cs && cs.name !== "hyperlink") o = chainRun(cs.id, o); }
        return applyRun(rpr, o);
      }
      function chainRun(id, o) {
        var list = [], seen = 0;
        while (id && styles[id] && seen++ < 20) { list.unshift(styles[id]); id = styles[id].basedOn; }
        list.forEach(function (s) { if (s.rPr) o = applyRun(s.rPr, o); });
        return o;
      }
      function applyRun(rpr, o) {
        var e;
        if ((e = kid(rpr, "b"))) { if (onOff(e)) o.b = true; else delete o.b; }
        if ((e = kid(rpr, "i"))) { if (onOff(e)) o.i = true; else delete o.i; }
        if ((e = kid(rpr, "u"))) { if ((w(e, "val") || "single") !== "none") o.u = true; else delete o.u; }
        if ((e = kid(rpr, "strike")) || (e = kid(rpr, "dstrike"))) { if (onOff(e)) o.s = true; else delete o.s; }
        if ((e = kid(rpr, "color"))) { var c = hexOf(w(e, "val")); if (c && c !== "#000000") o.c = c; else delete o.c; }
        if ((e = kid(rpr, "highlight"))) { var hv = w(e, "val"); if (M.HL[hv]) o.h = M.HL[hv]; else delete o.h; }
        else if ((e = kid(rpr, "shd"))) { var f = hexOf(w(e, "fill")); if (f && f !== "#ffffff") o.h = f; }
        if ((e = kid(rpr, "sz"))) { var z = int(w(e, "val"), 2, 800, 22) / 2; if (Math.abs(z - base.size) > 0.01) o.z = z; else delete o.z; }
        if ((e = kid(rpr, "rFonts"))) { var fn = w(e, "ascii") || w(e, "hAnsi") || w(e, "cs"); if (fn && M.FONT_RE.test(fn) && fn !== base.font) o.f = fn; else if (fn === base.font) delete o.f; }
        if ((e = kid(rpr, "vertAlign"))) { var va = w(e, "val"); delete o.sup; delete o.sub; if (va === "superscript") o.sup = true; if (va === "subscript") o.sub = true; }
        if (kid(rpr, "vanish") && onOff(kid(rpr, "vanish"))) o.hidden = true; else delete o.hidden;
        return o;
      }
      // walk the inline content of a paragraph; calls out.text / out.image / out.pagebreak
      function inline(el, m, out, fld) {
        kids(el).forEach(function (c) {
          var n = c.localName;
          if (n === "r") return run(c, m, out, fld);
          if (n === "hyperlink") {
            var r = rel(c, "id"), u = r && rels[r] && rels[r].external ? M.safeUrl(rels[r].target || "") : null;
            return inline(c, u ? Object.assign({}, m, { a: u }) : m, out, fld);
          }
          if (n === "ins" || n === "smartTag" || n === "customXml" || n === "bdo" || n === "dir" || n === "moveTo") return inline(c, m, out, fld);
          if (n === "sdt") return inline(kid(c, "sdtContent"), m, out, fld);
          if (n === "fldSimple") {
            var ins = (w(c, "instr") || "").trim().toUpperCase();
            if (out.hf && /^PAGE\b/.test(ins)) return out.text("{page}", m);
            if (out.hf && /^NUMPAGES\b/.test(ins)) return out.text("{pages}", m);
            return inline(c, m, out, fld);
          }
          if (n === "AlternateContent") { var ch = kid(c, "Choice"), fb = kid(c, "Fallback"); return inline(ch && ch.getElementsByTagNameNS("*", "blip").length ? ch : fb || ch, m, out, fld); }
        });
      }
      function run(r, m, out, fld) {
        var mk = runMarks(kid(r, "rPr"), m);
        kids(r).forEach(function (c) {
          var n = c.localName;
          if (n === "fldChar") {
            var t = w(c, "fldCharType");
            if (t === "begin") fld.push({ instr: "", sep: false, skip: false });
            else if (t === "separate" && fld.length) {
              var f = fld[fld.length - 1], ins = f.instr.trim().toUpperCase();
              f.sep = true;
              if (out.hf && /^PAGE\b/.test(ins)) { out.text("{page}", mk); f.skip = true; }
              else if (out.hf && /^NUMPAGES\b/.test(ins)) { out.text("{pages}", mk); f.skip = true; }
            } else if (t === "end") fld.pop();
            return;
          }
          var cur = fld[fld.length - 1];
          if (n === "instrText") { if (cur) cur.instr += c.textContent; return; }
          if (cur && (!cur.sep || cur.skip)) return;
          if (mk.hidden) return;
          if (n === "t") return out.text(c.textContent, mk);
          if (n === "tab" || n === "ptab") return out.text("\t", mk);
          if (n === "br") { if (w(c, "type") === "page") return out.pagebreak(); if (w(c, "type") === "column") return out.text("\n", mk); return out.text("\n", mk); }
          if (n === "cr") return out.text("\n", mk);
          if (n === "noBreakHyphen") return out.text("‑", mk);
          if (n === "drawing" || n === "pict" || n === "object") return drawing(c, out);
        });
      }
      function drawing(el, out) {
        if (out.hf) return;
        var blips = el.getElementsByTagNameNS("*", "blip"), id = null;
        for (var k = 0; k < blips.length && !id; k++) id = rel(blips[k], "embed");
        if (!id) { var vi = el.getElementsByTagNameNS("*", "imagedata")[0]; if (vi) id = rel(vi, "id"); }
        if (!id || !rels[id] || rels[id].external || !rels[id].target) { if (blips.length || el.getElementsByTagNameNS("*", "imagedata").length) warnings.push("linked"); return; }
        var ext = el.getElementsByTagNameNS("*", "extent")[0], cx = ext ? +ext.getAttribute("cx") : 0, cy = ext ? +ext.getAttribute("cy") : 0;
        var wpx = isFinite(cx) && cx > 0 ? cx / 9525 : 300, hpx = isFinite(cy) && cy > 0 ? cy / 9525 : 200;
        var k2 = Math.min(1, 4000 / wpx, 6000 / hpx);
        wpx = Math.max(4, Math.round(wpx * k2)); hpx = Math.max(4, Math.round(hpx * k2));
        var dp = el.getElementsByTagNameNS("*", "docPr")[0], alt = dp ? (dp.getAttribute("descr") || dp.getAttribute("title") || "") : "";
        if (imgCount >= L.maxImages) { warnings.push("too many images"); return; }
        imgCount++;
        var bid = M.uid();
        images.push({ id: bid, path: rels[id].target });
        out.image(M.image(bid, wpx, hpx, { alt: alt.slice(0, 500) }));
      }

      // ---- paragraphs ----
      var blocks = [];
      function pProps(ppr) {
        var o = { style: "normal", align: "", dir: "", ind: 0, fi: "", list: null, pbb: false, styleId: null };
        var ps = kid(ppr, "pStyle"); o.styleId = ps ? w(ps, "val") : null;
        if (!o.styleId) Object.keys(styles).some(function (id) { if (styles[id].type === "paragraph" && styles[id].def) { o.styleId = id; return true; } return false; });
        o.style = mapStyle(o.styleId);
        function get(name) { var direct = kid(ppr, name); if (direct) return direct; return chain(o.styleId, function (s) { return s.pPr ? kid(s.pPr, name) : null; }); }
        var jc = get("jc"), jv = jc ? w(jc, "val") : "";
        o.align = jv === "center" ? "center" : jv === "right" || jv === "end" ? "right" : jv === "both" || jv === "distribute" || jv === "lowKashida" || jv === "mediumKashida" || jv === "highKashida" ? "justify" : jv === "left" || jv === "start" ? "left" : "";
        if (o.align === "left") o.align = "";
        var bidi = get("bidi"); if (bidi && onOff(bidi)) o.dir = "rtl";
        var np = get("numPr");
        if (np) {
          var il = kid(np, "ilvl"), ni = kid(np, "numId");
          if (!ni) ni = chain(o.styleId, function (s) { var q = s.pPr && kid(s.pPr, "numPr"); return q && kid(q, "numId"); });
          o.list = listFor(ni ? w(ni, "val") : null, il ? int(w(il, "val"), 0, 8, 0) : 0);
          // styles like "List Bullet 2" keep level 0 and indent instead: take the deeper of the two
          var li = get("ind"), lt = li ? int(w(li, "left") || w(li, "start"), 0, 31680, 0) : 0;
          if (o.list && lt) o.list.l = Math.max(o.list.l, Math.min(8, Math.round(lt / 720) - 1));
        }
        var ind = kid(ppr, "ind");
        if (ind && !o.list) {
          var left = int(w(ind, "left") || w(ind, "start"), -31680, 31680, 0), fl = int(w(ind, "firstLine"), 0, 31680, 0), hg = int(w(ind, "hanging"), 0, 31680, 0);
          if (hg >= 360) { o.fi = "hang"; left -= hg; } else if (fl >= 360) o.fi = "first";
          o.ind = Math.max(0, Math.min(8, Math.round(left / 720)));
          o.indTw = left + hg;
        }
        var sp = kid(ppr, "spacing");
        if (sp) {
          if (w(sp, "before") != null) o.sb = int(w(sp, "before"), 0, 4000, 0) / 20;
          if (w(sp, "after") != null) o.sa = int(w(sp, "after"), 0, 4000, 0) / 20;
          var ln = w(sp, "line"), rule = w(sp, "lineRule");
          if (ln != null) { var lv = int(ln, 1, 20000, 240); o.lh = !rule || rule === "auto" ? lv / 240 : Math.max(0.8, Math.min(4, (lv / 20) / (base.size * 1.17))); o.lh = Math.round(o.lh * 100) / 100; }
        }
        var pb = kid(ppr, "pageBreakBefore"); if (pb && onOff(pb)) o.pbb = true;
        return o;
      }
      function paragraph(p, target, hfMode) {
        var ppr = kid(p, "pPr"), pr = pProps(ppr), m = {};
        if (pr.style === "normal" && pr.styleId && styles[pr.styleId]) m = chainRun(pr.styleId, {});
        else if (pr.style === "normal" || pr.style === "toc") m = {};
        delete m.hidden;
        if (pr.pbb && target.length) target.push(M.pagebreak());
        var cur = null, out;
        function fresh() {
          cur = M.para(pr.style === "toc" ? "normal" : pr.style, null, { align: pr.align, dir: pr.dir, ind: pr.ind, fi: pr.fi, list: pr.list ? { k: pr.list.k, l: pr.list.l } : null });
          if (pr.sb != null) cur.sb = Math.min(200, pr.sb);
          if (pr.sa != null) cur.sa = Math.min(200, pr.sa);
          if (pr.lh) cur.lh = Math.max(0.5, Math.min(5, pr.lh));
          if (pr.style === "toc") cur._toc = true;
          cur.text = [];
        }
        fresh();
        var any = false, split = false;
        out = {
          hf: hfMode,
          text: function (t, mk) {
            if (chars > L.maxChars) return;
            t = String(t).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
            chars += t.length;
            var o = M.withText(mk, t); delete o.hidden; cur.text.push(o); any = true;
          },
          image: function (img) { flush(true); img.align = pr.align === "right" ? "right" : pr.align === "center" ? "center" : "left"; target.push(img); fresh(); split = true; },
          pagebreak: function () { flush(true); target.push(M.pagebreak()); fresh(); split = true; }
        };
        function flush(partial) {
          cur.text = M.cleanSpans(cur.text);
          if (!partial || M.len(cur.text)) target.push(cur);
        }
        inline(p, m, out, []);
        if (M.len(M.cleanSpans(cur.text)) || !split) flush(false);
        // checklist paragraphs written as "☐ text" / "☒ text"
        var b = target[target.length - 1];
        if (b && b.type === "p") {
          var t0 = M.plain(b.text), mm = /^([☐☒☑])\s?/.exec(t0);
          if (mm && (!b.list || b.list.k === "ul")) { b.text = M.slice(b.text, mm[0].length); b.list = { k: "todo", l: Math.max(0, Math.min(8, Math.round((pr.indTw || 720) / 720) - 1)), c: mm[1] !== "☐" }; b.ind = 0; b.fi = ""; }
        }
        return target;
      }
      function tableBlock(tbl, target) {
        var grid = kid(tbl, "tblGrid"), cws = kids(grid, "gridCol").map(function (g) { return int(w(g, "w"), 1, 100000, 1000); }).slice(0, 63);
        var rows = kids(tbl, "tr").slice(0, 1000), cols = Math.max(1, cws.length), grid2 = [], hdr = false, tops = {};
        rows.forEach(function (tr, r) {
          grid2[r] = [];
          var trPr = kid(tr, "trPr");
          if (r === 0 && trPr && kid(trPr, "tblHeader") && onOff(kid(trPr, "tblHeader"))) hdr = true;
          var c = 0, gb = trPr ? kid(trPr, "gridBefore") : null;
          if (gb) c = int(w(gb, "val"), 0, 62, 0);
          for (var k = 0; k < c; k++) grid2[r][k] = M.cell();
          kids(tr, "tc").forEach(function (tc) {
            if (c >= 63) return;
            var pr = kid(tc, "tcPr"), gs = pr ? kid(pr, "gridSpan") : null, vm = pr ? kid(pr, "vMerge") : null, shd = pr ? kid(pr, "shd") : null;
            var cs = Math.max(1, Math.min(63 - c, gs ? int(w(gs, "val"), 1, 63, 1) : 1));
            var cont = vm && (w(vm, "val") == null || w(vm, "val") === "continue");
            var cl = M.cell();
            if (cont && tops[c]) {
              tops[c].rs++;
              for (var k2 = c; k2 < c + cs; k2++) grid2[r][k2] = { text: [], bg: "", al: "", cs: 1, rs: 1, m: true };
              c += cs; return;
            }
            var parts = [], al = "";
            kids(tc).forEach(function (x) {
              if (x.localName === "p") { var tmp = paragraph(x, [], false); tmp.forEach(function (b) { if (b.type === "p") { if (!al && b.align) al = b.align; if (parts.length) parts.push({ t: "\n" }); parts = parts.concat(b.text); } }); }
              else if (x.localName === "tbl") { var inner = []; tableBlock(x, inner); inner.forEach(function (tb) { tb.rows.forEach(function (row) { if (parts.length) parts.push({ t: "\n" }); parts.push({ t: row.filter(function (q) { return !q.m; }).map(function (q) { return M.plain(q.text); }).join("  ") }); }); }); }
              else if (x.localName === "sdt") { kids(kid(x, "sdtContent"), "p").forEach(function (pp) { paragraph(pp, [], false).forEach(function (b) { if (b.type === "p") { if (parts.length) parts.push({ t: "\n" }); parts = parts.concat(b.text); } }); }); }
            });
            cl.text = M.cleanSpans(parts, 20000);
            cl.al = al === "justify" || al === "center" || al === "right" ? al : "";
            var bg = shd ? hexOf(w(shd, "fill")) : null; if (bg && bg !== "#ffffff") cl.bg = bg;
            cl.cs = cs;
            grid2[r][c] = cl;
            for (var k3 = c + 1; k3 < c + cs; k3++) grid2[r][k3] = { text: [], bg: "", al: "", cs: 1, rs: 1, m: true };
            if (vm) tops[c] = cl; else delete tops[c];
            c += cs;
          });
          cols = Math.max(cols, grid2[r].length);
        });
        cols = Math.min(63, cols);
        if (!rows.length) return;
        var t = M.table(grid2.length, cols);
        t.rows = grid2.map(function (row) { var o = []; for (var c = 0; c < cols; c++) o.push(row[c] || M.cell()); return o; });
        var sum = 0; for (var k = 0; k < cols; k++) sum += cws[k] || 1000;
        t.cw = []; for (k = 0; k < cols; k++) t.cw.push((cws[k] || 1000) / sum);
        t.hdr = hdr;
        if (hdr) t.rows[0].forEach(function (cl) { if (cl.bg === "#eef1f5") cl.bg = ""; if (M.len(cl.text) && M.hasMark(cl.text, 0, M.len(cl.text), "b")) cl.text = M.setMark(cl.text, 0, M.len(cl.text), "b", null); });
        target.push(M.cleanBlock(t));
      }
      function bodyInto(el, target) {
        kids(el).forEach(function (c) {
          if (target.length > L.maxBlocks) return;
          var n = c.localName;
          if (n === "p") paragraph(c, target, false);
          else if (n === "tbl") tableBlock(c, target);
          else if (n === "sdt") bodyInto(kid(c, "sdtContent"), target);
          else if (n === "customXml" || n === "ins") bodyInto(c, target);
        });
      }
      var body = dx.getElementsByTagNameNS("*", "body")[0];
      if (!body) throw new Error("This file has no document body.");
      bodyInto(body, blocks);
      if (blocks.length > L.maxBlocks) warnings.push("truncated");
      // contents entries become one live contents block
      var outB = [];
      blocks.forEach(function (b) {
        if (b._toc) { if (!outB.length || outB[outB.length - 1].type !== "toc") outB.push(M.toc()); return; }
        outB.push(b);
      });
      // ---- section: page size, margins, header / footer ----
      var sect = kids(body, "sectPr").pop() || body.getElementsByTagNameNS("*", "sectPr")[0];
      var page = { size: "a4", orient: "portrait", mt: 25.4, mr: 25.4, mb: 25.4, ml: 25.4, hd: 12.7, fd: 12.7 }, hf = { header: "", footer: "", ha: "center", fa: "center", first: false }, hfJobs = [];
      if (sect) {
        var sz = kid(sect, "pgSz"), mg = kid(sect, "pgMar");
        if (sz) {
          var wmm = int(w(sz, "w"), 2835, 68000, 11906) / TW, hmm = int(w(sz, "h"), 2835, 68000, 16838) / TW;
          var small = Math.min(wmm, hmm), big = Math.max(wmm, hmm), found = "custom";
          Object.keys(M.SIZES).forEach(function (k) { var s = M.SIZES[k]; if (Math.abs(s[0] - small) < 2 && Math.abs(s[1] - big) < 2) found = k; });
          page.size = found; page.orient = wmm > hmm ? "landscape" : "portrait"; page.w = wmm; page.h = hmm;
        }
        if (mg) ["top", "right", "bottom", "left", "header", "footer"].forEach(function (k, i) {
          var v = w(mg, k);
          if (v != null) page[["mt", "mr", "mb", "ml", "hd", "fd"][i]] = Math.max(0, Math.min(150, Math.abs(int(v, -85000, 85000, 1440)) / TW));
        });
        var tp = kid(sect, "titlePg"); var titlePg = tp && onOff(tp);
        kids(sect, "headerReference").concat(kids(sect, "footerReference")).forEach(function (ref) {
          var type = w(ref, "type") || "default", id = rel(ref, "id"), r2 = rels[id];
          if (!r2 || r2.external || !r2.target) return;
          hfJobs.push({ kind: ref.localName === "headerReference" ? "header" : "footer", type: type, path: r2.target });
        });
        hf.first = !!titlePg;
      }
      return Promise.all(hfJobs.map(function (j) {
        return text(j.path).then(function (s) {
          if (!s) return;
          var hx = xml(s, j.path), list = [];
          kids(hx.documentElement, "p").forEach(function (p) { paragraph(p, list, true); });
          kids(hx.documentElement, "sdt").forEach(function (x) { kids(kid(x, "sdtContent"), "p").forEach(function (p) { paragraph(p, list, true); }); });
          var tx = list.filter(function (b) { return b.type === "p"; }).map(function (b) { return M.plain(b.text).replace(/[\t\n]+/g, "   ").trim(); }).filter(Boolean).join("   ");
          var al = (list.filter(function (b) { return b.type === "p" && M.plain(b.text).trim(); })[0] || {}).align || "left";
          j.text = tx.slice(0, 500); j.align = al === "justify" ? "left" : al;
        }).catch(function () { warnings.push("header unreadable"); });
      })).then(function () {
        hfJobs.forEach(function (j) {
          if (j.type === "default" && j.text != null) { hf[j.kind] = j.text; hf[j.kind === "header" ? "ha" : "fa"] = j.align || "left"; }
        });
        if (hf.first) {
          var firstHas = hfJobs.some(function (j) { return j.type === "first" && j.text; });
          if (firstHas) hf.first = false; // a first-page header with its own text: show the normal one everywhere (one header per document here)
        }
        var title = "";
        var dc = null;
        return loadImages().then(function (imgs) {
          dc = M.cleanDoc({ page: page, hf: hf, base: base, blocks: outB, title: title });
          // drop image blocks whose bytes could not be used
          var ok = {}; imgs.forEach(function (x) { ok[x.id] = true; });
          dc.blocks = dc.blocks.filter(function (b) { return b.type !== "img" || ok[b.blob]; });
          M.normalize(dc);
          return { doc: dc, images: imgs, warnings: summarize(warnings) };
        });
      });

      function loadImages() {
        var total = 0, out = [];
        return images.reduce(function (pr, im) {
          return pr.then(function () {
            var f = zip.file(im.path);
            if (!f) { warnings.push("missing image"); return; }
            var u = f._data && f._data.uncompressedSize;
            if (typeof u === "number" && u > L.maxMedia) { warnings.push("image too large"); return; }
            return f.async("uint8array").then(function (u8) {
              if (u8.length > L.maxMedia || total + u8.length > L.maxTotal) { warnings.push("image too large"); return; }
              var type = sniff(u8);
              if (!type) { warnings.push("unsupported image"); return; }
              total += u8.length;
              out.push({ id: im.id, data: u8, type: type });
            });
          });
        }, Promise.resolve()).then(function () { return out; });
      }
    }
  }
  function summarize(ws) {
    var c = {}; ws.forEach(function (x) { c[x] = (c[x] || 0) + 1; });
    var msg = { linked: "linked (external) images were ignored", "unsupported image": "images in formats browsers can't show (EMF, WMF, TIFF…) were skipped", "missing image": "images were missing from the file", "image too large": "images were too large and skipped", "too many images": "images over the limit were skipped", truncated: "the document was very long and was cut short", "header unreadable": "a header or footer could not be read", "theme ignored": "the theme part was ignored" };
    return Object.keys(c).map(function (k) { return c[k] + " " + (msg[k] || k); });
  }

  window.DX = { exportDocx: exportDocx, importDocx: importDocx, sniff: sniff, LIMITS: LIMITS };
})();
