/* PowerPoint (.pptx) reader for the MCQ and SQ scrapers.
 *
 * A slide deck made from a Word file is the same text laid out page by page: text boxes hold the stems,
 * options, "উত্তর:" and "ব্যাখ্যা:" lines, and pictures float on top. This file turns every slide into
 * the paragraph stream the Word parsers already read (a synthetic <w:body>) and then calls them,
 * unchanged, so a .pptx gives the same sheet as the same content in a .docx.
 *
 *   slide  ->  text boxes / tables / pictures, put in reading order (full-width items, then the left
 *              column, then the right column, each top to bottom)
 *   a:p    ->  w:p   (italic, sub/superscript, line breaks, Word equations kept as OMML)
 *   a:tbl  ->  w:tbl
 *   p:pic  ->  a picture inside the paragraph it floats over, so option pictures land on their (ক)/(খ) label
 *
 * Works in the page and in Node (pass { JSZip, DOMParser } as deps).
 */
(function (root) {
  "use strict";

  const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
  const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
  const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
  const A14 = "http://schemas.microsoft.com/office/drawing/2010/main";
  const XML = "http://www.w3.org/XML/1998/namespace";
  const EMU = 12700; // EMU per point

  const kids = (el) => Array.from((el && el.children) || []);
  const local = (el) => (el && el.localName) || "";
  const child = (el, ns, name) => kids(el).find((c) => c.localName === name && c.namespaceURI === ns) || null;

  // resolve a relationship target against the folder of the part that owns it
  function resolve(target, baseDir) {
    if (target.startsWith("/")) return target.slice(1);
    const out = [];
    for (const p of (baseDir + target).split("/")) {
      if (p === "..") out.pop();
      else if (p !== ".") out.push(p);
    }
    return out.join("/");
  }
  async function readXml(zip, path, Parser) {
    const f = zip.file(path);
    return f ? new Parser().parseFromString(await f.async("string"), "application/xml") : null;
  }
  async function readRels(zip, partPath, Parser) {
    const dir = partPath.slice(0, partPath.lastIndexOf("/") + 1), name = partPath.slice(dir.length);
    const doc = await readXml(zip, `${dir}_rels/${name}.rels`, Parser);
    const rels = {};
    if (doc) {
      for (const r of Array.from(doc.getElementsByTagName("Relationship"))) {
        rels[r.getAttribute("Id")] = { type: r.getAttribute("Type") || "", target: resolve(r.getAttribute("Target"), dir) };
      }
    }
    return rels;
  }

  // ---------- geometry ----------
  // a shape's frame in its own coordinate space, or null when it inherits one from the layout
  function frameOf(el) {
    const tag = local(el);
    const t = tag === "graphicFrame" ? child(el, P, "xfrm")
      : tag === "grpSp" ? child(child(el, P, "grpSpPr"), A, "xfrm")
        : child(child(el, P, "spPr"), A, "xfrm");
    if (!t) return null;
    const off = child(t, A, "off"), ext = child(t, A, "ext");
    if (!off || !ext) return null;
    const f = { x: +off.getAttribute("x"), y: +off.getAttribute("y"), w: +ext.getAttribute("cx"), h: +ext.getAttribute("cy") };
    const co = child(t, A, "chOff"), ce = child(t, A, "chExt");
    if (co && ce && +ce.getAttribute("cx") && +ce.getAttribute("cy")) {
      f.chx = +co.getAttribute("x"); f.chy = +co.getAttribute("y");
      f.sx = f.w / +ce.getAttribute("cx"); f.sy = f.h / +ce.getAttribute("cy");
    }
    return f;
  }
  // frames of the enclosing groups (innermost first) carry a child -> parent mapping
  function absolute(frame, groups) {
    let { x, y, w, h } = frame;
    for (const g of groups) {
      if (g.sx === undefined) continue;
      x = g.x + (x - g.chx) * g.sx; y = g.y + (y - g.chy) * g.sy;
      w *= g.sx; h *= g.sy;
    }
    return { x: x / EMU, y: y / EMU, w: w / EMU, h: h / EMU }; // points
  }

  // ---------- slide contents ----------
  function* leaves(node, groups) {
    for (const c of kids(node)) {
      const t = local(c);
      if (t === "AlternateContent") {
        const pick = child(c, MC, "Choice") || child(c, MC, "Fallback");
        if (pick) yield* leaves(pick, groups);
      } else if (t === "sp" || t === "pic" || t === "graphicFrame") {
        yield { el: c, tag: t, groups };
      } else if (t === "grpSp") {
        const f = frameOf(c);
        yield* leaves(c, f ? [f, ...groups] : groups);
      }
    }
  }

  const SKIP_PLACEHOLDERS = new Set(["dt", "sldNum", "ftr", "hdr"]);
  const hasText = (el) => Array.from(el.getElementsByTagNameNS(A, "t")).some((t) => (t.textContent || "").trim()) ||
    el.getElementsByTagNameNS(M, "t").length > 0;

  function classify(leaf) {
    const { el, tag } = leaf;
    if (tag === "pic") return child(child(el, P, "blipFill"), A, "blip") ? "pic" : null;
    if (tag === "graphicFrame") return el.getElementsByTagNameNS(A, "tbl").length ? "table" : null;
    const ph = child(child(child(el, P, "nvSpPr"), P, "nvPr"), P, "ph");
    if (ph && SKIP_PLACEHOLDERS.has(ph.getAttribute("type"))) return null;
    return child(el, P, "txBody") && hasText(el) ? "text" : null;
  }

  // ---------- text boxes: vertical position of each paragraph (for placing floating pictures) ----------
  function paraStats(ap, boxW) {
    let chars = 0, tabs = 0, breaks = 0, sz = 0;
    const scan = (node) => {
      for (const c of kids(node)) {
        const t = local(c);
        if (t === "r" || t === "fld") {
          const rp = child(c, A, "rPr");
          if (rp && rp.getAttribute("sz")) sz = Math.max(sz, +rp.getAttribute("sz"));
          for (const x of kids(c)) if (local(x) === "t") { const s = x.textContent || ""; tabs += (s.match(/\t/g) || []).length; chars += s.replace(/\t/g, "").length; }
        } else if (t === "br") breaks++;
        else if (t === "m") chars += Array.from(c.getElementsByTagNameNS(M, "t")).reduce((n, x) => n + (x.textContent || "").length, 0);
        else if (t === "AlternateContent") { const p = child(c, MC, "Choice"); if (p) scan(p); }
      }
    };
    scan(ap);
    if (!sz) { const e = child(ap, A, "endParaRPr"); sz = e && e.getAttribute("sz") ? +e.getAttribute("sz") : 1000; }
    const pt = sz / 100;
    const width = chars * pt * 0.52 + tabs * 30;
    const lines = Math.max(1, Math.ceil(width / Math.max(40, boxW - 14))) + breaks;
    const pPr = child(ap, A, "pPr");
    const sp = (n) => { const v = child(child(pPr, A, n), A, "spcPts"); return v ? +v.getAttribute("val") / 100 : 0; };
    return { height: lines * pt * 1.28 + sp("spcBef") + sp("spcAft") };
  }
  // top of every paragraph, scaled so the paragraphs fill the box exactly (text boxes auto-fit their text)
  function paraTops(paras, box) {
    const hs = paras.map((p) => paraStats(p, box.w).height);
    const sum = hs.reduce((a, b) => a + b, 0) || 1;
    const f = Math.min(1.7, Math.max(0.55, box.h / sum));
    const tops = [], ends = [];
    let y = box.y;
    hs.forEach((h) => { tops.push(y); y += h * f; ends.push(y); });
    tops.ends = ends;
    return tops;
  }

  // ---------- building the Word-style body ----------
  function el(doc, name, attrs) {
    const e = doc.createElementNS(W, "w:" + name);
    for (const [k, v] of Object.entries(attrs || {})) e.setAttributeNS(W, "w:" + k, v);
    return e;
  }
  function textRun(doc, text, props) {
    const r = el(doc, "r");
    if (props) {
      const pr = el(doc, "rPr");
      if (props.i) pr.append(el(doc, "i"));
      if (props.b) pr.append(el(doc, "b"));
      if (props.va) pr.append(el(doc, "vertAlign", { val: props.va }));
      if (pr.children.length) r.append(pr);
    }
    const t = el(doc, "t");
    t.setAttributeNS(XML, "xml:space", "preserve");
    t.textContent = text;
    r.append(t);
    return r;
  }
  function runProps(rPr) {
    if (!rPr) return null;
    const base = +(rPr.getAttribute("baseline") || 0);
    return { i: rPr.getAttribute("i") === "1", b: rPr.getAttribute("b") === "1", va: base > 0 ? "superscript" : base < 0 ? "subscript" : "" };
  }
  function addInline(doc, parent, ap) {
    for (const c of kids(ap)) {
      const t = local(c);
      if (t === "r") {
        const tt = child(c, A, "t");
        if (tt) parent.append(textRun(doc, tt.textContent || "", runProps(child(c, A, "rPr"))));
      } else if (t === "fld") {
        const tt = child(c, A, "t");
        if (tt) parent.append(textRun(doc, tt.textContent || "", null));
      } else if (t === "br") {
        const r = el(doc, "r"); r.append(el(doc, "br")); parent.append(r);
      } else if (t === "m" && c.namespaceURI === A14) { // a Word equation: the same OMML Word uses
        for (const om of kids(c)) if (om.namespaceURI === M) parent.append(doc.importNode(om, true));
      } else if (t === "AlternateContent") {
        const pick = child(c, MC, "Choice");
        if (pick) addInline(doc, parent, pick);
      } else if (t === "Choice") addInline(doc, parent, c);
    }
  }
  function buildPara(doc, ap, heading) {
    const p = el(doc, "p");
    if (heading) { const pPr = el(doc, "pPr"); pPr.append(el(doc, "pStyle", { val: "Heading1" })); p.append(pPr); }
    addInline(doc, p, ap);
    return p;
  }
  function pictureRun(doc, blipFill, rid) {
    const r = el(doc, "r"), d = el(doc, "drawing");
    const bf = doc.importNode(blipFill, true); // keeps a:blip and its a:srcRect crop side by side
    const blip = child(bf, A, "blip");
    for (const a of Array.from(blip.attributes)) if (a.namespaceURI === R) blip.removeAttributeNS(R, a.localName);
    blip.setAttributeNS(R, "r:embed", rid);
    d.append(bf); r.append(d);
    return r;
  }
  function buildTable(doc, tbl) {
    const t = el(doc, "tbl"), grid = el(doc, "tblGrid");
    const cols = kids(child(tbl, A, "tblGrid")).map((g) => Math.round(+g.getAttribute("w") / 635));
    cols.forEach((w) => grid.append(el(doc, "gridCol", { w })));
    t.append(grid);
    for (const tr of kids(tbl).filter((x) => local(x) === "tr")) {
      const row = el(doc, "tr");
      let ci = 0;
      for (const tc of kids(tr).filter((x) => local(x) === "tc")) {
        const cell = el(doc, "tc"), pr = el(doc, "tcPr");
        const span = +(tc.getAttribute("gridSpan") || 1);
        pr.append(el(doc, "tcW", { w: cols.slice(ci, ci + span).reduce((a, b) => a + b, 0) || 1000, type: "dxa" }));
        if (span > 1) pr.append(el(doc, "gridSpan", { val: span }));
        if (tc.getAttribute("hMerge") === "1") { ci += 1; continue; } // covered by the cell to its left
        if (tc.getAttribute("vMerge") === "1") pr.append(el(doc, "vMerge"));
        else if (tc.getAttribute("rowSpan")) pr.append(el(doc, "vMerge", { val: "restart" }));
        cell.append(pr);
        const tx = child(tc, A, "txBody");
        const ps = tx ? kids(tx).filter((x) => local(x) === "p") : [];
        (ps.length ? ps : [null]).forEach((ap) => cell.append(ap ? buildPara(doc, ap, false) : el(doc, "p")));
        row.append(cell);
        ci += span;
      }
      t.append(row);
    }
    return t;
  }

  // reading order of one slide: full-width items split the page into bands; inside a band the left
  // column is read before the right one, each from top to bottom
  function readingOrder(items, pageW) {
    const full = (it) => it.box.w > pageW * 0.6;
    const byY = (a, b) => a.box.y - b.box.y || a.order - b.order;
    const out = [];
    const bands = [[]];
    for (const it of items.slice().sort(byY)) {
      if (full(it)) { out.push(...flush(bands.pop())); out.push(it); bands.push([]); } else bands[bands.length - 1].push(it);
    }
    out.push(...flush(bands.pop()));
    function flush(band) {
      const mid = pageW / 2;
      const colOf = (it) => (it.box.x + it.box.w / 2 < mid ? 0 : 1);
      return band.filter((i) => colOf(i) === 0).sort(byY).concat(band.filter((i) => colOf(i) === 1).sort(byY));
    }
    return out;
  }

  const LABEL = /\(\s*[কখগঘ]\s*\)/g;
  const LABEL_ANY = /\(\s*[কখগঘ]\s*\)/;

  // put a picture inside a paragraph, right after the option label it sits beside
  function injectAfterLabel(doc, wp, labelIndex, run) {
    let seen = 0;
    for (const r of kids(wp)) {
      if (local(r) !== "r") continue;
      const t = child(r, W, "t");
      if (!t) continue;
      const text = t.textContent;
      for (const m of text.matchAll(LABEL)) {
        if (seen++ === labelIndex) {
          const end = m.index + m[0].length;
          if (end < text.length) { // split the run after the label so the picture lands between the pieces
            const rest = r.cloneNode(true);
            child(rest, W, "t").textContent = text.slice(end);
            t.textContent = text.slice(0, end);
            r.after(rest);
          }
          r.after(run);
          return true;
        }
      }
    }
    return false;
  }

  /**
   * Build the synthetic Word body for a whole deck.
   * Returns { doc, body, rels, slides } where rels maps picture ids to their path inside the .pptx.
   */
  async function buildBody(zip, Parser, progress) {
    const pres = await readXml(zip, "ppt/presentation.xml", Parser);
    if (!pres) throw new Error("This is not a PowerPoint (.pptx) file.");
    const presRels = await readRels(zip, "ppt/presentation.xml", Parser);
    const sz = pres.getElementsByTagNameNS(P, "sldSz")[0];
    const pageW = sz ? +sz.getAttribute("cx") / EMU : 612;
    const slidePaths = Array.from(pres.getElementsByTagNameNS(P, "sldId"))
      .map((s) => presRels[s.getAttributeNS(R, "id")]).filter(Boolean).map((r) => r.target);

    const doc = new Parser().parseFromString(
      `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:m="${M}" xmlns:a="${A}"><w:body/></w:document>`, "application/xml");
    const body = doc.getElementsByTagNameNS(W, "body")[0];
    const rels = {};
    const layoutChapter = {}; // layout path -> chapter text (the page header text lives in the layout)
    const parseChapter = root.MCQCore && root.MCQCore.parseChapter;

    for (let si = 0; si < slidePaths.length; si++) {
      if (si % 5 === 0 && progress) await progress(`Reading slide ${si + 1} of ${slidePaths.length}…`);
      const sdoc = await readXml(zip, slidePaths[si], Parser);
      if (!sdoc) continue;
      const srels = await readRels(zip, slidePaths[si], Parser);
      const tree = sdoc.getElementsByTagNameNS(P, "spTree")[0];
      if (!tree) continue;

      // chapter from the slide's layout
      const lay = Object.values(srels).find((r) => /\/slideLayout$/.test(r.type));
      let chapter = "";
      if (lay) {
        if (!(lay.target in layoutChapter)) {
          const ld = await readXml(zip, lay.target, Parser);
          const text = ld ? Array.from(ld.getElementsByTagNameNS(A, "p")).map((p) => Array.from(p.getElementsByTagNameNS(A, "t")).map((t) => t.textContent).join("")).join(" ") : "";
          layoutChapter[lay.target] = parseChapter ? parseChapter(text) : "";
        }
        chapter = layoutChapter[lay.target];
      }

      // collect the items of this slide
      const items = [];
      let order = 0;
      for (const leaf of leaves(tree, [])) {
        const kind = classify(leaf);
        if (!kind) continue;
        const f = frameOf(leaf.el);
        const box = f ? absolute(f, leaf.groups) : { x: 0, y: items.length ? items[items.length - 1].box.y : 0, w: pageW, h: 20 };
        const it = { kind, el: leaf.el, box, order: order++ };
        if (kind === "pic") {
          const bf = child(leaf.el, P, "blipFill");
          const rid = child(bf, A, "blip").getAttributeNS(R, "embed");
          const target = srels[rid] && srels[rid].target;
          if (!target) continue;
          it.id = `s${si}_${rid}`;
          rels[it.id] = target;
          it.blipFill = bf;
        } else if (kind === "text") {
          it.paras = kids(child(leaf.el, P, "txBody")).filter((c) => local(c) === "p");
          const ph = child(child(child(leaf.el, P, "nvSpPr"), P, "nvPr"), P, "ph");
          it.heading = !!ph && /^(title|ctrTitle)$/.test(ph.getAttribute("type") || "");
        }
        items.push(it);
      }

      // pictures that float over a text box go into that box's paragraphs
      const boxes = items.filter((i) => i.kind === "text");
      boxes.forEach((b) => { b.tops = paraTops(b.paras, b.box); b.labelRow = b.paras.map((p) => LABEL_ANY.test(Array.from(p.getElementsByTagNameNS(A, "t")).map((t) => t.textContent).join(""))); b.wps = b.paras.map((p) => buildPara(doc, p, b.heading)); b.inlinePics = b.paras.map(() => []); });
      const standalone = [];
      for (const pic of items.filter((i) => i.kind === "pic")) {
        let best = null, bestOv = 0;
        for (const b of boxes) {
          const ov = Math.min(pic.box.x + pic.box.w, b.box.x + b.box.w) - Math.max(pic.box.x, b.box.x);
          const inside = pic.box.y >= b.box.y - 6 && pic.box.y <= b.box.y + b.box.h + 6;
          if (inside && ov / Math.max(1, pic.box.w) >= 0.6 && ov > bestOv) { best = b; bestOv = ov; }
        }
        if (!best || pic.box.w > pageW * 0.6) { standalone.push(pic); continue; } // banners and the like
        // option pictures are centred on their label's line: prefer an option-label row the centre falls on
        // (a little slack either side, as paragraph positions are estimates); otherwise the paragraph under the centre
        const centre = pic.box.y + Math.min(pic.box.h / 2, 24);
        let k = -1;
        best.tops.forEach((t, i) => {
          if (best.labelRow[i] && centre >= t - 4 && centre <= best.tops.ends[i] + 6 && (k < 0 || Math.abs(centre - (t + best.tops.ends[i]) / 2) < Math.abs(centre - (best.tops[k] + best.tops.ends[k]) / 2))) k = i;
        });
        if (k < 0) { k = 0; best.tops.forEach((t, i) => { if (t <= centre + 1) k = i; }); }
        best.inlinePics[k].push(pic);
      }
      for (const b of boxes) {
        b.inlinePics.forEach((pics, k) => {
          if (!pics.length) return;
          const wp = b.wps[k];
          const labels = (b.paras[k] ? Array.from(b.paras[k].getElementsByTagNameNS(A, "t")).map((t) => t.textContent).join("") : "").match(LABEL) || [];
          const sorted = pics.slice().sort((p, q) => p.box.x - q.box.x);
          const placed = new Set();
          if (labels.length) {
            // a row of option labels: each picture goes after the label to its left
            const rowStart = b.box.x, rowW = b.box.w;
            sorted.forEach((pic) => {
              const rel = (pic.box.x + pic.box.w / 2 - rowStart) / rowW;
              const li = labels.length === 1 ? 0 : Math.min(labels.length - 1, Math.floor(rel * labels.length));
              if (injectAfterLabel(doc, wp, li, pictureRun(doc, pic.blipFill, pic.id))) placed.add(pic);
            });
          }
          const rest = sorted.filter((p) => !placed.has(p));
          if (rest.length) { // a figure under the stem: a paragraph of its own, right after the paragraph it floats over
            const holder = el(doc, "p");
            rest.forEach((pic) => holder.append(pictureRun(doc, pic.blipFill, pic.id)));
            b.after = b.after || {};
            b.after[k] = holder;
          }
        });
      }

      // emit in reading order
      const emitted = [];
      const ordered = readingOrder(items.filter((i) => i.kind !== "pic" || standalone.includes(i)), pageW);
      for (const it of ordered) {
        if (it.kind === "text") {
          it.wps.forEach((wp, k) => { emitted.push(wp); if (it.after && it.after[k]) emitted.push(it.after[k]); });
        } else if (it.kind === "table") {
          emitted.push(buildTable(doc, it.el.getElementsByTagNameNS(A, "tbl")[0]));
        } else if (it.kind === "pic") {
          const p = el(doc, "p"); p.append(pictureRun(doc, it.blipFill, it.id)); emitted.push(p);
        }
      }
      emitted.forEach((n) => body.append(n));

      // a section break after every slide carries that slide's chapter
      const sp = el(doc, "sectPr");
      sp.setAttribute("data-chapter", chapter);
      const brk = el(doc, "p"), pPr = el(doc, "pPr");
      pPr.append(sp); brk.append(pPr); body.append(brk);
    }
    body.append(el(doc, "sectPr"));
    return { doc, body, rels, slides: slidePaths.length };
  }

  /**
   * Extract questions from a .pptx.
   *   data     ArrayBuffer / Uint8Array of the .pptx
   *   kind     "mcq" (default) or "sq"
   *   encode   async (entries) => data URL (MCQ pictures; see images.js)
   */
  async function extract(data, { kind, encode, onProgress, deps } = {}) {
    const JSZip = (deps && deps.JSZip) || root.JSZip;
    const Parser = (deps && deps.DOMParser) || root.DOMParser;
    const progress = onProgress || (() => {});
    progress("Reading presentation…");
    const zip = await JSZip.loadAsync(data);
    const { body, rels, slides } = await buildBody(zip, Parser, progress);
    if (!slides) throw new Error("This presentation has no slides.");
    const chapterOf = (sp) => sp.getAttribute("data-chapter") || "";
    const styleNames = {};

    progress("Finding questions…");
    if (kind === "sq") {
      const found = root.SQParser.parse(body, rels, styleNames, chapterOf);
      return { mcqs: found.items, skipped: 0, keys: 0, slides, rows: root.SQParser.toRows(found.items) };
    }
    const { mcqs, registry, skipped, keys } = root.MCQParser.parse(body, rels, styleNames, chapterOf);
    const entry = (i) => registry[i].group
      ? { group: registry[i].group, rels: registry[i].rels, getBytes: async (target) => zip.file(target).async("uint8array") }
      : registry[i].table ? { table: registry[i].table } : ({
        crop: registry[i].crop,
        getBytes: async () => zip.file(registry[i].target).async("uint8array"),
      });
    for (let n = 0; n < mcqs.length; n++) {
      const q = mcqs[n];
      q.image = q.questionImages.length && encode ? await encode(q.questionImages.map(entry)) : "";
      q.optionImageUrls = {};
      for (const [letter, idxs] of Object.entries(q.optionImages)) {
        q.optionImageUrls[letter] = encode ? await encode(idxs.map(entry)) : "";
      }
      if (n % 25 === 0) await progress(`Processing question ${n + 1} of ${mcqs.length}…`);
    }
    return { mcqs, skipped, keys, slides, rows: root.MCQParser.toRows(mcqs) };
  }

  const api = { extract, buildBody, readingOrder, _paraTops: paraTops, _paraStats: paraStats };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.PPTXCore = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
