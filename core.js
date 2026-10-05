/* Glue between the .docx container, the parser and the .xlsx writer. Works in the page and in Node. */
(function (root) {
  "use strict";

  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

  // resolve a relationship target against the folder of the part that owns the relationships
  function resolveTarget(target, baseDir) {
    if (target.startsWith("/")) return target.slice(1);
    const parts = ((baseDir === undefined ? "word/" : baseDir) + target).split("/");
    const out = [];
    for (const p of parts) {
      if (p === "..") out.pop();
      else if (p !== ".") out.push(p);
    }
    return out.join("/");
  }

  async function readXml(zip, path, Parser) {
    const f = zip.file(path);
    return f ? new Parser().parseFromString(await f.async("string"), "application/xml") : null;
  }

  // "রসায়ন প্রথম পত্র ▶ অধ্যায় ০ ১ ▶ ল্যাবরেটরির নিরাপদ ব্যবহার"  ->  "অধ্যায় ০১: ল্যাবরেটরির নিরাপদ ব্যবহার"
  function parseChapter(text) {
    // NFD so "য়" is "য" + nukta whichever way the document stored it
    const m = /অ\s*ধ্যা\s*য\u09BC?\s*((?:[০-৯0-9]\s*)+)[\s▶►▸▷>:|\-–—]*(.*)$/u.exec(text.normalize("NFD"));
    if (!m) return "";
    // no leading zeros: "০১" -> "১" (a lone "০" stays)
    const num = m[1].replace(/\s+/g, "").replace(/^[০0]+(?=[০-৯1-9])/, "");
    const name = m[2].normalize("NFC").replace(/\s+/g, " ").replace(/\s*[0-9০-৯]+$/, "").trim();
    return name ? `অধ্যায় ${num}: ${name}` : `অধ্যায় ${num}`;
  }

  /**
   * Extract MCQs from a .docx.
   *   data     ArrayBuffer / Uint8Array of the .docx
   *   encode   async (entries) => data URL; entries = [{ crop, getBytes: async () => Uint8Array }]
   *   deps     { JSZip, DOMParser } (defaults to the page globals)
   */
  async function extract(data, { encode, onProgress, deps, kind } = {}) {
    const JSZip = (deps && deps.JSZip) || root.JSZip;
    const Parser = (deps && deps.DOMParser) || root.DOMParser;
    const progress = onProgress || (() => {});
    const zip = await JSZip.loadAsync(data);

    progress("Reading document…");
    // .docx and .docm share one layout; find the main part through the package relationships
    let mainPath = "word/document.xml";
    const pkgRels = await readXml(zip, "_rels/.rels", Parser);
    if (pkgRels) {
      const r = Array.from(pkgRels.getElementsByTagName("Relationship")).find((x) => /\/officeDocument$/.test(x.getAttribute("Type") || ""));
      if (r) mainPath = resolveTarget(r.getAttribute("Target"), "");
    }
    const mainDir = mainPath.includes("/") ? mainPath.slice(0, mainPath.lastIndexOf("/") + 1) : "";
    const doc = await readXml(zip, mainPath, Parser);
    if (!doc) throw new Error("This is not a Word (.docx or .docm) file.");
    const relsDoc = await readXml(zip, `${mainDir}_rels/${mainPath.slice(mainDir.length)}.rels`, Parser);
    const rels = {};
    if (relsDoc) {
      for (const r of Array.from(relsDoc.getElementsByTagName("Relationship"))) {
        rels[r.getAttribute("Id")] = resolveTarget(r.getAttribute("Target"), mainDir);
      }
    }
    const stylesDoc = await readXml(zip, `${mainDir}styles.xml`, Parser);
    const styleNames = {};
    if (stylesDoc) {
      for (const s of Array.from(stylesDoc.getElementsByTagNameNS(W, "style"))) {
        const name = Array.from(s.children).find((c) => c.localName === "name");
        styleNames[s.getAttributeNS(W, "styleId")] = name ? name.getAttributeNS(W, "val") : "";
      }
    }

    // chapter name of every page header, by relationship id
    const headerChapters = {};
    if (relsDoc) {
      for (const r of Array.from(relsDoc.getElementsByTagName("Relationship"))) {
        if (!/\/header$/.test(r.getAttribute("Type") || "")) continue;
        const hdr = await readXml(zip, resolveTarget(r.getAttribute("Target"), mainDir), Parser);
        if (!hdr) continue;
        const text = Array.from(hdr.getElementsByTagNameNS(W, "p"))
          .map((p) => Array.from(p.getElementsByTagNameNS(W, "t")).map((t) => t.textContent).join(""))
          .join(" ");
        headerChapters[r.getAttribute("Id")] = parseChapter(text);
      }
    }
    const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const order = { default: 0, even: 1, first: 2 };
    const chapterOf = (sectPr) => {
      const refs = Array.from(sectPr.children).filter((c) => c.localName === "headerReference")
        .sort((a, b) => (order[a.getAttributeNS(W, "type")] ?? 3) - (order[b.getAttributeNS(W, "type")] ?? 3));
      for (const ref of refs) {
        const ch = headerChapters[ref.getAttributeNS(R, "id")];
        if (ch) return ch;
      }
      return "";
    };

    progress("Finding MCQs…");
    const body = doc.getElementsByTagNameNS(W, "body")[0];
    if (kind === "sq") { // short questions: no pictures or options, so no image step
      const found = root.SQParser.parse(body, rels, styleNames, chapterOf);
      return { mcqs: found.items, skipped: 0, keys: 0, rows: root.SQParser.toRows(found.items) };
    }
    const { mcqs, registry, skipped, keys } = root.MCQParser.parse(body, rels, styleNames, chapterOf);

    const entry = (i) => ({
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
      if (n % 25 === 0) progress(`Processing question ${n + 1} of ${mcqs.length}…`);
    }
    return { mcqs, skipped, keys, rows: root.MCQParser.toRows(mcqs) };
  }

  // spec (optional): { sheet, headers, widths, imageCols }; defaults are the MCQ sheet
  async function buildWorkbook(ExcelJS, rows, spec) {
    spec = spec || {};
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet(spec.sheet || "MCQs");
    ws.addRow(spec.headers || root.MCQParser.HEADERS);
    rows.forEach((r) => ws.addRow(r));
    const widths = spec.widths || [8, 6, 8, 30, 14, 45, 12, 70];
    ws.columns.forEach((col, i) => (col.width = i < widths.length ? widths[i] : 18));
    const head = ws.getRow(1);
    head.eachCell((c) => {
      c.font = { bold: true };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9EAD3" } };
      c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    });
    ws.eachRow((row, n) => {
      if (n > 1) row.eachCell({ includeEmpty: true }, (c) => (c.alignment = { vertical: "middle", wrapText: true }));
    });
    // picture cells (question image, then one per option): links become clickable hyperlinks (the cell text
    // stays the URL) and these cells do not wrap, so a long link or base64 text does not make the row huge
    for (const col of spec.imageCols || [7, 11, 14, 17, 20]) {
      ws.getColumn(col).eachCell((c, n) => {
        if (n === 1) return;
        c.alignment = { vertical: "middle", wrapText: false };
        if (typeof c.value === "string" && /^https?:\/\//.test(c.value)) {
          c.value = { text: c.value, hyperlink: c.value };
          c.font = { color: { argb: "FF0563C1" }, underline: true };
        }
      });
    }
    ws.views = [{ state: "frozen", ySplit: 1 }];
    return wb;
  }

  const api = { extract, buildWorkbook, parseChapter };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.MCQCore = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
