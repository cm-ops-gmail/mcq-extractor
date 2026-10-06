/* Browser-only picture handling: crop the white margin, shrink, and encode as a base64 PNG URL
 * that fits in one Excel cell (limit 32,767 characters). */
(function (root) {
  "use strict";

  const MAX_CELL_CHARS = 32000;

  const mime = (bytes) => {
    if (bytes[0] === 0x89 && bytes[1] === 0x50) return "image/png";
    if (bytes[0] === 0xff && bytes[1] === 0xd8) return "image/jpeg";
    if (bytes[0] === 0x47 && bytes[1] === 0x49) return "image/gif";
    if (bytes[0] === 0x42 && bytes[1] === 0x4d) return "image/bmp";
    if (bytes[0] === 0x52 && bytes[1] === 0x49) return "image/webp";
    return "";
  };

  async function decode(bytes) {
    try {
      const type = mime(bytes);
      if (!type) return null; // EMF/WMF and other formats a browser cannot draw
      return await createImageBitmap(new Blob([bytes], { type }));
    } catch (e) {
      return null;
    }
  }

  function newCanvas(w, h) {
    const c = document.createElement("canvas");
    c.width = Math.max(1, Math.round(w));
    c.height = Math.max(1, Math.round(h));
    return c;
  }

  // draw on white (transparent PNGs), applying Word's own crop (fractions l, t, r, b)
  function toCanvas(bitmap, crop) {
    let sx = 0, sy = 0, sw = bitmap.width, sh = bitmap.height;
    if (crop && crop.some((v) => v)) {
      const [l, t, r, b] = crop;
      const x0 = Math.round(sw * l), y0 = Math.round(sh * t);
      const x1 = Math.round(sw * (1 - r)), y1 = Math.round(sh * (1 - b));
      if (x1 > x0 && y1 > y0) { sx = x0; sy = y0; sw = x1 - x0; sh = y1 - y0; }
    }
    const c = newCanvas(sw, sh);
    const g = c.getContext("2d");
    g.fillStyle = "#fff";
    g.fillRect(0, 0, c.width, c.height);
    g.drawImage(bitmap, sx, sy, sw, sh, 0, 0, c.width, c.height);
    return c;
  }

  // crop the white margin around the picture
  function trim(canvas, pad = 6, tol = 12) {
    const { width: w, height: h } = canvas;
    const d = canvas.getContext("2d").getImageData(0, 0, w, h).data;
    let l = w, t = h, r = -1, b = -1;
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        const i = (y * w + x) * 4;
        if (255 - Math.min(d[i], d[i + 1], d[i + 2]) > tol) {
          if (x < l) l = x;
          if (x > r) r = x;
          if (y < t) t = y;
          if (y > b) b = y;
        }
      }
    }
    if (r < 0) return canvas;
    l = Math.max(l - pad, 0); t = Math.max(t - pad, 0);
    r = Math.min(r + 1 + pad, w); b = Math.min(b + 1 + pad, h);
    const out = newCanvas(r - l, b - t);
    out.getContext("2d").drawImage(canvas, l, t, r - l, b - t, 0, 0, r - l, b - t);
    return out;
  }

  function stack(canvases) {
    if (canvases.length === 1) return canvases[0];
    const gap = 10;
    const out = newCanvas(Math.max(...canvases.map((c) => c.width)),
      canvases.reduce((s, c) => s + c.height, 0) + gap * (canvases.length - 1));
    const g = out.getContext("2d");
    g.fillStyle = "#fff";
    g.fillRect(0, 0, out.width, out.height);
    let y = 0;
    for (const c of canvases) { g.drawImage(c, 0, y); y += c.height + gap; }
    return out;
  }

  // fewer colours compress much better as PNG
  function posterize(canvas, step) {
    const g = canvas.getContext("2d");
    const img = g.getImageData(0, 0, canvas.width, canvas.height);
    const d = img.data;
    for (let i = 0; i < d.length; i += 4) {
      d[i] = Math.min(255, Math.round(d[i] / step) * step);
      d[i + 1] = Math.min(255, Math.round(d[i + 1] / step) * step);
      d[i + 2] = Math.min(255, Math.round(d[i + 2] / step) * step);
    }
    g.putImageData(img, 0, 0);
    return canvas;
  }

  function toDataUrl(canvas) {
    for (const side of [900, 700, 500, 400, 300, 220, 160, 120, 80]) {
      const scale = Math.min(1, side / Math.max(canvas.width, canvas.height));
      for (const step of [0, 32, 64]) {
        const c = newCanvas(canvas.width * scale, canvas.height * scale);
        const g = c.getContext("2d");
        g.imageSmoothingQuality = "high";
        g.drawImage(canvas, 0, 0, c.width, c.height);
        if (step) posterize(c, step);
        const url = c.toDataURL("image/png");
        if (url.length <= MAX_CELL_CHARS) return url;
      }
    }
    return "";
  }

  // ---------- Word tables drawn as a picture ----------
  const TW = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const kidsOf = (el) => Array.from(el.children || []);
  const named = (el, n) => kidsOf(el).find((c) => c.localName === n && c.namespaceURI === TW) || null;
  const attrW = (el, n) => (el ? el.getAttributeNS(TW, n) : null);

  function cellInfo(tc) {
    const paras = kidsOf(tc).filter((c) => c.localName === "p");
    const lines = paras.map((p) => (root.MCQParser ? root.MCQParser.paraText(p, {}, null) : p.textContent).replace(/\t/g, " ").trim());
    const bold = paras.some((p) => Array.from(p.getElementsByTagNameNS(TW, "b")).some((b) => attrW(b, "val") !== "0" && attrW(b, "val") !== "false"));
    const jc = paras.map((p) => attrW(named(named(p, "pPr"), "jc"), "val")).find(Boolean) || "left";
    const pr = named(tc, "tcPr");
    const shd = named(pr, "shd");
    const fill = shd && attrW(shd, "fill") && attrW(shd, "fill") !== "auto" ? "#" + attrW(shd, "fill") : "";
    const vm = named(pr, "vMerge");
    return {
      text: lines.join("\n").replace(/\n{2,}/g, "\n").trim(),
      bold, align: jc === "center" ? "center" : jc === "right" || jc === "end" ? "right" : "left", fill,
      span: parseInt(attrW(named(pr, "gridSpan"), "val") || "1", 10) || 1,
      vmerge: vm ? (attrW(vm, "val") === "restart" ? "restart" : "continue") : "",
    };
  }

  // Draw a Word table (grid, merged cells, shading, bold, alignment) onto a white canvas.
  function drawTable(tbl) {
    const S = 2, PAD = 8, LINE = 20, FONT = '15px "Noto Sans Bengali","Geist",system-ui,sans-serif';
    const grid = named(tbl, "tblGrid");
    let cols = grid ? kidsOf(grid).filter((c) => c.localName === "gridCol").map((c) => Math.max(40, (parseInt(attrW(c, "w"), 10) || 1500) / 15)) : [];
    const rows = kidsOf(tbl).filter((r) => r.localName === "tr").map((tr) => {
      let col = 0;
      return kidsOf(tr).filter((c) => c.localName === "tc").map((tc) => { const ci = cellInfo(tc); ci.col = col; col += ci.span; return ci; });
    });
    const ncols = Math.max(cols.length, ...rows.map((r) => (r.length ? r[r.length - 1].col + r[r.length - 1].span : 0)));
    while (cols.length < ncols) cols.push(100);
    const total = cols.reduce((a, b) => a + b, 0);
    if (total > 900) cols = cols.map((w) => (w * 900) / total); // keep very wide tables readable
    const xs = [0]; cols.forEach((w) => xs.push(xs[xs.length - 1] + w));

    const meas = document.createElement("canvas").getContext("2d");
    const wrap = (cell, width) => {
      meas.font = (cell.bold ? "600 " : "") + FONT;
      const out = [];
      for (const para of cell.text.split("\n")) {
        let line = "";
        for (const word of para.split(/(\s+)/)) {
          const t = line + word;
          if (line && meas.measureText(t).width > width - 2 * PAD) { out.push(line.trimEnd()); line = word.trimStart(); } else line = t;
        }
        out.push(line.trimEnd());
      }
      return out;
    };
    // row heights come from cells that are not vertically merged
    const heights = rows.map((r) => Math.max(30, ...r.filter((c) => c.vmerge !== "continue" && c.text).map((c) => wrap(c, xs[Math.min(c.col + c.span, ncols)] - xs[c.col]).length * LINE + 2 * PAD - 4)));
    const ys = [0]; heights.forEach((h) => ys.push(ys[ys.length - 1] + h));

    const canvas = newCanvas((xs[ncols] + 2) * S, (ys[rows.length] + 2) * S);
    const g = canvas.getContext("2d");
    g.fillStyle = "#fff"; g.fillRect(0, 0, canvas.width, canvas.height);
    g.scale(S, S); g.translate(1, 1);
    g.textBaseline = "middle"; g.lineWidth = 1; g.strokeStyle = "#222";
    rows.forEach((r, ri) => {
      for (const c of r) {
        if (c.vmerge === "continue") continue;
        let rs = 1;
        if (c.vmerge === "restart") for (let k = ri + 1; k < rows.length; k++) { const below = rows[k].find((x) => x.col === c.col); if (below && below.vmerge === "continue") rs++; else break; }
        const x0 = xs[c.col], x1 = xs[Math.min(c.col + c.span, ncols)], y0 = ys[ri], y1 = ys[ri + rs];
        if (c.fill) { g.fillStyle = c.fill; g.fillRect(x0, y0, x1 - x0, y1 - y0); }
        g.strokeRect(x0, y0, x1 - x0, y1 - y0);
        if (!c.text) continue;
        const lines = wrap(c, x1 - x0);
        g.font = (c.bold ? "600 " : "") + FONT; g.fillStyle = "#000";
        g.textAlign = c.align;
        const tx = c.align === "center" ? (x0 + x1) / 2 : c.align === "right" ? x1 - PAD : x0 + PAD;
        const top = y0 + (y1 - y0 - lines.length * LINE) / 2 + LINE / 2;
        lines.forEach((ln, li) => g.fillText(ln, tx, top + li * LINE));
      }
    });
    return canvas;
  }

  // entries: [{ crop, getBytes } | { table }] -> data URL ("" when nothing could be decoded)
  async function encode(entries) {
    const parts = [];
    for (const e of entries) {
      if (e.table) { try { parts.push(trim(drawTable(e.table))); } catch (err) { console.warn("table skipped", err); } continue; }
      const bitmap = await decode(await e.getBytes());
      if (bitmap) parts.push(trim(toCanvas(bitmap, e.crop)));
    }
    return parts.length ? toDataUrl(stack(parts)) : "";
  }

  root.MCQImages = { encode };
})(typeof globalThis !== "undefined" ? globalThis : this);
