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
    for (const side of [500, 400, 300, 220, 160, 120, 80]) {
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

  // entries: [{ crop, getBytes }] -> data URL ("" when nothing could be decoded)
  async function encode(entries) {
    const parts = [];
    for (const e of entries) {
      const bitmap = await decode(await e.getBytes());
      if (bitmap) parts.push(trim(toCanvas(bitmap, e.crop)));
    }
    return parts.length ? toDataUrl(stack(parts)) : "";
  }

  root.MCQImages = { encode };
})(typeof globalThis !== "undefined" ? globalThis : this);
