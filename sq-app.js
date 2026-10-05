(function () {
  "use strict";
  const $ = (id) => document.getElementById(id);
  const app = $("app"), drop = $("drop"), fileInput = $("file"), errorBox = $("error");
  const reduceMotion = window.matchMedia && matchMedia("(prefers-reduced-motion: reduce)").matches;
  let workbookBlob = null, downloadUrl = null, outName = "sq.xlsx", busy = false, flashTimer = null;

  const tick = () => new Promise((r) => setTimeout(r, 0));
  const setState = (s) => { app.dataset.state = s; };

  // ---------- theme ----------
  const root = document.documentElement;
  const prefersDark = () => matchMedia("(prefers-color-scheme: dark)").matches;
  const applyTheme = (t) => { root.classList.toggle("dark", t === "dark"); root.classList.toggle("light", t === "light"); };
  try { const t = localStorage.getItem("theme"); if (t) applyTheme(t); } catch (e) { /* storage unavailable */ }
  $("theme").addEventListener("click", () => {
    const dark = root.classList.contains("dark") || (!root.classList.contains("light") && prefersDark());
    const next = dark ? "light" : "dark";
    applyTheme(next);
    try { localStorage.setItem("theme", next); } catch (e) { /* ignore */ }
  });

  // ---------- upload ----------
  drop.addEventListener("click", () => fileInput.click());
  drop.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); fileInput.click(); } });
  let dragDepth = 0;
  window.addEventListener("dragenter", (e) => { e.preventDefault(); dragDepth++; drop.classList.add("over"); });
  window.addEventListener("dragover", (e) => e.preventDefault());
  window.addEventListener("dragleave", () => { if (--dragDepth <= 0) { dragDepth = 0; drop.classList.remove("over"); } });
  window.addEventListener("drop", (e) => {
    e.preventDefault(); dragDepth = 0; drop.classList.remove("over");
    if (e.dataTransfer.files[0]) handle(e.dataTransfer.files[0]);
  });
  fileInput.addEventListener("change", () => { if (fileInput.files[0]) handle(fileInput.files[0]); });

  const forgetDownload = () => {
    workbookBlob = null;
    if (downloadUrl) URL.revokeObjectURL(downloadUrl);
    downloadUrl = null;
  };
  $("again").addEventListener("click", () => {
    forgetDownload();
    fileInput.value = "";
    setState("idle");
    window.scrollTo({ top: 0, behavior: reduceMotion ? "auto" : "smooth" });
  });
  $("download").addEventListener("click", () => {
    if (!workbookBlob) return;
    if (!downloadUrl) downloadUrl = URL.createObjectURL(workbookBlob);
    const a = document.createElement("a");
    a.href = downloadUrl;
    a.download = outName;
    a.style.display = "none";
    document.body.append(a);
    a.click();
    a.remove();
    const label = $("downloadLabel");
    label.textContent = "Downloaded ✓";
    clearTimeout(flashTimer);
    flashTimer = setTimeout(() => { label.textContent = "Download .xlsx"; }, 2500);
  });

  // ---------- progress ----------
  const steps = Array.from(document.querySelectorAll(".step"));
  const labels = ["Read file", "Find questions", "Check answers", "Build sheet"];
  steps.forEach((el, i) => { el.lastChild.textContent = labels[i]; });
  function setStep(n) {
    steps.forEach((el, i) => {
      el.classList.toggle("done", i + 1 < n);
      el.classList.toggle("active", i + 1 === n);
    });
  }
  function setProgress(pct, text) {
    pct = Math.max(0, Math.min(100, Math.round(pct)));
    $("barFill").style.width = pct + "%";
    $("progressPct").textContent = pct + "%";
    if (text) $("progressText").textContent = text;
  }
  function onProgress(msg) {
    if (/^Finding/.test(msg)) { setStep(2); setProgress(40, "Finding short questions…"); }
    else { setStep(1); setProgress(10, msg); }
    return tick();
  }
  const fmtSize = (b) => (b > 1048576 ? (b / 1048576).toFixed(1) + " MB" : Math.max(1, Math.round(b / 1024)) + " KB");

  // ---------- results ----------
  function countUp(el, to) {
    if (reduceMotion || to === 0) { el.textContent = to; return; }
    const t0 = performance.now(), dur = 900;
    (function frame(now) {
      const p = Math.min(1, (now - t0) / dur), e = 1 - Math.pow(1 - p, 3);
      el.textContent = Math.round(to * e);
      if (p < 1) requestAnimationFrame(frame);
    })(t0);
  }
  function showStats(items) {
    const box = $("stats");
    box.textContent = "";
    for (const [label, value, bad] of items) {
      const d = document.createElement("div");
      d.className = "stat" + (bad && value ? " bad" : "");
      const b = document.createElement("b"), s = document.createElement("span");
      b.textContent = "0"; s.textContent = label;
      d.append(b, s); box.append(d);
      countUp(b, value);
    }
  }
  function showPreview(rows) {
    const box = $("preview");
    box.textContent = "";
    const h = document.createElement("h3");
    const shown = Math.min(rows.length, 50);
    h.innerHTML = "<span>Preview</span><small></small>";
    h.querySelector("small").textContent = rows.length > shown ? `first ${shown} of ${rows.length} questions — the .xlsx has all of them` : `${rows.length} questions`;
    const wrap = document.createElement("div"); wrap.className = "scroll";
    const table = document.createElement("table");
    const head = ["Status", "SI", "Type", "Chapter", "Board", "Title", "Answer"];
    const thead = table.createTHead().insertRow();
    head.forEach((t) => { const th = document.createElement("th"); th.textContent = t; thead.append(th); });
    const tbody = table.createTBody();
    rows.slice(0, shown).forEach((r, i) => {
      const tr = tbody.insertRow();
      tr.style.animationDelay = Math.min(i, 24) * 28 + "ms";
      const cell = (text, cls) => { const td = tr.insertCell(); if (cls) td.className = cls; td.textContent = text; return td; };
      const st = cell("");
      if (r[0] === "failed") { const s = document.createElement("span"); s.className = "pill failed"; s.textContent = "failed"; st.append(s); }
      cell(r[1]);
      const ty = cell(""); const tp = document.createElement("span"); tp.className = "pill m1"; tp.textContent = r[2]; ty.append(tp);
      cell(r[3], "chapter");
      cell(r[4], "board");
      const t = cell("", "title"); const d1 = document.createElement("div"); d1.textContent = r[5]; t.append(d1);
      const a = cell("", "title"); const d2 = document.createElement("div"); d2.textContent = r[6]; a.append(d2);
    });
    wrap.append(table);
    box.append(h, wrap);
  }

  // ---------- main ----------
  async function handle(f) {
    if (!f || busy) return;
    if (!/\.(docx|docm)$/i.test(f.name)) {
      errorBox.textContent = `"${f.name}" is not a Word file. Please choose a .docx or .docm file.`;
      setState("idle"); void app.offsetWidth; setState("error");
      return;
    }
    busy = true;
    forgetDownload();
    errorBox.textContent = "";
    $("fileName").textContent = f.name;
    $("fileSize").textContent = fmtSize(f.size);
    $("docExt").textContent = /\.docm$/i.test(f.name) ? "DOCM" : "DOCX";
    setStep(1); setProgress(2, "Opening the file…");
    setState("working");
    try {
      await tick();
      const data = await f.arrayBuffer();
      const { rows } = await MCQCore.extract(data, { kind: "sq", onProgress });
      if (!rows.length) throw new Error('No short questions found. They need a numbered title followed by an "উত্তর:" line, under a "… প্রশ্নোত্তর" heading.');
      setStep(4); setProgress(90, "Building the spreadsheet…");
      await tick();
      const wb = await MCQCore.buildWorkbook(ExcelJS, rows, {
        sheet: "SQ", headers: SQParser.HEADERS, widths: [8, 6, 16, 30, 24, 50, 90], imageCols: [],
      });
      const buf = await wb.xlsx.writeBuffer();
      workbookBlob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      outName = f.name.replace(/\.(docx|docm)$/i, "") + "_SQ.xlsx";
      setProgress(100, "Done");
      steps.forEach((el) => { el.classList.remove("active"); el.classList.add("done"); });
      await new Promise((r) => setTimeout(r, reduceMotion ? 0 : 450));

      const failed = rows.filter((r) => r[0] === "failed").length;
      const types = {};
      rows.forEach((r) => { types[r[2]] = (types[r[2]] || 0) + 1; });
      const chapters = new Set(rows.map((r) => r[3]).filter(Boolean)).size;
      $("status").textContent = `Done: ${rows.length} short questions extracted` +
        (failed ? ` ${failed} are marked "failed" in the status column (no answer found).` : ".");
      const typeStats = Object.entries(types).slice(0, 3).map(([k, v]) => [k || "(no type)", v]);
      showStats([["Short questions", rows.length], ...typeStats, ["Chapters", chapters], ["Marked failed", failed, true]]);
      showPreview(rows);
      setState("done");
    } catch (e) {
      console.error(e);
      errorBox.textContent = e && e.message && /^No short/.test(e.message) ? e.message : "Could not read the file: " + (e && e.message ? e.message : e);
      setState("idle"); void app.offsetWidth; setState("error");
    } finally {
      busy = false;
      fileInput.value = "";
    }
  }
})();
