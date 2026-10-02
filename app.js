(function () {
  "use strict";
  const $ = (id) => document.getElementById(id);
  const app = $("app"), drop = $("drop"), fileInput = $("file"), errorBox = $("error");
  const reduceMotion = window.matchMedia && matchMedia("(prefers-reduced-motion: reduce)").matches;
  let workbookBlob = null, outName = "mcqs.xlsx", busy = false;

  const tick = () => new Promise((r) => setTimeout(r, 0)); // let the page repaint between steps
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

  // ---------- Google Drive upload (Apps Script web app set in config.js) ----------
  const driveUrl = ((window.MCQ_CONFIG || {}).driveUrl || "").trim();

  // columns of the sheet that hold a picture: question image, then one per option
  const IMAGE_COLS = [6, 10, 13, 16, 19];
  // Upload every picture and put its Drive link in the cell. Anything that cannot be uploaded
  // (script deleted, folder gone, no network) simply keeps its base64 text.
  async function uploadPictures(rows, baseName, onStep) {
    const url = driveUrl;
    const jobs = [], byData = new Map(); // identical pictures are uploaded once
    rows.forEach((r, ri) => IMAGE_COLS.forEach((c, k) => {
      if (typeof r[c] === "string" && r[c].startsWith("data:image")) {
        jobs.push({ ri, c, name: `${baseName}_Q${r[1]}_${k === 0 ? "question" : "option" + k}.png`, data: r[c] });
      }
    }));
    let done = 0, failed = 0, strikes = 0;
    const send = async (job) => {
      if (strikes >= 3) return null; // the endpoint is not answering: keep base64 for the rest, don't wait on each one
      for (let attempt = 0; attempt < 2; attempt++) {
        try {
          const res = await fetch(url, {
            method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" },
            body: JSON.stringify({ name: job.name, mime: "image/png", data: job.data.split(",")[1] }),
          });
          const j = await res.json();
          // Drive's own viewer page: picture centred, with the download button
          if (j && j.ok && (j.id || j.url)) { strikes = 0; return j.id ? `https://drive.google.com/file/d/${j.id}/view` : j.url; }
        } catch (e) { /* retry */ }
      }
      strikes++;
      return null;
    };
    let next = 0;
    const worker = async () => {
      while (next < jobs.length) {
        const job = jobs[next++];
        if (!byData.has(job.data)) byData.set(job.data, send(job));
        const link = await byData.get(job.data);
        if (link) rows[job.ri][job.c] = link; else failed++;
        onStep(++done, jobs.length);
      }
    };
    await Promise.all(Array.from({ length: Math.min(4, jobs.length) }, worker));
    return { total: jobs.length, failed };
  }

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

  // the download link is made once, when the sheet is ready, and kept until a new file is chosen
  let downloadUrl = null, flashTimer = null;
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
    document.body.append(a); // attached to the page, which some browsers need for the download to start
    a.click();
    a.remove();
    // visible confirmation, so a download that starts quietly is not mistaken for a click that did nothing
    const label = $("downloadLabel");
    label.textContent = "Downloaded ✓";
    clearTimeout(flashTimer);
    flashTimer = setTimeout(() => { label.textContent = "Download .xlsx"; }, 2500);
  });

  // ---------- progress ----------
  const steps = Array.from(document.querySelectorAll(".step"));
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
    const m = /(\d+) of (\d+)/.exec(msg);
    if (m) { setStep(3); setProgress(25 + (parseInt(m[1], 10) / parseInt(m[2], 10)) * 65, msg); }
    else if (/^Finding/.test(msg)) { setStep(2); setProgress(18, "Finding questions…"); }
    else { setStep(1); setProgress(6, msg); }
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
    const head = ["Status", "SI", "Type", "Chapter", "Board", "Title", "Image", "Answer"];
    const thead = table.createTHead().insertRow();
    head.forEach((t) => { const th = document.createElement("th"); th.textContent = t; thead.append(th); });
    const tbody = table.createTBody();
    rows.slice(0, shown).forEach((r, i) => {
      const tr = tbody.insertRow();
      tr.style.animationDelay = Math.min(i, 24) * 28 + "ms";
      const cell = (text, cls) => { const td = tr.insertCell(); if (cls) td.className = cls; td.textContent = text; return td; };
      const pill = (td, text, kind) => { td.textContent = ""; const s = document.createElement("span"); s.className = "pill " + kind; s.textContent = text; td.append(s); };
      pill(cell(""), r[0] === "failed" ? "failed" : "ok", r[0] === "failed" ? "failed" : "m1");
      if (r[0] !== "failed") tr.cells[0].firstChild.style.visibility = "hidden";
      cell(r[1]);
      pill(cell(""), r[2], r[2]);
      cell(r[3], "chapter");
      cell(r[4], "board");
      const t = cell("", "title"); const div = document.createElement("div"); div.textContent = r[5]; t.append(div);
      cell(r[6] ? "yes" : "");
      const correct = [0, 1, 2, 3].filter((k) => r[9 + 3 * k] === 1).map((k) => "কখগঘ"[k]).join("");
      const a = cell(""); if (correct) pill(a, correct, "ans");
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
      const { rows, keys } = await MCQCore.extract(data, { encode: MCQImages.encode, onProgress });
      let upload = null;
      if (rows.length && driveUrl) {
        setStep(3);
        upload = await uploadPictures(rows, f.name.replace(/\.(docx|docm)$/i, ""), (d, t) => {
          setProgress(90 + (d / t) * 4, `Uploading pictures to Drive: ${d} of ${t}…`);
          return tick();
        });
      }
      if (!rows.length) throw new Error('No MCQs found. Questions need a numbered title, (ক)–(ঘ) options and an "উত্তর:" line or answer table.');
      setStep(4); setProgress(94, "Building the spreadsheet…");
      await tick();
      const wb = await MCQCore.buildWorkbook(ExcelJS, rows);
      const buf = await wb.xlsx.writeBuffer();
      workbookBlob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      outName = f.name.replace(/\.(docx|docm)$/i, "") + ".xlsx";
      setProgress(100, "Done");
      steps.forEach((el) => { el.classList.remove("active"); el.classList.add("done"); });
      await new Promise((r) => setTimeout(r, reduceMotion ? 0 : 450));

      const failed = rows.filter((r) => r[0] === "failed").length;
      const withImages = rows.filter((r) => r[2] === "m4").length;
      const chapters = new Set(rows.map((r) => r[3]).filter(Boolean)).size;
      $("status").textContent =
        `Done: ${rows.length} questions extracted` + (withImages ? ` (${withImages} with pictures).` : ".") +
        (failed ? ` ${failed} question(s) are marked "failed" in the status column (no answer found, or fewer than four options).` : "") +
        ` Answer tables found: ${keys}.` +
        (upload && upload.total ? ` Pictures uploaded to Drive: ${upload.total - upload.failed} of ${upload.total}` +
          (upload.failed ? " (the rest stay as base64)." : ".") : "");
      showStats([["Questions", rows.length], ["With pictures", withImages], ["Chapters", chapters], ["Answer tables", keys], ["Marked failed", failed, true]]);
      showPreview(rows);
      setState("done");
    } catch (e) {
      console.error(e);
      errorBox.textContent = e && e.message && /^No MCQs/.test(e.message) ? e.message : "Could not read the file: " + (e && e.message ? e.message : e);
      setState("idle"); void app.offsetWidth; setState("error");
    } finally {
      busy = false;
      fileInput.value = "";
    }
  }
})();
