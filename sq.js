/* SQ parser: short questions ("১। নৌ বিমা কী? [বোর্ড]" followed by "উত্তর: …") under banners such as
 * "জ্ঞানমূলক প্রশ্নোত্তর" / "অনুধাবনমূলক প্রশ্নোত্তর". Uses the building blocks exported by parser.js. */
(function (root) {
  "use strict";
  const P = root.MCQParser.internals;
  const { iterBlocks, paragraphInfo, splitBoard, kid, QUESTION_START, KA, IMG_RE_G, W } = P;

  const HAS_OPTION = /\(\s*[কখগঘঙ]\s*[)\s]/; // "(ক)" or a mistyped "(ক "
  const BANNER = /প্রশ্নোত্তর/;
  const ANSWER_START = /^\s*উত্তর\s*[:：ঃ]?\s*/;
  const BOARD_ONLY = /^\s*\[[^\[\]]*\]\s*$/;
  const MCQ_BANNER = /বহুনির্বাচনি|বহুনির্বাচনী/;

  // creative questions (সৃজনশীল): "প্রশ্ন-০১: <উদ্দীপক> (ক) … (খ) … (গ) … (ঘ) …" and later "১নং প্রশ্নের উত্তর (ক) … (খ) …"
  const CQ_STEM = /^\s*প্রশ্ন\s*[-–]\s*([০-৯0-9]+)\s*[:ঃ]?\s*/;
  const CQ_SUB = /^\s*\(\s*([কখগঘ])\s*\)\s*/;
  const CQ_ANSWER_HEAD = /^\s*([০-৯0-9]+)\s*নং\s*প্রশ্নের\s*উত্তর\s*[:ঃ]?\s*$/;
  // only these parts are short questions; (গ) and (ঘ) are the long ones
  const CQ_TYPES = { "ক": "জ্ঞানমূলক", "খ": "অনুধাবনমূলক" };

  // "(ক) A (খ) B (গ) C" -> [["ক","A"],["খ","B"],["গ","C"]]
  const splitSubs = (text) => {
    const marks = [...text.matchAll(/\(\s*([কখগঘ])\s*\)/g)];
    if (!marks.length || marks[0].index !== 0) return [];
    return marks.map((m, i) => [m[1], text.slice(m.index + m[0].length, i + 1 < marks.length ? marks[i + 1].index : text.length).trim()]);
  };
  // a creative question's part reads like a question or an instruction; an MCQ option does not
  const READS_AS_QUESTION = /[?？]|কী|কেন|কত|কোন|কীভাবে|কর[োো]?\b|করুন|লেখো|দাও|বলো|ব্যাখ্যা|বিশ্লেষণ|বর্ণনা|নির্ণয়|মতামত|যাচাই|মূল্যায়ন|আলোচনা/;

  const HEADERS = ["status", "SI", "type", "chapter", "Board", "title", "answer"];

  function parse(bodyEl, rels, styleNames, chapterOf) {
    const pars = [];
    let section = 0;
    const sectPrs = [];
    for (const item of iterBlocks(bodyEl)) {
      if (item.key) continue; // answer tables belong to MCQs
      const info = paragraphInfo(item.p, rels, null, styleNames || {});
      info.section = section;
      // banners are often drawn as text boxes: the box text arrives as `box`, on a paragraph that may have no text
      if (info.box && BANNER.test(info.box.normalize("NFC")) && info.box.length <= 80) pars.push({ raw: info.box, italic: "", style: "", section });
      if (info.raw) pars.push(info);
      const sp = kid(kid(item.p, "pPr", W), "sectPr", W);
      if (sp) sectPrs[section++] = sp;
    }
    const lastSp = kid(bodyEl, "sectPr", W);
    if (lastSp) sectPrs[section] = lastSp;
    const chapters = [];
    let prev = "";
    for (let n = 0; n <= section; n++) {
      chapters[n] = (chapterOf && sectPrs[n] ? chapterOf(sectPrs[n]) : "") || prev;
      prev = chapters[n];
    }

    const items = [];
    let type = "", inSQ = false, cur = null, mode = null; // mode: title | answer
    let cq = null, cqMode = null, cqLetter = ""; // creative question being read: stem | sub | ans
    const cqs = [];
    const textOf = (p) => p.raw.replace(IMG_RE_G, "").normalize("NFC").trim(); // composed vowel signs, so "ো" always matches

    // a numbered line starts a short question when "উত্তর:" comes with it or right after it, with no (ক) options
    const startsSQ = (i) => {
      const t = textOf(pars[i]);
      if (!QUESTION_START.test(t)) return false;
      if (HAS_OPTION.test(t)) return false; // options on the line itself: an MCQ, whatever follows
      if (ANSWER_START.test(t.replace(QUESTION_START, "")) === false && /উত্তর\s*[:：ঃ]/.test(t)) return true; // answer on the same line
      for (let j = i + 1; j < Math.min(i + 4, pars.length); j++) {
        const u = textOf(pars[j]);
        if (QUESTION_START.test(u) || HAS_OPTION.test(u)) return false;
        if (ANSWER_START.test(u)) return true;
      }
      return false;
    };

    const cqItems = [];
    const finishCQ = () => { cq = null; cqMode = null; cqLetter = ""; };
    const finish = () => { if (cur) { cur.answer = cur.answer.trim(); items.push(cur); } cur = null; mode = null; };

    pars.forEach((par, i) => {
      const text = textOf(par);
      if (!text) return;

      // ---- creative questions ----
      const isBanner = text.length <= 60 && (BANNER.test(text) || MCQ_BANNER.test(text)) && !QUESTION_START.test(text);
      if (isBanner || /^(heading|title)/i.test(par.style)) finishCQ();
      const stemM = CQ_STEM.exec(text);
      if (stemM) {
        finish(); finishCQ();
        cq = { num: parseInt([...stemM[1]].map((c) => ("০১২৩৪৫৬৭৮৯".includes(c) ? "০১২৩৪৫৬৭৮৯".indexOf(c) : c)).join(""), 10),
               chapter: chapters[par.section] || "", board: "", stem: "", subs: {}, answers: {} };
        cqs.push(cq);
        cqMode = "stem";
        const [rest, board] = splitBoard(par, text.replace(CQ_STEM, ""));
        cq.stem = rest.trim(); cq.board = board;
        return;
      }
      const headM = CQ_ANSWER_HEAD.exec(text);
      if (headM) { // "১নং প্রশ্নের উত্তর": answers of the latest creative question with that number
        const n = parseInt([...headM[1]].map((c) => ("০১২৩৪৫৬৭৮৯".includes(c) ? "০১২৩৪৫৬৭৮৯".indexOf(c) : c)).join(""), 10);
        const target = [...cqs].reverse().find((c) => c.num === n && Object.keys(c.answers).length === 0);
        if (target) { finishCQ(); cq = target; cq.done = true; cqMode = "ans"; cqLetter = ""; return; }
      }
      if (cq && cqMode === "ans" && QUESTION_START.test(text) && pars.slice(i, i + 4).some((p) => KA.test(textOf(p)))) finishCQ(); // an MCQ starts
      if (cq) {
        const sub = CQ_SUB.exec(text);
        if (cqMode === "stem" || cqMode === "sub") {
          const pieces = sub ? splitSubs(text) : [];
          // several labels on one line with nothing that reads like a question: an MCQ's options, so this was no creative question
          if (pieces.length > 1 && !pieces.some((p) => READS_AS_QUESTION.test(p[1]))) {
            if (!Object.keys(cq.subs).length) cqs.splice(cqs.indexOf(cq), 1);
            finishCQ();
            return;
          }
          if (pieces.length) {
            cqMode = "sub";
            for (const [l, t] of pieces) { cqLetter = l; cq.subs[l] = t; }
            return;
          }
          if (cqMode === "sub") { cq.subs[cqLetter] = (cq.subs[cqLetter] + " " + text).trim(); return; }
          const [t, b] = splitBoard(par, text);
          if (t) cq.stem = (cq.stem ? cq.stem + "\n" : "") + t;
          cq.board = cq.board || b;
          return;
        }
        if (cqMode === "ans") {
          if (sub) { for (const [l, t] of splitSubs(text)) { cqLetter = l; cq.answers[l] = t; } return; }
          if (cqLetter) { cq.answers[cqLetter] = (cq.answers[cqLetter] + " " + text).trim(); return; }
          return;
        }
      }

      // section banners decide the type ("জ্ঞানমূলক", "অনুধাবনমূলক" …); MCQ sections switch collecting off
      if (isBanner) {
        finish();
        if (MCQ_BANNER.test(text)) { inSQ = false; return; }
        type = text.replace(BANNER, "").replace(/[\s:–-]+$/, "").trim();
        inSQ = true;
        return;
      }
      if (/^(heading|title)/i.test(par.style)) { finish(); return; }
      if (!inSQ) return;

      if (startsSQ(i)) {
        finish();
        let line = text.replace(QUESTION_START, "");
        let answerPart = "";
        const am = /(.*?)\s*উত্তর\s*[:：ঃ]\s*(.*)$/s.exec(line);
        if (am) { line = am[1]; answerPart = am[2]; }
        const [title, board] = splitBoard(par, line);
        cur = { type, chapter: chapters[par.section] || "", board, title: title.trim(), answer: answerPart };
        mode = am ? "answer" : "title";
        return;
      }
      if (!cur) return;

      if (BOARD_ONLY.test(text)) { // the board often sits on a line of its own, after the question
        if (!cur.board) cur.board = text.trim();
        return;
      }
      if (mode === "title" && ANSWER_START.test(text)) {
        cur.answer = text.replace(ANSWER_START, "");
        mode = "answer";
        return;
      }
      if (mode === "answer") cur.answer = (cur.answer + " " + text).trim();
      else if (mode === "title") {
        const [t, b] = splitBoard(par, text); // title wrapped over several lines
        if (t) cur.title = cur.title + "\n" + t;
        cur.board = cur.board || b;
      }
    });
    finish(); finishCQ();
    // creative questions are added once their answers have been read: in the order the stems appear
    for (const c of cqs) {
      for (const letter of Object.keys(CQ_TYPES)) {
        if (c.subs[letter] === undefined) continue;
        cqItems.push({ type: CQ_TYPES[letter], chapter: c.chapter, board: c.board,
          title: c.subs[letter].trim(), answer: (c.answers[letter] || "").trim() });
      }
    }
    return { items: items.concat(cqItems), cqCount: cqs.length, cqNoAnswer: cqs.filter((c) => !Object.keys(c.answers).length).length };
  }

  // Board is shown without its brackets: "[য. বো. ১৭]" -> "য. বো. ১৭"
  const unbracket = (b) => (b || "").replace(/^\s*\[\s*/, "").replace(/\s*\]\s*$/, "");

  function toRows(items) {
    return items.map((q, i) => [q.answer ? "" : "failed", i + 1, q.type, q.chapter, unbracket(q.board), q.title, q.answer]);
  }

  const api = { parse, toRows, HEADERS };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.SQParser = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
