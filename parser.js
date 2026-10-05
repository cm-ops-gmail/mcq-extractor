/* MCQ parser: turns the XML of a .docx into MCQ records. Pure logic, no browser-only APIs,
 * so it runs in the page and in Node tests. Equations are converted to KaTeX (LaTeX). */
(function (root) {
  "use strict";

  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
  const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const A = "http://schemas.openxmlformats.org/drawingml/2006/main";

  const BN = "০১২৩৪৫৬৭৮৯";
  const LETTERS = ["ক", "খ", "গ", "ঘ"];
  const ALL_LETTERS = ["ক", "খ", "গ", "ঘ", "ঙ"]; // some books print a fifth option; the sheet only has four

  const QUESTION_START = new RegExp(`^\\s*([${BN}0-9]+)\\s*[।.)]\\s*`);
  const SECTION_HEADING = /প্রশ্নোত্তর/;
  const SECTION_MARKER = /^\s*(mcq|cq)\s*\(\s*(start|end)\s*\)\s*$/i;
  // "[নিচের] উদ্দীপক/অনুচ্ছেদ/তথ্য/চিত্র/ছক … পড়ো/পড়ে/লক্ষ্য কর … ৩৯ ও ৪০ নং প্রশ্নের উত্তর দাও:"
  const STIMULUS_WORD = /^\s*(নিচের\s+)?(উদ্দীপক|অনুচ্ছেদ|তথ্য|চিত্র|ছক|কবিতা|গদ্যাংশ)\S*\s/;
  const STIMULUS_START = { test: (t) => STIMULUS_WORD.test(t) && /উত্তর\s*দাও/.test(t.slice(0, 160)) && /[০-৯0-9]/.test(t.slice(0, 120)) && !QUESTION_START.test(t) };
  const OPTION_MARK = /\(\s*([কখগঘঙ])\s*\)+/;
  const OPTION_MARK_G = /\(\s*([কখগঘঙ])\s*\)+/g; // "(ঘ))" typos are accepted
  const OPTION_OPEN_G = /\(\s*([কখগঘঙ])(?=\s+\S)/g; // "(গ অর্থনৈতিক" typo: closing bracket missing
  const ANSWER = /উত্তর\s*[:：]?\s*\(?\s*([কখগঘঙ])\s*\)?/;
  const EXPLANATION_ANY = /ব্যাখ্যা\s*[:：]?\s*/;
  const EXPLANATION_START = /^ব্যাখ্যা\s*[:：]?\s*/;
  const BOARD_TAIL = /\s*(\[[^\[\]]*\])\s*$/;
  const KA = /\(\s*ক\s*\)/;

  const IMG_L = "", IMG_R = "";
  const IMG_RE_G = new RegExp(`${IMG_L}(\\d+)${IMG_R}`, "g");

  // ---------- small DOM helpers ----------
  const kids = (el) => Array.from(el.children || []);
  const local = (el) => el.localName || "";
  function kid(el, name, ns) {
    if (!el) return null;
    for (const c of kids(el)) if (c.localName === name && c.namespaceURI === (ns || M)) return c;
    return null;
  }
  function mval(el, dflt) {
    if (!el) return dflt;
    const v = el.getAttributeNS(M, "val");
    return v === null || v === undefined ? dflt : v;
  }
  function descendants(el) { return Array.from(el.getElementsByTagName("*")); }

  // ---------- Word equations (OMML) -> KaTeX ----------
  const GREEK = {
    "α": "alpha", "β": "beta", "γ": "gamma", "δ": "delta", "ε": "epsilon", "ϵ": "epsilon",
    "ζ": "zeta", "η": "eta", "θ": "theta", "ϑ": "vartheta", "ι": "iota", "κ": "kappa",
    "λ": "lambda", "μ": "mu", "ν": "nu", "ξ": "xi", "π": "pi", "ρ": "rho", "σ": "sigma",
    "ς": "varsigma", "τ": "tau", "υ": "upsilon", "φ": "phi", "ϕ": "varphi", "χ": "chi",
    "ψ": "psi", "ω": "omega", "Γ": "Gamma", "Δ": "Delta", "∆": "Delta", "Θ": "Theta",
    "Λ": "Lambda", "Ξ": "Xi", "Π": "Pi", "Σ": "Sigma", "Φ": "Phi", "Ψ": "Psi", "Ω": "Omega",
  };
  const SYMBOLS = {
    "×": "\\times", "÷": "\\div", "±": "\\pm", "∓": "\\mp", "·": "\\cdot", "⋅": "\\cdot",
    "∗": "\\ast", "∘": "\\circ", "∙": "\\bullet", "≤": "\\leq", "≥": "\\geq", "≠": "\\neq",
    "≈": "\\approx", "∼": "\\sim", "≅": "\\cong", "≡": "\\equiv", "∝": "\\propto",
    "≪": "\\ll", "≫": "\\gg", "→": "\\rightarrow", "←": "\\leftarrow", "↔": "\\leftrightarrow",
    "⇒": "\\Rightarrow", "⇐": "\\Leftarrow", "⇔": "\\Leftrightarrow", "⟶": "\\longrightarrow",
    "⟵": "\\longleftarrow", "⇌": "\\rightleftharpoons", "↑": "\\uparrow", "↓": "\\downarrow",
    "↦": "\\mapsto", "∞": "\\infty", "∂": "\\partial", "∇": "\\nabla", "∅": "\\emptyset",
    "∀": "\\forall", "∃": "\\exists", "∈": "\\in", "∉": "\\notin", "⊂": "\\subset",
    "⊃": "\\supset", "⊆": "\\subseteq", "⊇": "\\supseteq", "∩": "\\cap", "∪": "\\cup",
    "∧": "\\land", "∨": "\\lor", "¬": "\\neg", "∠": "\\angle", "△": "\\triangle",
    "°": "^{\\circ}", "′": "'", "″": "''", "−": "-", "–": "-", "—": "-", "∑": "\\sum",
    "∏": "\\prod", "∫": "\\int", "∬": "\\iint", "∭": "\\iiint", "∮": "\\oint",
    "⋃": "\\bigcup", "⋂": "\\bigcap", "…": "\\ldots", "⋯": "\\cdots", "⋮": "\\vdots",
    "⋱": "\\ddots", "∴": "\\therefore", "∵": "\\because", "⊥": "\\perp", "∥": "\\parallel",
    "⊕": "\\oplus", "⊗": "\\otimes", "ℏ": "\\hbar", "ℓ": "\\ell", "ℝ": "\\mathbb{R}",
    "ℕ": "\\mathbb{N}", "ℤ": "\\mathbb{Z}", "ℚ": "\\mathbb{Q}", "ℂ": "\\mathbb{C}",
    "⟨": "\\langle", "⟩": "\\rangle", "⌊": "\\lfloor", "⌋": "\\rfloor", "⌈": "\\lceil",
    "⌉": "\\rceil", "‖": "\\|", "∣": "|", " ": "\\ ", " ": "\\,", " ": "\\ ",
    "%": "\\%", "#": "\\#", "$": "\\$", "{": "\\{", "}": "\\}", "_": "\\_",
    "~": "\\sim", "^": "\\hat{}", "\\": "\\backslash ",
  };
  const SUB_CHARS = {};
  [..."₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎"].forEach((c, i) => (SUB_CHARS[c] = "0123456789+-=()"[i]));
  const SUP_CHARS = {};
  [..."⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱ"].forEach((c, i) => (SUP_CHARS[c] = "0123456789+-=()ni"[i]));
  const ACCENTS = {
    "̂": "hat", "^": "hat", "̃": "tilde", "~": "tilde", "̄": "bar",
    "¯": "bar", "̅": "bar", "⃗": "vec", "→": "vec", "̇": "dot",
    "˙": "dot", "̈": "ddot", "¨": "ddot", "̌": "check", "ˇ": "check",
    "́": "acute", "´": "acute", "̀": "grave", "`": "grave", "̆": "breve",
  };
  const NARY = {
    "∑": "\\sum", "∏": "\\prod", "∫": "\\int", "∬": "\\iint", "∭": "\\iiint",
    "∮": "\\oint", "⋃": "\\bigcup", "⋂": "\\bigcap", "⨁": "\\bigoplus",
  };
  const ARROWS = new Set(["→", "⟶", "⟵", "←", "↔", "⇌", "⇒", "⇐", "⇔"]);
  const XARROW = {
    "→": "xrightarrow", "⟶": "xrightarrow", "←": "xleftarrow", "⟵": "xleftarrow",
    "↔": "xleftrightarrow", "⇒": "xRightarrow", "⇐": "xLeftarrow", "⇔": "xLeftrightarrow",
    "⇌": "rightleftharpoons",
  };
  const FUNCS = new Set(["sin", "cos", "tan", "cot", "sec", "csc", "log", "ln", "exp", "lim",
    "max", "min", "sup", "inf", "det", "gcd", "arcsin", "arccos", "arctan", "sinh", "cosh", "tanh"]);
  const DELIMS = {
    "": ".", "(": "(", ")": ")", "[": "[", "]": "]", "{": "\\{", "}": "\\}", "|": "|",
    "‖": "\\|", "⟨": "\\langle", "⟩": "\\rangle", "⌊": "\\lfloor", "⌋": "\\rfloor",
    "⌈": "\\lceil", "⌉": "\\rceil", "〈": "\\langle", "〉": "\\rangle",
  };
  const has = (o, k) => Object.prototype.hasOwnProperty.call(o, k);

  // join LaTeX pieces, adding a space where a control word meets a letter
  function cat(...parts) {
    let out = "";
    for (const p of parts) {
      if (!p) continue;
      if (out && /\\[A-Za-z]+$/.test(out) && /^[A-Za-z]/.test(p)) out += " ";
      out += p;
    }
    return out;
  }
  function balanced(t) {
    let d = 0;
    for (const c of t) {
      d += (c === "{") - (c === "}");
      if (d < 0) return false;
    }
    return d === 0;
  }
  // wrap in braces unless already a single token
  function group(tex) {
    if (/^(\\[A-Za-z]+|[^\\{}])$/.test(tex) || (tex.startsWith("{") && tex.endsWith("}") && balanced(tex.slice(1, -1)))) {
      return tex ? tex : "{}";
    }
    return tex ? "{" + tex + "}" : "{}";
  }
  function charTex(ch) {
    if (has(GREEK, ch)) return "\\" + GREEK[ch];
    if (has(SUB_CHARS, ch)) return "_{" + SUB_CHARS[ch] + "}";
    if (has(SUP_CHARS, ch)) return "^{" + SUP_CHARS[ch] + "}";
    return has(SYMBOLS, ch) ? SYMBOLS[ch] : ch;
  }

  function runMath(r, inEqArr) {
    const pr = kid(r, "rPr");
    let sty = null, nor = false;
    if (pr) {
      const s = kid(pr, "sty");
      sty = s ? mval(s, null) : null;
      const n = kid(pr, "nor");
      nor = !!n && !["0", "false"].includes(mval(n, "1"));
    }
    const text = kids(r).filter((c) => c.localName === "t" && c.namespaceURI === M).map((c) => c.textContent).join("");
    if (!text.trim()) return text && !nor ? "\\ " : "";
    if (nor) return "\\text{" + text.replace(/[\\{}$]/g, "") + "}";
    let body = "";
    for (const ch of text) {
      if (ch === " " || ch === "\t") continue;
      if (ch === "&" && inEqArr) body += "&";
      else if (ch === "&") body += "\\&";
      else body = cat(body, charTex(ch));
    }
    if (FUNCS.has(text) && sty !== "i") return "\\" + text;
    const simple = /^[A-Za-z0-9.]+$/.test(body);
    if (sty === "p") return simple ? "\\mathrm{" + body + "}" : body;
    return body; // bold / bold-italic are not marked up: plain letters are enough
  }

  function kidTex(el, name, eq) { const k = kid(el, name); return k ? tex(k, eq) : ""; }
  function chrOf(el, prop, dflt) {
    const pr = kid(el, prop);
    if (!pr) return dflt;
    const c = kid(pr, "chr");
    return c ? mval(c, null) : dflt;
  }
  function flag(pr, name) {
    if (!pr) return false;
    const f = kid(pr, name);
    return !!f && !["0", "false"].includes(mval(f, "1"));
  }
  const allText = (el) => (el ? descendants(el).filter((t) => t.localName === "t" && t.namespaceURI === M).map((t) => t.textContent).join("").trim() : "");

  function tex(el, eq) {
    const tag = local(el);
    if (tag === "r" && el.namespaceURI === M) return runMath(el, eq);
    if (tag === "sSub") return cat(group(kidTex(el, "e", eq)), "_{" + kidTex(el, "sub", eq) + "}");
    if (tag === "sSup") return cat(group(kidTex(el, "e", eq)), "^{" + kidTex(el, "sup", eq) + "}");
    if (tag === "sSubSup") {
      return cat(group(kidTex(el, "e", eq)), "_{" + kidTex(el, "sub", eq) + "}^{" + kidTex(el, "sup", eq) + "}");
    }
    if (tag === "sPre") {
      return cat("{}_{" + kidTex(el, "sub", eq) + "}^{" + kidTex(el, "sup", eq) + "}", group(kidTex(el, "e", eq)));
    }
    if (tag === "f") {
      const num = kidTex(el, "num", eq), den = kidTex(el, "den", eq);
      const t = kid(kid(el, "fPr"), "type");
      const ftype = t ? mval(t, "bar") : "bar";
      if (ftype === "noBar") return "\\binom{" + num + "}{" + den + "}";
      if (ftype === "lin") return cat(group(num), "/", group(den));
      return "\\frac{" + num + "}{" + den + "}";
    }
    if (tag === "rad") {
      const e = kidTex(el, "e", eq);
      const hide = flag(kid(el, "radPr"), "degHide");
      const deg = kidTex(el, "deg", eq);
      return hide || !deg ? "\\sqrt{" + e + "}" : "\\sqrt[" + deg + "]{" + e + "}";
    }
    if (tag === "d") {
      const pr = kid(el, "dPr");
      const pick = (name, dflt) => { const c = pr ? kid(pr, name) : null; return c ? mval(c, "") : dflt; };
      const beg = pick("begChr", "("), end = pick("endChr", ")"), sep = pick("sepChr", "|");
      const items = kids(el).filter((c) => c.localName === "e" && c.namespaceURI === M).map((e) => tex(e, eq));
      const dl = (c) => (has(DELIMS, c) ? DELIMS[c] : c);
      const inner = items.join(" " + dl(sep) + " ");
      return "\\left" + dl(beg) + " " + inner + " \\right" + dl(end);
    }
    if (tag === "nary") {
      const pr = kid(el, "naryPr");
      const c = chrOf(el, "naryPr", "∫");
      const op = has(NARY, c) ? NARY[c] : (c === "" || c === null ? "\\int" : charTex(c));
      const sub = kidTex(el, "sub", eq), sup = kidTex(el, "sup", eq);
      const ll = pr ? kid(pr, "limLoc") : null;
      const und = !!ll && mval(ll, null) === "undOvr";
      let out = op;
      if (und) out += "\\limits";
      if (sub && !flag(pr, "subHide")) out += "_{" + sub + "}";
      if (sup && !flag(pr, "supHide")) out += "^{" + sup + "}";
      return cat(out, kidTex(el, "e", eq));
    }
    if (tag === "func") {
      const fn = kid(el, "fName");
      const raw = allText(fn);
      const simple = !!fn && !kid(fn, "limLow") && !kid(fn, "limUpp");
      let name;
      if (simple && FUNCS.has(raw)) name = "\\" + raw;
      else if (simple && /^\p{L}+$/u.test(raw)) name = "\\operatorname{" + raw + "}";
      else name = kidTex(el, "fName", eq);
      return cat(name, kidTex(el, "e", eq));
    }
    if (tag === "acc") {
      const c = chrOf(el, "accPr", "̂");
      const cmd = has(ACCENTS, c) ? ACCENTS[c] : "hat";
      return "\\" + cmd + "{" + kidTex(el, "e", eq) + "}";
    }
    if (tag === "bar") {
      const p = kid(kid(el, "barPr"), "pos");
      const pos = p ? mval(p, "top") : "top";
      return (pos === "bot" ? "\\underline{" : "\\overline{") + kidTex(el, "e", eq) + "}";
    }
    if (tag === "borderBox") return "\\boxed{" + kidTex(el, "e", eq) + "}";
    if (tag === "groupChr") {
      const c = chrOf(el, "groupChrPr", "⏟");
      const p = kid(kid(el, "groupChrPr"), "pos");
      const pos = p ? mval(p, "bot") : "bot";
      const e = kidTex(el, "e", eq);
      if (ARROWS.has(c)) return allText(kid(el, "e")) ? "\\" + (pos === "top" ? "overset" : "underset") + "{" + charTex(c) + "}{" + e + "}" : charTex(c);
      return c === "⏞" || c === "︷" || pos === "top" ? "\\overbrace{" + e + "}" : "\\underbrace{" + e + "}";
    }
    if (tag === "limLow" || tag === "limUpp") {
      const eEl = kid(el, "e");
      let base = allText(eEl);
      const gc = kid(eEl, "groupChr"); // Word draws a reaction arrow as an arrow-character group over blanks
      if (gc && !base && ARROWS.has(chrOf(gc, "groupChrPr", ""))) base = chrOf(gc, "groupChrPr", "");
      const lim = kidTex(el, "lim", eq);
      const e = eEl ? tex(eEl, eq) : "";
      if (ARROWS.has(base)) {
        const cmd = XARROW[base];
        if (cmd === "rightleftharpoons") {
          return tag === "limUpp" ? "\\overset{" + lim + "}{\\rightleftharpoons}" : "\\underset{" + lim + "}{\\rightleftharpoons}";
        }
        return tag === "limUpp" ? "\\" + cmd + "{" + lim + "}" : "\\" + cmd + "[" + lim + "]{}";
      }
      if (FUNCS.has(base)) return "\\" + base + (tag === "limLow" ? "_{" + lim + "}" : "^{" + lim + "}");
      return tag === "limLow" ? "\\underset{" + lim + "}{" + e + "}" : "\\overset{" + lim + "}{" + e + "}";
    }
    if (tag === "m") {
      const rows = kids(el).filter((c) => c.localName === "mr").map((mr) =>
        kids(mr).filter((c) => c.localName === "e").map((e) => tex(e, eq)).join(" & "));
      return "\\begin{matrix} " + rows.join(" \\\\ ") + " \\end{matrix}";
    }
    if (tag === "eqArr") {
      const rows = kids(el).filter((c) => c.localName === "e").map((e) => tex(e, true));
      return "\\begin{aligned} " + rows.join(" \\\\ ") + " \\end{aligned}";
    }
    if (tag === "phant") {
      const show = kid(kid(el, "phantPr"), "show");
      const e = kidTex(el, "e", eq);
      return show && ["0", "false"].includes(mval(show, null)) ? "\\phantom{" + e + "}" : e;
    }
    // containers (oMath, e, num, den, sub, sup, lim, deg ...) and unknown elements
    let out = "";
    for (const k of kids(el)) {
      const n = local(k);
      if (n.endsWith("Pr") || n === "ctrlPr") continue;
      out = cat(out, tex(k, eq));
    }
    return out;
  }

  // ---------- pictures ----------
  function elementImages(el, rels) {
    const out = [], seen = new Set();
    for (const x of descendants(el)) {
      const t = local(x);
      let rid = null, crop = null;
      if (t === "blip") {
        rid = x.getAttributeNS(R, "embed");
        const rect = kid(x.parentNode, "srcRect", A);
        if (rect) crop = ["l", "t", "r", "b"].map((k) => parseInt(rect.getAttribute(k) || "0", 10) / 100000);
      } else if (t === "imagedata") {
        rid = x.getAttributeNS(R, "id");
      } else continue;
      const target = rels[rid];
      if (!target || seen.has(target)) continue; // Word stores the same picture twice (drawing + fallback)
      seen.add(target);
      out.push({ target, crop });
    }
    return out;
  }

  // ---------- paragraph text: equations -> KaTeX, sub/superscript runs -> KaTeX, pictures -> marks ----------
  const LATIN_TAIL = /[A-Za-z()\[\]]+$/;
  const LATIN_HEAD = /^[A-Za-z()\[\]]+/;

  function paraText(p, rels, registry) {
    const segs = []; // [kind, text, chem]  kind: "t" text | "m" math

    function addText(t) {
      if (!t) return;
      const last = segs[segs.length - 1];
      if (last && last[0] === "m" && last[2]) {
        const h = LATIN_HEAD.exec(t);
        if (h) {
          last[1] += h[0];
          t = t.slice(h[0].length);
          if (!t) return;
        }
      }
      const l2 = segs[segs.length - 1];
      if (l2 && l2[0] === "t") l2[1] += t;
      else segs.push(["t", t, false]);
    }
    function addScript(t, marker) {
      const script = `${marker}{${t}}`;
      const last = segs[segs.length - 1];
      if (last && last[0] === "m") {
        last[1] += script;
        last[2] = true;
        return;
      }
      if (last && last[0] === "t") {
        const m = LATIN_TAIL.exec(last[1]);
        if (m) {
          last[1] = last[1].slice(0, m.index);
          if (!last[1]) segs.pop();
          segs.push(["m", m[0] + script, true]);
          return;
        }
      }
      segs.push(["m", "{}" + script, true]);
    }
    function addMath(t) {
      if (!t.trim()) return;
      const last = segs[segs.length - 1];
      if (last && last[0] === "m") {
        last[1] = cat(last[1], t);
        last[2] = false;
      } else segs.push(["m", t, false]);
    }
    function walk(node) {
      for (const c of kids(node)) {
        const t = local(c);
        if (t === "r" && c.namespaceURI === W) {
          const va = kid(kid(c, "rPr", W), "vertAlign", W);
          const mode = va ? va.getAttributeNS(W, "val") : null;
          let out = "";
          const flush = () => {
            if ((mode === "subscript" || mode === "superscript") && out.trim()) addScript(out.trim(), mode === "subscript" ? "_" : "^");
            else addText(out);
            out = "";
          };
          for (const x of kids(c)) {
            const xt = local(x);
            if (xt === "t") out += x.textContent;
            else if (xt === "tab") out += "\t";
            else if (xt === "br" || xt === "cr") out += " ";
            else if (["drawing", "pict", "object", "AlternateContent"].includes(xt) && registry) {
              flush();
              for (const im of elementImages(x, rels)) {
                addText(`${IMG_L}${registry.length}${IMG_R}`);
                registry.push(im);
              }
            }
          }
          flush();
        } else if (t === "oMath") addMath(tex(c, false));
        else if (t === "oMathPara") {
          for (const om of kids(c)) if (local(om) === "oMath") addMath(tex(om, false));
        } else if (["hyperlink", "smartTag", "sdt", "sdtContent", "ins", "fldSimple"].includes(t)) walk(c);
      }
    }
    walk(p);
    return segs.map((s) => s[1]).join("");
  }

  // ---------- picture / option-label tokens ----------
  function tokens(raw) {
    const toks = [];
    for (const piece of raw.split(new RegExp(`(${IMG_L}\\d+${IMG_R})`))) {
      const m = new RegExp(`^${IMG_L}(\\d+)${IMG_R}$`).exec(piece);
      if (m) { toks.push(["img", parseInt(m[1], 10)]); continue; }
      let pos = 0;
      for (const om of piece.matchAll(OPTION_MARK_G)) {
        if (piece.slice(pos, om.index).trim()) toks.push(["text", piece.slice(pos, om.index)]);
        toks.push(["mark", om[1]]);
        pos = om.index + om[0].length;
      }
      if (piece.slice(pos).trim()) toks.push(["text", piece.slice(pos)]);
    }
    return toks;
  }

  // Option pictures sit next to their "(ক)" label: before it or after it, on the same line or in a
  // row above/below the labels. Returns [questionPictures, {letter: [pictures]}].
  function assignImages(toks) {
    const segs = [];
    for (const [kind, val] of toks) {
      const last = segs[segs.length - 1];
      if (kind === "text") segs.push(["text", [val]]);
      else if (last && last[0] === kind) last[1].push(val);
      else segs.push([kind, [val]]);
    }
    const allImgs = toks.filter((t) => t[0] === "img").map((t) => t[1]);
    const first = segs.findIndex((g) => g[0] === "mark");
    if (first < 0) return [allImgs, {}];
    let last = first;
    segs.forEach((g, i) => { if (g[0] === "mark") last = i; });
    if (!segs.slice(first).some((g) => g[0] === "img")) return [allImgs, {}];
    const after = segs.slice(last + 1).some((g) => g[0] === "img");
    const options = {}, used = new Set();
    segs.forEach((g, i) => {
      if (g[0] !== "img") return;
      const imgs = g[1];
      let pairs;
      if (after && i > 0 && segs[i - 1][0] === "mark") {
        const marks = segs[i - 1][1], n = Math.min(imgs.length, marks.length);
        pairs = imgs.slice(0, n).map((im, k) => [im, marks.slice(marks.length - n)[k]]);
      } else if (!after && i + 1 < segs.length && segs[i + 1][0] === "mark") {
        const marks = segs[i + 1][1], n = Math.min(imgs.length, marks.length);
        pairs = imgs.slice(imgs.length - n).map((im, k) => [im, marks.slice(0, n)[k]]);
      } else return;
      for (const [im, mk] of pairs) {
        (options[mk] = options[mk] || []).push(im);
        used.add(im);
      }
    });
    const afterLast = new Set(segs.slice(last + 1).filter((g) => g[0] === "img").flatMap((g) => g[1]));
    return [allImgs.filter((i) => !used.has(i) && !afterLast.has(i)), options]; // a picture after the last option is stray
  }

  // ---------- question parsing ----------
  const toInt = (digits) => parseInt([...digits].map((c) => (BN.includes(c) ? BN.indexOf(c) : c)).join(""), 10);

  function stimulusNumbers(text) {
    const nums = (text.match(new RegExp(`[${BN}0-9]+`, "g")) || []).map(toInt);
    if (nums.length === 2 && /থেকে|[-–—]/.test(text)) {
      const s = new Set();
      for (let n = nums[0]; n <= nums[1]; n++) s.add(n);
      return s;
    }
    return new Set(nums);
  }

  function optionsFollow(pars, i) {
    for (let j = i; j < Math.min(i + 12, pars.length); j++) {
      const t = pars[j].raw;
      if (j > i && QUESTION_START.test(t)) return false;
      if (KA.test(t)) return true;
    }
    return false;
  }

  // board = italic [..] text, else a trailing [..]
  function splitBoard(par, text) {
    const m = /\[[^\[\]]*\]/.exec(par.italic.trim());
    if (m) {
      const board = m[0];
      const idx = text.lastIndexOf(board);
      if (idx >= 0) text = text.slice(0, idx) + text.slice(idx + board.length);
      else text = text.replace(/\s*\[[^\[\]]*\]\s*$/, "");
      return [text.trim(), board];
    }
    const t = BOARD_TAIL.exec(text);
    if (t) return [text.slice(0, t.index).trim(), t[1]];
    return [text.trim(), ""];
  }

  function parseOptions(text, mcq) {
    // strict "(ক)" labels, plus "(গ " with a missing bracket when it is the next letter in order
    const strict = [...text.matchAll(OPTION_MARK_G)].map((m) => ({ index: m.index, 0: m[0], 1: m[1], ok: true }));
    const open = [...text.matchAll(OPTION_OPEN_G)].map((m) => ({ index: m.index, 0: m[0], 1: m[1], ok: false }))
      .filter((m) => !strict.some((s) => s.index === m.index));
    let last = Math.max(-1, ...Object.keys(mcq.options).map((k) => ALL_LETTERS.indexOf(k)));
    const marks = [];
    for (const m of [...strict, ...open].sort((a, b) => a.index - b.index)) {
      const idx = ALL_LETTERS.indexOf(m[1]);
      if (m.ok || idx === last + 1) { marks.push(m); last = idx; }
    }
    marks.forEach((m, i) => {
      const end = i + 1 < marks.length ? marks[i + 1].index : text.length;
      mcq.options[m[1]] = text.slice(m.index + m[0].length, end).trim();
    });
  }

  // An answer table ("SL | Ans | SL | Ans ..." with serial numbers and ক/খ/গ/ঘ) -> {serial: letter}, or null
  function answerKey(tbl) {
    const map = {};
    let pairs = 0, cells = 0, header = false;
    for (const tr of kids(tbl)) {
      if (local(tr) !== "tr") continue;
      const texts = kids(tr).filter((c) => local(c) === "tc").map((tc) =>
        Array.from(tc.getElementsByTagNameNS(W, "t")).map((t) => t.textContent).join("").replace(/[\u200b-\u200d\u2060\ufeff]/g, "").trim());
      cells += texts.length;
      for (let i = 0; i < texts.length; i++) {
        if (/^(sl|ans|serial|ক্রম|উত্তর)\.?$/i.test(texts[i])) { header = true; continue; }
        const letter = (texts[i + 1] || "").match(/^\(?\s*([কখগঘঙ])\s*\)?$/);
        if (new RegExp(`^[${BN}0-9]+$`).test(texts[i]) && letter) {
          map[toInt(texts[i])] = letter[1];
          pairs++;
          i++;
        }
      }
    }
    return (header && pairs >= 1) || (pairs >= 3 && pairs * 2 >= cells * 0.8) ? map : null;
  }

  // block-level items in document order: paragraphs (tables are flattened) and answer tables
  function* iterBlocks(node) {
    for (const c of kids(node)) {
      const t = local(c);
      if (c.namespaceURI !== W) continue;
      if (t === "p") yield { p: c };
      else if (t === "tbl") {
        const key = answerKey(c);
        if (key) yield { key };
        else {
          for (const tr of kids(c)) {
            if (local(tr) !== "tr") continue;
            for (const tc of kids(tr)) if (local(tc) === "tc") yield* iterBlocks(tc);
          }
        }
      } else if (t === "sdt") {
        const content = kid(c, "sdtContent", W);
        if (content) yield* iterBlocks(content);
      }
    }
  }

  function paragraphInfo(p, rels, registry, styleNames) {
    const raw = paraText(p, rels, registry).trim();
    // italic text of the paragraph's own runs (used to find the board)
    let italic = "";
    for (const r of kids(p)) {
      if (local(r) !== "r" || r.namespaceURI !== W) continue;
      const i = kid(kid(r, "rPr", W), "i", W);
      if (!i) continue;
      const v = i.getAttributeNS(W, "val");
      if (v === "0" || v === "false") continue;
      for (const x of kids(r)) {
        if (local(x) === "t") italic += x.textContent;
        else if (local(x) === "tab") italic += "\t";
      }
    }
    const ps = kid(kid(p, "pPr", W), "pStyle", W);
    const id = ps ? ps.getAttributeNS(W, "val") : "";
    const style = (styleNames[id] || id || "").trim();
    const hasMath = kids(p).some((c) => local(c) === "oMath" || local(c) === "oMathPara");
    // text boxes anchored in this paragraph (e.g. a data panel beside an explanation); skip the
    // duplicate copy Word keeps in mc:Fallback
    const boxes = [];
    for (const tb of Array.from(p.getElementsByTagNameNS(W, "txbxContent"))) {
      let anc = tb.parentNode, fallback = false;
      while (anc && anc !== p) { if (local(anc) === "Fallback") { fallback = true; break; } anc = anc.parentNode; }
      if (fallback) continue;
      const t = kids(tb).filter((c) => local(c) === "p").map((q) => paraText(q, rels, null).trim()).filter(Boolean).join(" ");
      if (t) boxes.push(t);
    }
    return { raw, italic, style, hasMath, box: boxes.join(" ") };
  }

  // chapterOf(sectPr) -> chapter text from that section's page header (optional)
  function parse(bodyEl, rels, styleNames, chapterOf) {
    const registry = [];
    const pars = [];
    let section = 0;
    const sectPrs = []; // sectPrs[n] = the sectPr that closes section n (its page headers give the chapter)
    for (const item of iterBlocks(bodyEl)) {
      if (item.key) { pars.push({ raw: "\ue002", italic: "", style: "", key: item.key, section }); continue; }
      const info = paragraphInfo(item.p, rels, registry, styleNames || {});
      info.section = section;
      if (info.raw) pars.push(info);
      const sp = kid(kid(item.p, "pPr", W), "sectPr", W);
      if (sp) sectPrs[section++] = sp; // this paragraph ends the section
    }
    const lastSp = kid(bodyEl, "sectPr", W);
    if (lastSp) sectPrs[section] = lastSp;
    const chapters = [];
    let prevChapter = "";
    for (let n = 0; n <= section; n++) {
      chapters[n] = (chapterOf && sectPrs[n] ? chapterOf(sectPrs[n]) : "") || prevChapter; // no header of its own: inherit
      prevChapter = chapters[n];
    }

    // The "উত্তরমালা" banner is a picture printed just before an answer table. Pictures (and the
    // notes) between the last question and the table are not part of any question.
    pars.forEach((p, k) => {
      if (!p.key) return;
      for (let j = k - 1; j >= 0; j--) {
        const t = pars[j].raw.replace(IMG_RE_G, "").trim();
        if (t && !SECTION_MARKER.test(t) && !/^উত্তরমালা/.test(t)) break;
        pars[j].banner = true;
      }
    });

    let mcqs = [], cur = null, mode = null; // mode: title | options | answer | explanation
    let stimLines = [], stim = null;
    let prefixLines = [], prefixAt = -2; // equation lines right after a question's options belong to the NEXT question
    const newMCQ = (num) => ({ num, chapter: "", title: "", board: "", options: {}, answer: "", explanation: "", tokens: [], boxes: [] });

    pars.forEach((par, i) => {
      const raw = par.raw;
      if (par.banner) { cur = null; mode = null; return; }
      if (par.key) {
        // answer table: fill in the answers of the block of questions right above it. A block is the
        // run of unanswered questions numbered 1..N going backwards from the table; an older block
        // that never got a table must not use up these serial numbers.
        const block = [];
        let prev = Infinity;
        for (let k = mcqs.length - 1; k >= 0; k--) {
          const q = mcqs[k];
          if (q.num > prev) break; // the numbering restarted: an earlier block (questions with an answer line are walked over)
          block.unshift(q);
          prev = q.num;
          if (q.num === 1) break;
        }
        const used = new Set();
        for (const q of block) {
          if (q.num in par.key && !used.has(q.num)) { q.answer = par.key[q.num]; used.add(q.num); }
        }
        cur = null; mode = null; stimLines = []; stim = null; prefixLines = [];
        return;
      }
      let text = raw.replace(IMG_RE_G, "").trim();
      // "(ঘ) পণ্যের দাম উত্তর: ক": the answer is written at the end of the last option's line
      let inlineAnswer = "";
      if (!/^উত্তর/.test(text) && OPTION_MARK.test(text)) {
        const mid = /\s*উত্তর\s*[:：]\s*\(?\s*([কখগঘঙ])\s*\)?\s*$/.exec(text);
        if (mid) { inlineAnswer = mid[1]; text = text.slice(0, mid.index).trim(); }
      }
      if (SECTION_MARKER.test(text)) return; // stray "mcq (end)" style notes carry no meaning
      if (/^উত্তরমালা/.test(text)) { cur = null; mode = null; return; } // "answer key" banner
      if (/^(heading|title)/i.test(par.style)) { // chapter/section heading ends the current MCQ
        cur = null; mode = null; stimLines = []; stim = null;
        return;
      }

      // things that cannot belong to an MCQ end it (and its explanation)
      if (cur && ["explanation", "answer", "options"].includes(mode)) {
        const ends =
          (text.length <= 60 && SECTION_HEADING.test(text)) ||
          (mode === "explanation" && text.startsWith("উত্তর")) ||
          (QUESTION_START.test(text) &&
            pars.slice(i + 1, i + 4).some((p) => p.raw.startsWith("উত্তর")) &&
            !pars.slice(i, i + 4).some((p) => KA.test(p.raw)));
        if (ends) { cur = null; mode = null; return; }
      }

      if (STIMULUS_START.test(text)) {
        cur = null; mode = null;
        // the header sentence ends at "দাও:"; the passage may follow in the same paragraph
        const hd = /^(.*?উত্তর\s*দাও\s*[:ঃ]?)\s*(.*)$/s.exec(text);
        stimLines = hd ? [hd[1], ...(hd[2] ? [hd[2]] : [])] : [text];
        stim = null;
        return;
      }

      const qm = QUESTION_START.exec(text);
      if (qm && optionsFollow(pars, i)) {
        const num = toInt(qm[1]);
        if (stimLines.length) {
          const nums = stimulusNumbers(stimLines[0]);
          // a board tag printed at the end of the passage belongs in the Board cell, not in the passage
          let stimBoard = "";
          const lastLine = stimLines[stimLines.length - 1], bt = stimLines.length > 1 ? BOARD_TAIL.exec(lastLine) : null;
          if (bt) { stimBoard = bt[1]; stimLines[stimLines.length - 1] = lastLine.slice(0, bt.index).trim(); }
          stim = { text: stimLines.join("\n"), nums: nums.size ? nums : new Set([num]), board: stimBoard };
          stimLines = [];
        }
        cur = newMCQ(num);
        cur.chapter = chapters[par.section] || "";
        mcqs.push(cur);
        cur.tokens.push(...tokens(raw));
        const [stripped, board] = splitBoard(par, text);
        cur.board = board;
        cur.title = stripped.replace(QUESTION_START, "").trim();
        if (inlineAnswer) cur.answer = inlineAnswer;
        const firstMark = OPTION_MARK.exec(cur.title); // options written on the title line itself
        if (firstMark) {
          parseOptions(cur.title.slice(firstMark.index), cur);
          cur.title = cur.title.slice(0, firstMark.index).trim();
          mode = "options";
        } else mode = "title";
        if (prefixLines.length && prefixAt === i - 1) cur.title = prefixLines.join("\n") + "\n" + cur.title;
        prefixLines = [];
        if (stim && stim.nums.has(num)) { cur.title = stim.text + "\n" + cur.title; cur.board = cur.board || stim.board; }
        if (stim && num >= Math.max(...stim.nums)) stim = null;
        return;
      }

      if (stimLines.length) { // passage text belonging to the stimulus
        if (text) stimLines.push(text);
        return;
      }
      if (!cur) return;

      if (par.box) cur.boxes.push(par.box);
      if (mode === "options" && par.hasMath && !OPTION_MARK.test(text) && LETTERS.every((k) => k in cur.options) &&
          !/^(উত্তর|ব্যাখ্যা)/.test(text)) {
        prefixLines.push(text); // e.g. a reaction scheme printed before the next question
        prefixAt = i;
        return;
      }

      // remember pictures and option labels in order; sorted out per question later
      if ((mode === "title" || mode === "options") && !text.startsWith("উত্তর")) {
        cur.tokens.push(...tokens(raw));
      }
      if (!text) return;

      if (mode !== "explanation") {
        const a = ANSWER.exec(text);
        if (a && text.trimStart().startsWith("উত্তর")) {
          cur.answer = a[1];
          const rest = text.slice(a.index + a[0].length);
          const e = EXPLANATION_ANY.exec(rest);
          if (e) {
            cur.explanation = rest.slice(e.index + e[0].length).trim();
            mode = "explanation";
          } else mode = "answer";
          return;
        }
      }

      const e = EXPLANATION_START.exec(text);
      if (e && mode !== "explanation") {
        cur.explanation = text.slice(e[0].length).trim();
        mode = "explanation";
        return;
      }

      if (mode === "explanation") {
        cur.explanation = (cur.explanation + " " + text).trim();
      } else if (OPTION_MARK.test(text)) {
        if (mode === "options" && OPTION_MARK.exec(text)[1] in cur.options) { // labels start over: not this question
          cur = null; mode = null;
          return;
        }
        parseOptions(text, cur);
        if (inlineAnswer) cur.answer = inlineAnswer;
        mode = "options";
      } else if (mode === "title") { // title wrapped over several paragraphs
        const [t, b] = splitBoard(par, text);
        if (t) cur.title = cur.title ? cur.title + "\n" + t : t;
        cur.board = cur.board || b;
      } else if (mode === "options" && Object.keys(cur.options).length) { // continuation of the last option
        const last = ALL_LETTERS.filter((k) => k in cur.options).pop();
        cur.options[last] = (cur.options[last] + " " + text).trim();
      }
    });

    // keep only complete MCQs: a title, at least two options and an answer
    const complete = mcqs.filter((q) => q.title && Object.keys(q.options).length >= 2);
    const skipped = complete.filter((q) => !q.answer).length; // no "উত্তর:" line and not in any answer table
    mcqs = complete;
    for (const q of mcqs) {
      if (q.boxes.length) q.explanation = [q.explanation, ...q.boxes].filter(Boolean).join(" ");
      delete q.boxes;
      const [question, options] = assignImages(q.tokens);
      q.questionImages = question;
      q.optionImages = options;
      delete q.tokens;
    }
    return { mcqs, registry, skipped, keys: pars.filter((p) => p.key).length };
  }

  // ---------- sheet rows ----------
  const HEADERS = ["status", "SI", "type", "chapter", "Board", "title", "image", "explanation"];
  for (let i = 1; i <= 4; i++) HEADERS.push(`options_${i}_answer`, `options_${i}_is_correct`, `options_${i}_image`);

  // `images`: {question: dataUrl, options: {letter: dataUrl}} per MCQ (filled by the page)
  function toRows(mcqs) {
    return mcqs.map((q, i) => {
      const hasImage = !!q.image || Object.values(q.optionImageUrls || {}).some(Boolean);
      // "failed": no answer found for it, or fewer than four options were read
      const failed = !LETTERS.includes(q.answer) || LETTERS.some((l) => q.options[l] === undefined && !(q.optionImageUrls || {})[l]);
      const row = [failed ? "failed" : "", i + 1, hasImage ? "m4" : "m1", q.chapter || "", q.board, q.title, q.image || "", q.explanation];
      for (const l of LETTERS) {
        const picture = (q.optionImageUrls || {})[l] || "";
        let text = q.options[l] === undefined ? "" : q.options[l];
        if (!text && picture) text = `(${l})`; // a picture-only option still gets its label, so the answer cell is never empty
        row.push(text, q.answer === l ? 1 : 0, picture);
      }
      return row;
    });
  }

  // building blocks shared with the short-question (SQ) scraper
  const internals = { iterBlocks, paragraphInfo, splitBoard, kid, QUESTION_START, KA, IMG_RE_G, W };
  const api = { parse, toRows, HEADERS, LETTERS, paraText, tex, assignImages, tokens, internals };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.MCQParser = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
