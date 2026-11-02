/**
 * AI Rate Checker — local heuristic detector (not an official academic verdict)
 */
(function () {
  "use strict";

  // ——— AI boilerplate phrases (ZH + EN; keep ZH for Chinese docs) ———
  const AI_PHRASES_ZH = [
    "综上所述", "总而言之", "值得注意的是", "需要指出的是", "不可否认",
    "众所周知", "在当今社会", "随着科技的发展", "随着时代的发展",
    "具有重要意义", "发挥着重要作用", "至关重要", "不可或缺",
    "不仅如此", "与此同时", "此外", "因此", "由此可见",
    "进一步来说", "换言之", "具体而言", "从本质上讲",
    "为了更好地", "在一定程度上", "在此基础上", "归根结底",
    "毫无疑问", "显而易见", "不言而喻", "尤为重要",
    "本文将", "本章将", "首先", "其次", "最后", "总之",
    "一方面", "另一方面", "与此相对", "与此相关",
    "深入探讨", "进行讨论", "展开分析", "加以阐述",
    "提供了新的视角", "具有深远影响", "带来了挑战与机遇",
    "我们应当", "有必要", "必须认识到", "应当重视",
  ];

  const AI_PHRASES_EN = [
    "in conclusion", "to summarize", "it is important to note",
    "it should be noted", "it is worth noting", "as a matter of fact",
    "in today's world", "in this day and age", "with the rapid development",
    "plays a crucial role", "plays an important role", "is of great significance",
    "furthermore", "moreover", "additionally", "in addition",
    "on the one hand", "on the other hand", "that being said",
    "needless to say", "it goes without saying", "last but not least",
    "in light of", "taking into account", "from this perspective",
    "this paper will", "this essay will", "this article aims",
    "delve into", "shed light on", "pave the way",
    "a double-edged sword", "a myriad of", "plethora of",
  ];

  const CIRC = 2 * Math.PI * 58; // gauge circumference

  // ——— DOM ———
  const $ = (id) => document.getElementById(id);
  const dropzone = $("dropzone");
  const fileInput = $("fileInput");
  const fileInfo = $("fileInfo");
  const fileName = $("fileName");
  const fileSize = $("fileSize");
  const clearFile = $("clearFile");
  const analyzeBtn = $("analyzeBtn");
  const analyzePasteBtn = $("analyzePasteBtn");
  const sampleBtn = $("sampleBtn");
  const pasteArea = $("pasteArea");
  const loading = $("loading");
  const results = $("results");
  const errorEl = $("error");
  const themeBtn = $("themeBtn");

  let currentFile = null;
  let extractedText = "";

  // ——— Theme ———
  function applyTheme(mode) {
    const html = document.documentElement;
    html.classList.remove("light", "auto");
    if (mode === "light") html.classList.add("light");
    else if (mode === "dark") { /* default dark vars */ }
    else html.classList.add("auto");
    localStorage.setItem("ai-rate-theme", mode);
  }
  (function initTheme() {
    const saved = localStorage.getItem("ai-rate-theme") || "auto";
    applyTheme(saved === "dark" ? "dark" : saved === "light" ? "light" : "auto");
  })();
  themeBtn.addEventListener("click", () => {
    const html = document.documentElement;
    if (html.classList.contains("light")) applyTheme("dark");
    else if (!html.classList.contains("auto") && !html.classList.contains("light")) applyTheme("light");
    else applyTheme(window.matchMedia("(prefers-color-scheme: dark)").matches ? "light" : "dark");
  });

  // ——— UI helpers ———
  function showError(msg) {
    errorEl.textContent = msg;
    errorEl.classList.add("visible");
  }
  function clearError() {
    errorEl.classList.remove("visible");
    errorEl.textContent = "";
  }
  function setLoading(on) {
    loading.classList.toggle("visible", on);
    if (on) results.classList.remove("visible");
  }
  function formatBytes(n) {
    if (n < 1024) return n + " B";
    if (n < 1024 * 1024) return (n / 1024).toFixed(1) + " KB";
    return (n / (1024 * 1024)).toFixed(1) + " MB";
  }
  function setFile(file) {
    currentFile = file;
    extractedText = "";
    if (!file) {
      fileInfo.classList.remove("visible");
      analyzeBtn.disabled = true;
      return;
    }
    fileName.textContent = file.name;
    fileSize.textContent = formatBytes(file.size);
    fileInfo.classList.add("visible");
    analyzeBtn.disabled = false;
    clearError();
  }

  // ——— File / drop ———
  dropzone.addEventListener("click", () => fileInput.click());
  dropzone.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); fileInput.click(); }
  });
  fileInput.addEventListener("change", () => {
    const f = fileInput.files && fileInput.files[0];
    if (f) acceptFile(f);
  });
  ["dragenter", "dragover"].forEach((ev) => {
    dropzone.addEventListener(ev, (e) => {
      e.preventDefault();
      dropzone.classList.add("dragover");
    });
  });
  ["dragleave", "drop"].forEach((ev) => {
    dropzone.addEventListener(ev, (e) => {
      e.preventDefault();
      dropzone.classList.remove("dragover");
    });
  });
  dropzone.addEventListener("drop", (e) => {
    const f = e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) acceptFile(f);
  });
  clearFile.addEventListener("click", () => {
    fileInput.value = "";
    setFile(null);
  });

  function acceptFile(file) {
    const name = (file.name || "").toLowerCase();
    if (!name.endsWith(".docx")) {
      showError("Please upload a .docx Word file (legacy .doc is not supported—save as .docx first).");
      setFile(null);
      return;
    }
    if (file.size > 15 * 1024 * 1024) {
      showError("File is too large (please keep under 15MB).");
      setFile(null);
      return;
    }
    setFile(file);
  }

  // ——— Docx extract ———
  async function extractDocx(file) {
    if (typeof mammoth === "undefined") {
      throw new Error("mammoth.js failed to load; cannot parse Word.");
    }
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.extractRawText({ arrayBuffer });
    const text = (result.value || "").replace(/\r\n/g, "\n").trim();
    if (!text) throw new Error("No text could be extracted (empty document or image-only).");
    return text;
  }

  // ——— Text utilities ———
  function isCJK(ch) {
    return /[\u4e00-\u9fff\u3400-\u4dbf]/.test(ch);
  }

  function splitSentences(text) {
    const cleaned = text.replace(/\u00a0/g, " ").trim();
    if (!cleaned) return [];
    // Split on CJK / EN sentence enders while keeping content
    const parts = cleaned.split(/(?<=[。！？；.!?;])\s*|\n+/);
    return parts
      .map((s) => s.trim())
      .filter((s) => s.length >= 4);
  }

  function tokenize(text) {
    const tokens = [];
    let buf = "";
    for (const ch of text) {
      if (isCJK(ch)) {
        if (buf) { tokens.push(buf.toLowerCase()); buf = ""; }
        tokens.push(ch);
      } else if (/[A-Za-z0-9']/.test(ch)) {
        buf += ch;
      } else {
        if (buf) { tokens.push(buf.toLowerCase()); buf = ""; }
      }
    }
    if (buf) tokens.push(buf.toLowerCase());
    return tokens.filter(Boolean);
  }

  function charLen(s) {
    return [...s].length;
  }

  function mean(arr) {
    if (!arr.length) return 0;
    return arr.reduce((a, b) => a + b, 0) / arr.length;
  }

  function stddev(arr) {
    if (arr.length < 2) return 0;
    const m = mean(arr);
    return Math.sqrt(arr.reduce((s, x) => s + (x - m) ** 2, 0) / arr.length);
  }

  function clamp(x, lo, hi) {
    return Math.max(lo, Math.min(hi, x));
  }

  // ——— Features ———
  function detectPhrases(text) {
    const lower = text.toLowerCase();
    const hits = [];
    for (const p of AI_PHRASES_ZH) {
      let idx = 0;
      while ((idx = text.indexOf(p, idx)) !== -1) {
        hits.push({ phrase: p, index: idx });
        idx += p.length;
      }
    }
    for (const p of AI_PHRASES_EN) {
      let idx = 0;
      while ((idx = lower.indexOf(p, idx)) !== -1) {
        hits.push({ phrase: p, index: idx });
        idx += p.length;
      }
    }
    return hits;
  }

  function typeTokenRatio(tokens) {
    if (!tokens.length) return 1;
    return new Set(tokens).size / tokens.length;
  }

  function punctuationUniformity(sentences) {
    // AI text often ends almost every sentence with 。 or .
    // Chinese human prose also mostly uses 。 — only trust this on longer docs.
    if (sentences.length < 8) return 0;
    const endings = sentences.map((s) => {
      const t = s.trim();
      return t[t.length - 1] || "";
    });
    const counts = {};
    endings.forEach((e) => { counts[e] = (counts[e] || 0) + 1; });
    const maxShare = Math.max(...Object.values(counts)) / endings.length;
    // High uniformity of period → higher AI signal
    return clamp((maxShare - 0.7) / 0.3, 0, 1) * 0.6;
  }

  function burstiness(lengths) {
    // Human writing often has higher variance (burstiness); AI is smoother
    if (lengths.length < 6) return 0.25;
    const m = mean(lengths);
    if (m < 1) return 0.25;
    const cv = stddev(lengths) / m; // coefficient of variation
    // Typical AI CV often lower; map low CV → high AI score
    // CV ~0.2 → very even (AI-ish), CV ~0.7+ → human-ish
    return clamp(1 - (cv - 0.2) / 0.5, 0, 1);
  }

  function paragraphEvenness(text) {
    const paras = text.split(/\n+/).map((p) => p.trim()).filter((p) => charLen(p) > 20);
    if (paras.length < 3) return 0.3;
    const lens = paras.map(charLen);
    const cv = stddev(lens) / (mean(lens) || 1);
    return clamp(1 - (cv - 0.1) / 0.6, 0, 1);
  }

  function listDensity(text) {
    const lines = text.split(/\n+/);
    if (!lines.length) return 0;
    const listish = lines.filter((l) =>
      /^\s*([（(]?[0-9一二三四五六七八九十]+[）).、．]|[-*•·]|第[一二三四五六七八九十\d]+[章节条款点])/.test(l)
    ).length;
    const ratio = listish / lines.length;
    // Moderate lists OK; very high enumeration density is LLM-ish
    return clamp((ratio - 0.05) / 0.35, 0, 1);
  }

  function connectorDensity(text, sentences) {
    const connectors = [
      "因此", "然而", "此外", "同时", "另外", "所以", "但是", "不过",
      "并且", "以及", "从而", "进而", "故而", "于是",
      "however", "therefore", "moreover", "furthermore", "thus",
      "hence", "nevertheless", "consequently", "additionally",
    ];
    const lower = text.toLowerCase();
    let count = 0;
    for (const c of connectors) {
      const src = /[a-z]/.test(c) ? lower : text;
      let idx = 0;
      while ((idx = src.indexOf(c, idx)) !== -1) {
        count++;
        idx += c.length;
      }
    }
    const perSent = count / Math.max(sentences.length, 1);
    // ~0.15–0.25 natural; >0.5 often template-heavy
    return clamp((perSent - 0.15) / 0.55, 0, 1);
  }

  function scoreSentence(sent, phraseHitsInSent) {
    let score = 0;
    const reasons = [];
    const len = charLen(sent);

    if (phraseHitsInSent.length) {
      score += Math.min(0.45, 0.18 * phraseHitsInSent.length);
      reasons.push("Boilerplate: " + phraseHitsInSent.map((h) => h.phrase).slice(0, 3).join(", "));
    }

    // Very mid-length, polished single clause feel
    if (len >= 40 && len <= 120) {
      score += 0.08;
    }

    // Parallel structure / 「不仅…而且…」
    if (/不仅.{2,20}(而且|还)|一方面.{2,30}另一方面|既.{1,15}又/.test(sent)) {
      score += 0.15;
      reasons.push("Parallel / balanced structure");
    }

    // English AIisms
    const low = sent.toLowerCase();
    if (/\b(delve|crucial|underscore|multifaceted|landscape|tapestry)\b/.test(low)) {
      score += 0.2;
      reasons.push("Common English AI buzzwords");
    }

    // Ending with abstract evaluation
    if (/(具有重要意义|至关重要|不可或缺|深远影响|挑战与机遇)[。！.!]?$/.test(sent)) {
      score += 0.2;
      reasons.push("Abstract closing cliché");
    }

    // Extremely uniform punctuation-only end + long
    if (len > 80 && /[。．.]$/.test(sent.trim()) && (sent.match(/[，,]/g) || []).length >= 4) {
      score += 0.1;
      reasons.push("Long sentence with even commas");
    }

    const tokens = tokenize(sent);
    if (tokens.length >= 8) {
      const ttr = typeTokenRatio(tokens);
      if (ttr < 0.55) {
        score += 0.12;
        reasons.push("Low lexical variety");
      }
    }

    return { score: clamp(score, 0, 1), reasons };
  }

  function jaccard(a, b) {
    const A = new Set(tokenize(a));
    const B = new Set(tokenize(b));
    if (!A.size || !B.size) return 0;
    let inter = 0;
    for (const x of A) if (B.has(x)) inter++;
    return inter / (A.size + B.size - inter);
  }

  function findDuplicates(sentences) {
    const pairs = [];
    const n = Math.min(sentences.length, 120); // cap for perf
    for (let i = 0; i < n; i++) {
      for (let j = i + 1; j < n; j++) {
        if (Math.abs(charLen(sentences[i]) - charLen(sentences[j])) > 40) continue;
        const sim = jaccard(sentences[i], sentences[j]);
        if (sim >= 0.72) {
          pairs.push({ i, j, sim, a: sentences[i], b: sentences[j] });
        }
      }
    }
    pairs.sort((x, y) => y.sim - x.sim);
    return pairs.slice(0, 12);
  }

  // ——— Main analysis ———
  function analyze(text) {
    const sentences = splitSentences(text);
    if (sentences.length < 2 && charLen(text) < 40) {
      throw new Error("Text is too short for a reliable estimate (need several full sentences).");
    }

    const tokens = tokenize(text);
    const phraseHits = detectPhrases(text);
    const lengths = sentences.map(charLen);

    const feat = {
      phrase: clamp(phraseHits.length / Math.max(sentences.length * 0.35, 1), 0, 1),
      burstiness: burstiness(lengths),
      ttr: (() => {
        const ttr = typeTokenRatio(tokens);
        // Low TTR → higher AI-ish (esp. long docs). Short docs naturally high TTR.
        if (tokens.length < 40) return 0.35;
        return clamp((0.72 - ttr) / 0.35, 0, 1);
      })(),
      punct: punctuationUniformity(sentences),
      para: paragraphEvenness(text),
      list: listDensity(text),
      connector: connectorDensity(text, sentences),
    };

    // Weighted aggregate
    const weights = {
      phrase: 0.34,
      burstiness: 0.14,
      ttr: 0.12,
      punct: 0.06,
      para: 0.08,
      list: 0.08,
      connector: 0.18,
    };

    let weighted = 0;
    for (const k of Object.keys(weights)) weighted += feat[k] * weights[k];
    // Extra boost when many distinct boilerplate hits
    if (phraseHits.length >= 8) weighted = clamp(weighted + 0.08, 0, 1);
    if (phraseHits.length >= 15) weighted = clamp(weighted + 0.06, 0, 1);

    // Sentence-level boost
    const sentScores = sentences.map((s) => {
      const hits = phraseHits.filter((h) => {
        const slice = text.slice(Math.max(0, h.index - 2), h.index + h.phrase.length + 2);
        return s.includes(h.phrase) || slice.includes(s.slice(0, Math.min(12, s.length)));
      });
      // Simpler: phrases contained in sentence
      const inSent = phraseHits.filter((h) => s.includes(h.phrase));
      // dedupe by phrase
      const uniq = [];
      const seen = new Set();
      for (const h of inSent) {
        if (!seen.has(h.phrase)) { seen.add(h.phrase); uniq.push(h); }
      }
      return { text: s, ...scoreSentence(s, uniq) };
    });

    const topSentMean = mean(
      [...sentScores].sort((a, b) => b.score - a.score).slice(0, Math.max(3, Math.ceil(sentScores.length * 0.25))).map((x) => x.score)
    );

    let aiRate = clamp(weighted * 0.68 + topSentMean * 0.32, 0, 1);

    // Soften short / phrase-free docs (avoid false positives on diary-like notes)
    if (tokens.length < 80) aiRate *= 0.75;
    if (phraseHits.length === 0 && tokens.length < 200) aiRate *= 0.7;
    if (phraseHits.length === 0) aiRate = Math.min(aiRate, 0.55);

    const pct = Math.round(aiRate * 100);
    const duplicates = findDuplicates(sentences);

    const featureRows = [
      { key: "phrase", name: "AI boilerplate / template phrases", score: feat.phrase, detail: `${phraseHits.length} hit(s)` },
      { key: "burstiness", name: "Even sentence length (low burstiness)", score: feat.burstiness, detail: `Length SD ${stddev(lengths).toFixed(1)}` },
      { key: "ttr", name: "Low lexical diversity (TTR)", score: feat.ttr, detail: `TTR ${typeTokenRatio(tokens).toFixed(3)}` },
      { key: "connector", name: "Discourse-connector density", score: feat.connector, detail: "Template-like transitions" },
      { key: "para", name: "Paragraph length evenness", score: feat.para, detail: "Whether paragraphs are overly uniform" },
      { key: "list", name: "List / enumeration density", score: feat.list, detail: "LLM-style bulletization" },
      { key: "punct", name: "Uniform sentence endings", score: feat.punct, detail: "Highly consistent end punctuation" },
    ];

    return {
      pct,
      aiRate,
      sentences: sentences.length,
      chars: charLen(text),
      tokens: tokens.length,
      phraseCount: phraseHits.length,
      featureRows,
      sentScores,
      duplicates,
      ttr: typeTokenRatio(tokens),
    };
  }

  function riskLevel(pct) {
    if (pct < 35) return { cls: "risk-low", label: "Low risk", title: "More human-like signals", summary: "Statistical features look closer to human writing, but you should still read the text yourself." };
    if (pct < 65) return { cls: "risk-mid", label: "Medium risk", title: "Mixed or polished traces", summary: "Some template / uniformity signals—possibly AI-assisted or highly formulaic writing. Check highlighted sentences." };
    return { cls: "risk-high", label: "High risk", title: "More AI-like signals", summary: "Several heuristic indicators are elevated (boilerplate, even sentence length, dense connectors, etc.). Local estimate only—not an official similarity verdict." };
  }

  function barColor(score) {
    if (score >= 0.65) return "var(--danger)";
    if (score >= 0.4) return "var(--warning)";
    return "var(--accent)";
  }

  function render(report) {
    const risk = riskLevel(report.pct);
    $("scorePct").textContent = report.pct + "%";
    $("scoreTitle").textContent = risk.title;
    const badge = $("riskBadge");
    badge.className = "risk " + risk.cls;
    badge.textContent = risk.label;
    $("scoreSummary").textContent = risk.summary + " (Heuristic estimate—not an official academic verdict.)";

    const fill = $("gaugeFill");
    const offset = CIRC * (1 - report.pct / 100);
    fill.style.strokeDasharray = String(CIRC);
    fill.style.strokeDashoffset = String(offset);
    fill.setAttribute("stroke", report.pct >= 65 ? "#ef4444" : report.pct >= 35 ? "#f59e0b" : "#22c55e");

    $("statsRow").innerHTML = [
      ["Sentences", report.sentences],
      ["Characters", report.chars],
      ["Tokens", report.tokens],
      ["Boilerplate hits", report.phraseCount],
      ["TTR", report.ttr.toFixed(3)],
      ["Similar pairs", report.duplicates.length],
    ].map(([k, v]) => `<div class="stat"><div class="k">${k}</div><div class="v">${v}</div></div>`).join("");

    $("features").innerHTML = report.featureRows.map((f) => {
      const pct = Math.round(f.score * 100);
      return `<div class="feature">
        <div class="feature-head"><span class="name">${f.name}</span><span class="val">${pct}% · ${f.detail}</span></div>
        <div class="bar"><i style="width:${pct}%;background:${barColor(f.score)}"></i></div>
      </div>`;
    }).join("");

    const suspicious = [...report.sentScores]
      .filter((s) => s.score >= 0.25)
      .sort((a, b) => b.score - a.score)
      .slice(0, 15);

    const susEl = $("suspicious");
    if (!suspicious.length) {
      susEl.innerHTML = `<p class="empty-hint">No high-scoring suspicious sentences found (or the text is short).</p>`;
    } else {
      susEl.innerHTML = suspicious.map((s) => {
        const level = s.score >= 0.55 ? "high" : "mid";
        const label = s.score >= 0.55 ? "High" : "Med";
        const reason = s.reasons.length ? s.reasons.join("; ") : "Combined statistical signals";
        return `<div class="sent ${level}"><span class="badge">${label} ${Math.round(s.score * 100)}</span>${escapeHtml(s.text)}<div class="reason">${escapeHtml(reason)}</div></div>`;
      }).join("");
    }

    const dupEl = $("duplicates");
    if (!report.duplicates.length) {
      dupEl.innerHTML = `<p class="empty-hint">No highly similar sentence pairs in this document (threshold 0.72). Simple in-doc check only—not an external corpus.</p>`;
    } else {
      dupEl.innerHTML = report.duplicates.map((d) => {
        return `<div class="sent dup"><span class="badge">Similar ${Math.round(d.sim * 100)}%</span>
          <div>${escapeHtml(d.a)}</div>
          <div class="reason" style="margin-top:0.4rem">⟷ ${escapeHtml(d.b)}</div>
        </div>`;
      }).join("");
    }

    results.classList.add("visible");
    results.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  async function runWithText(text) {
    clearError();
    setLoading(true);
    try {
      // yield to UI
      await new Promise((r) => setTimeout(r, 30));
      const report = analyze(text);
      render(report);
    } catch (err) {
      showError(err.message || String(err));
      results.classList.remove("visible");
    } finally {
      setLoading(false);
    }
  }

  analyzeBtn.addEventListener("click", async () => {
    if (!currentFile) return;
    clearError();
    setLoading(true);
    try {
      extractedText = await extractDocx(currentFile);
      await runWithText(extractedText);
    } catch (err) {
      showError(err.message || String(err));
      results.classList.remove("visible");
      setLoading(false);
    }
  });

  analyzePasteBtn.addEventListener("click", () => {
    const t = pasteArea.value.trim();
    if (!t) {
      showError("Please paste some text first.");
      return;
    }
    runWithText(t);
  });

  const SAMPLE_AI = `In today's world, artificial intelligence plays a crucial role in higher education. It is important to note that digital learning platforms not only improve efficiency but also provide personalized learning experiences. Needless to say, educational equity and quality are of great significance.

On the one hand, intelligent tutoring systems can adapt dynamically to student progress; on the other hand, big data analysis sheds light on teaching decisions. Furthermore, we must recognize that technology is a double-edged sword that brings both challenges and opportunities. Therefore, it is worth noting that a more robust ethical framework is essential.

In conclusion, this paper will delve into the integration of AI and higher education from policy, technology, and practice perspectives. Last but not least, only by balancing humanistic care with innovation can we pave the way for a truly transformative educational landscape.`;

  sampleBtn.addEventListener("click", () => {
    pasteArea.value = SAMPLE_AI;
    $("pasteDetails").open = true;
    runWithText(SAMPLE_AI);
  });

  // Expose for smoke tests
  window.__AIRateChecker = { analyze, extractDocx, splitSentences };
})();
