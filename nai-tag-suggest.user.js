// ==UserScript==
// @name         NAI TAG Helper (iPad)
// @namespace    https://github.com/kazuta-creator/aaa
// @version      0.13.0-ios1
// @description  NovelAI tag suggest for iPad (tap-based, Userscripts.app)
// @match        https://novelai.net/*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(() => {
  "use strict";
  console.log("[NAI TAG iOS] userscript loaded");

  /* =========================================================
   * ▼▼▼ ここだけ編集してください ▼▼▼
   * GitHub Pages にアップロードした CSV の置き場 URL。
   * 末尾スラッシュ必須。例:
   *   https://yourname.github.io/nai-tag-suggest-ipad/csv/
   * ========================================================= */
  const CSV_BASE_URL = "https://kazuta-creator.github.io/aaa/csv/";

  // gzip 版 (.csv.gz) を置いたなら true、生の .csv のままなら false
  const USE_GZIP = false;

  // キャッシュのバージョン。CSV を更新したらここを 2,3... と上げると再DL
  const CACHE_VERSION = 1;

  /* =========================================================
   * ▼ 基本設定
   * ========================================================= */
  const CSV_FILES = [
    "danbooru.csv",
    "danbooru_e621_merged.csv",
    "danbooru_2025.csv",
  ];
  const JP_PRECISE_FILE = "danbooru-precise-jp.csv";
  const JP_FALLBACK_FILE = "danbooru-machine-jp.csv";

  const MAX_RESULTS = 500;
  const CONVERT_UNDERSCORE_TO_SPACE = false;
  const AUTO_HIDE_DELAY_MS = 60000; // iPad ではタップ操作がゆっくりなので長め

  /* =========================================================
   * ▼ state
   * ========================================================= */
  let TAGS = [];
  let TAGS_BY_KEY = new Map();
  let TAG_ALIAS_ENTRIES = [];
  let TAG_ALIASES_BY_TAG_KEY = new Map();
  let TAGS_HITS_CACHE = [];
  let TAGS_JP_MAP = new Map();

  let FAVORITES = new Set();
  let HISTORY = [];

  let activeTab = "all";
  let jpSearchQuery = "";
  let sortMode = "prefix";

  let activeInput = null;
  let savedRange = null; // iOS: 選択範囲を明示的に覚える

  let emphasizeMode = 0;
  let emphasizeValue = 2.0;

  let pinEnabled = false;
  let hideTimer = null;

  /* =========================================================
   * ▼ utils (元コードから踏襲)
   * ========================================================= */
  const normalize = s =>
    s.toLowerCase()
      .replace(/[_\s]+/g, "_")
      .replace(/:+/g, "")
      .replace(/[^a-z0-9_]/g, "");

  const compactForSearch = s => normalize(s).replace(/_/g, "");
  const stripDecorations = s => s.replace(/[-\d.]*::/g, "");

  function parseCSVLine(line) {
    const out = [];
    let cell = "", inQuote = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === "\"") {
        if (inQuote && line[i + 1] === "\"") { cell += "\""; i++; }
        else inQuote = !inQuote;
        continue;
      }
      if (ch === "," && !inQuote) { out.push(cell); cell = ""; continue; }
      cell += ch;
    }
    out.push(cell);
    return out;
  }

  function parseAliasTokens(aliasRaw) {
    if (!aliasRaw) return [];
    return String(aliasRaw).split(",").map(v => v.trim()).filter(v => v);
  }

  function betterAliasChoice(nextAlias, prevAlias, query) {
    if (!prevAlias) return true;
    const nextNorm = normalize(nextAlias);
    const prevNorm = normalize(prevAlias);
    const isNextExact = nextNorm === query;
    const isPrevExact = prevNorm === query;
    if (isNextExact !== isPrevExact) return isNextExact;
    if (nextAlias.length !== prevAlias.length) return nextAlias.length < prevAlias.length;
    return nextAlias < prevAlias;
  }

  function highlightMatch(text, query) {
    if (!query) return text;
    const i = text.toLowerCase().indexOf(query);
    if (i === -1) return text;
    return text.slice(0, i)
      + `<span style="background:#7a6a00;color:#fff6bf;border-radius:2px;padding:0 2px;font-weight:600;">`
      + text.slice(i, i + query.length)
      + `</span>`
      + text.slice(i + query.length);
  }

  function tokenBounds(text, pos) {
    let s = pos;
    while (s > 0 && text[s - 1] !== "," && !/\s/.test(text[s - 1])) s--;
    let e = pos;
    while (e < text.length && text[e] !== "," && !/\s/.test(text[e])) e++;
    return { s, e };
  }

  const containsNonAscii = str => /[^\x00-\x7F]/.test(str);

  function shouldKeepUnderscore(tag) {
    if (/^[\^;TtOo><=][-.'\w]*_[-.'\w]*[\^;TtOo><=]$/.test(tag)) return true;
    return false;
  }
  function convertUnderscoreForDisplay(s) {
    if (!CONVERT_UNDERSCORE_TO_SPACE) return s;
    return shouldKeepUnderscore(s) ? s : s.replace(/_/g, " ");
  }
  const convertUnderscoreForInsert = convertUnderscoreForDisplay;

  function displayTag(tag) {
    if (tag.includes(":")) {
      const [ns, restRaw] = tag.split(":", 2);
      let rest = (restRaw || "").replace(/^_+/, "");
      rest = convertUnderscoreForDisplay(rest);
      return ns + ": " + rest;
    }
    return convertUnderscoreForDisplay(tag);
  }
  function formatInsertTag(tag) {
    if (tag.includes(":")) {
      const [ns, restRaw] = tag.split(":", 2);
      let rest = (restRaw || "").replace(/^_+/, "");
      rest = convertUnderscoreForInsert(rest);
      return ns + ": " + rest;
    }
    return convertUnderscoreForInsert(tag);
  }
  function normalizeJPText(s) {
    return String(s || "").replace(/\s+/g, " ").trim();
  }
  function wikiUrlForTag(tag) {
    return `https://danbooru.donmai.us/wiki_pages/${encodeURIComponent(tag)}`;
  }

  /* =========================================================
   * ▼ ストレージ: お気に入り/履歴は localStorage
   * ========================================================= */
  const LS_FAV = "nai_tag_favorites_v1";
  const LS_HIST = "nai_tag_history_v1";
  const HISTORY_MAX = 50;

  function loadUserLists() {
    try {
      const fav = JSON.parse(localStorage.getItem(LS_FAV) || "[]");
      const hist = JSON.parse(localStorage.getItem(LS_HIST) || "[]");
      FAVORITES = new Set(Array.isArray(fav) ? fav : []);
      HISTORY = Array.isArray(hist) ? hist : [];
    } catch (e) {
      console.warn("[NAI TAG iOS] loadUserLists failed:", e);
      FAVORITES = new Set(); HISTORY = [];
    }
  }
  function saveFavorites() {
    try { localStorage.setItem(LS_FAV, JSON.stringify([...FAVORITES])); } catch {}
  }
  function saveHistory() {
    try { localStorage.setItem(LS_HIST, JSON.stringify(HISTORY)); } catch {}
  }
  function addToHistory(tagKey) {
    HISTORY = HISTORY.filter(k => k !== tagKey);
    HISTORY.unshift(tagKey);
    if (HISTORY.length > HISTORY_MAX) HISTORY.length = HISTORY_MAX;
    saveHistory();
  }

  /* =========================================================
   * ▼ IndexedDB: CSV 本文キャッシュ
   * ========================================================= */
  const IDB_NAME = "naiTagSuggestCache";
  const IDB_STORE = "files";

  function idbOpen() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(IDB_NAME, 1);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains(IDB_STORE)) db.createObjectStore(IDB_STORE);
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }
  function idbGet(key) {
    return idbOpen().then(db => new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE, "readonly");
      const r = tx.objectStore(IDB_STORE).get(key);
      r.onsuccess = () => resolve(r.result);
      r.onerror = () => reject(r.error);
    }));
  }
  function idbPut(key, value) {
    return idbOpen().then(db => new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE, "readwrite");
      tx.objectStore(IDB_STORE).put(value, key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    }));
  }

  async function gunzipToText(arrayBuffer) {
    if (typeof DecompressionStream === "undefined") {
      throw new Error("DecompressionStream 未対応 (iPadOS 16.4 以上が必要)");
    }
    const ds = new DecompressionStream("gzip");
    const stream = new Response(arrayBuffer).body.pipeThrough(ds);
    return await new Response(stream).text();
  }

  async function fetchCsvText(relName) {
    const cacheKey = `v${CACHE_VERSION}:${relName}`;
    const cached = await idbGet(cacheKey).catch(() => null);
    if (cached && typeof cached === "string") {
      return cached;
    }
    const url = CSV_BASE_URL + relName + (USE_GZIP ? ".gz" : "");
    console.log("[NAI TAG iOS] fetching", url);
    const res = await fetch(url, { cache: "no-store", credentials: "omit" });
    if (!res.ok) throw new Error(`HTTP ${res.status} for ${url}`);
    let text;
    if (USE_GZIP) {
      const buf = await res.arrayBuffer();
      text = await gunzipToText(buf);
    } else {
      text = await res.text();
    }
    idbPut(cacheKey, text).catch(e => console.warn("[NAI TAG iOS] idb put failed", e));
    return text;
  }

  /* =========================================================
   * ▼ tag colors
   * ========================================================= */
  const TAG_TYPE_COLORS = {
    0: "#7fe9ff", 1: "#7db7ff", 3: "#ff9f9f",
    4: "#8fe388", 5: "#ffcc66", 7: "#b39ddb", 8: "#9ad1d4"
  };
  const colorByType = id => TAG_TYPE_COLORS[id] || "#aaa";

  /* =========================================================
   * ▼ CSV パース (元コードとほぼ同じ)
   * ========================================================= */
  function parseJPMapText(text) {
    const linesRaw = String(text || "").split(/\r?\n/);
    const map = new Map();
    let parsedRows = 0;
    for (const lineRaw of linesRaw) {
      const line = (lineRaw || "").trim();
      if (!line) continue;
      let parts = null;
      if (line.includes("\t")) parts = line.split("\t");
      else if (line.includes(",")) parts = parseCSVLine(line);
      if (!parts || parts.length < 2) continue;
      if (parts.length >= 3 && /^\d+$/.test((parts[0] || "").trim())) {
        const tag = (parts[1] || "").trim();
        const jp = normalizeJPText(parts.slice(2).join(","));
        if (!tag || !jp) continue;
        map.set(normalize(tag), jp);
        parsedRows++;
        continue;
      }
      const tag = (parts[0] || "").trim();
      const jp = normalizeJPText(parts.slice(1).join(","));
      if (!tag || !jp) continue;
      map.set(normalize(tag), jp);
      parsedRows++;
    }
    if (parsedRows === 0) {
      const compact = linesRaw.map(l => (l || "").trim()).filter(l => l);
      for (let i = 0; i + 2 < compact.length; i += 3) {
        const tag = (compact[i + 1] || "").trim();
        const jp = normalizeJPText(compact[i + 2]);
        if (!tag || !jp) continue;
        map.set(normalize(tag), jp);
      }
    }
    return map;
  }

  async function loadJPFile(path, label) {
    try {
      const text = await fetchCsvText(path);
      const map = parseJPMapText(text);
      console.log(`[NAI TAG iOS] ${label} loaded:`, map.size);
      return map;
    } catch (e) {
      console.warn(`[NAI TAG iOS] ${label} load failed:`, e, "file=", path);
      return new Map();
    }
  }

  function applyJPMapToTags() {
    for (const t of TAGS) t.jp = TAGS_JP_MAP.get(t.key) || "";
  }

  async function loadJPMaps() {
    const preciseMap = await loadJPFile(JP_PRECISE_FILE, "precise jp");
    if (preciseMap.size > 0) { TAGS_JP_MAP = preciseMap; return; }
    const fallbackMap = await loadJPFile(JP_FALLBACK_FILE, "machine jp (fallback)");
    TAGS_JP_MAP = fallbackMap;
  }

  async function loadAllCSV() {
    const map = new Map();
    const aliasEntries = [];
    const aliasesByTag = new Map();
    const aliasPairSet = new Set();

    for (const path of CSV_FILES) {
      let text;
      try { text = await fetchCsvText(path); }
      catch (e) { console.warn("[NAI TAG iOS] csv failed", path, e); continue; }
      for (const line of text.split(/\r?\n/)) {
        if (!line) continue;
        const cols = parseCSVLine(line);
        const tag = (cols[0] || "").trim();
        const typeIdRaw = (cols[1] || "").trim();
        const countRaw = (cols[2] || "").trim();
        const aliasRaw = (cols[3] || "").trim();
        if (!tag) continue;
        const key = normalize(tag);
        const typeId = Number(typeIdRaw);
        const count = Number(countRaw || 0);
        if (!map.has(key)) {
          map.set(key, { tag, key, compactKey: compactForSearch(tag), typeId, count });
        } else {
          map.get(key).count = Math.max(map.get(key).count, count);
        }
        const aliases = parseAliasTokens(aliasRaw);
        if (aliases.length > 0) {
          let tagAliasList = aliasesByTag.get(key);
          if (!tagAliasList) { tagAliasList = []; aliasesByTag.set(key, tagAliasList); }
          for (const alias of aliases) {
            const aliasKey = normalize(alias);
            if (!aliasKey || aliasKey === key) continue;
            const pairKey = `${key}|${aliasKey}`;
            if (aliasPairSet.has(pairKey)) continue;
            aliasPairSet.add(pairKey);
            const aliasEntry = { tagKey: key, alias, aliasKey, aliasCompact: compactForSearch(alias) };
            aliasEntries.push(aliasEntry);
            tagAliasList.push(aliasEntry);
          }
        }
      }
    }
    TAGS = [...map.values()];
    TAGS_BY_KEY = map;
    TAG_ALIAS_ENTRIES = aliasEntries;
    TAG_ALIASES_BY_TAG_KEY = aliasesByTag;
    applyJPMapToTags();
    console.log("[NAI TAG iOS] tags:", TAGS.length, "aliases:", TAG_ALIAS_ENTRIES.length);
  }

  /* =========================================================
   * ▼ UI (iPad 向け: 大きめ / タップ重視)
   * ========================================================= */
  const wrapper = document.createElement("div");
  wrapper.id = "nai-tag-ios";
  wrapper.style.cssText = `
    position:fixed;
    z-index:2147483647;
    background:#111;
    border:1px solid #444;
    border-radius:10px;
    display:none;
    overflow:hidden;
    width:min(560px, 92vw);
    max-height:min(70vh, 640px);
    box-shadow:0 10px 30px rgba(0,0,0,.5);
    touch-action:manipulation;
    font-family:-apple-system,BlinkMacSystemFont,sans-serif;
    color:#ddd;
  `;

  const header = document.createElement("div");
  header.style.cssText = `
    background:#1a1a1a;
    padding:8px;
    display:flex;
    flex-direction:column;
    gap:8px;
    touch-action:none;
  `;

  /* row1: emphasize + pin + close */
  const row1 = document.createElement("div");
  row1.style.cssText = "display:flex;align-items:center;gap:6px;";

  const btnStrong = document.createElement("button");
  btnStrong.textContent = "強調";
  const btnWeak = document.createElement("button");
  btnWeak.textContent = "弱め";
  const btnMinus = document.createElement("button");
  btnMinus.textContent = "−";
  const valueLabel = document.createElement("span");
  const btnPlus = document.createElement("button");
  btnPlus.textContent = "+";
  const pinBtn = document.createElement("button");
  pinBtn.textContent = "📌";
  const closeBtn = document.createElement("button");
  closeBtn.textContent = "×";

  valueLabel.style.cssText = "color:#ddd;font-size:14px;min-width:32px;text-align:center;";

  row1.append(btnStrong, btnWeak, btnMinus, valueLabel, btnPlus);
  const spacer1 = document.createElement("div");
  spacer1.style.flex = "1";
  row1.append(spacer1, pinBtn, closeBtn);

  /* row2: tabs + sort + jp search */
  const row2 = document.createElement("div");
  row2.style.cssText = "display:flex;align-items:center;gap:6px;";

  const tabAll = document.createElement("button");
  tabAll.textContent = "ALL";
  const tabFav = document.createElement("button");
  tabFav.textContent = "★";
  const btnSort = document.createElement("button");

  const jpBox = document.createElement("input");
  jpBox.type = "search";
  jpBox.placeholder = "日本語で検索";
  jpBox.autocapitalize = "off";
  jpBox.autocorrect = "off";
  jpBox.spellcheck = false;
  jpBox.style.cssText = `
    flex:1;
    min-width:0;
    font-size:15px;
    padding:8px 10px;
    border-radius:8px;
    border:1px solid #555;
    background:#0c0c0c;
    color:#eee;
    outline:none;
  `;

  row2.append(tabAll, tabFav, btnSort, jpBox);

  /* row3: プロンプト入力欄 (iOS のソフトキーボードでも入力できるように内蔵) */
  const row3 = document.createElement("div");
  row3.style.cssText = "display:flex;align-items:center;gap:6px;";

  const tagBox = document.createElement("input");
  tagBox.type = "search";
  tagBox.placeholder = "英語タグ検索 / ここで入力してタップで挿入";
  tagBox.autocapitalize = "off";
  tagBox.autocorrect = "off";
  tagBox.spellcheck = false;
  tagBox.style.cssText = `
    flex:1;
    min-width:0;
    font-size:15px;
    padding:8px 10px;
    border-radius:8px;
    border:1px solid #555;
    background:#0c0c0c;
    color:#eee;
    outline:none;
  `;
  const hitLabel = document.createElement("span");
  hitLabel.style.cssText = "color:#888;font-size:12px;white-space:nowrap;";
  row3.append(tagBox, hitLabel);

  // 共通ボタンスタイル (タップしやすく大きめ)
  [btnStrong, btnWeak, btnPlus, btnMinus, pinBtn, closeBtn, tabAll, tabFav, btnSort].forEach(b => {
    b.style.cssText = `
      font-size:14px;
      min-width:40px;
      min-height:36px;
      padding:6px 10px;
      border-radius:8px;
      border:1px solid #555;
      background:#222;
      color:#ddd;
      touch-action:manipulation;
      -webkit-user-select:none;
      user-select:none;
    `;
  });

  header.append(row1, row2, row3);

  const list = document.createElement("div");
  list.style.cssText = `
    max-height:min(56vh, 520px);
    overflow-y:auto;
    -webkit-overflow-scrolling:touch;
    font-size:15px;
  `;

  wrapper.append(header, list);
  document.body.appendChild(wrapper);

  // 起動トグルボタン (画面右下に常駐)
  const fab = document.createElement("button");
  fab.textContent = "TAG";
  fab.style.cssText = `
    position:fixed;
    right:16px;
    bottom:16px;
    z-index:2147483646;
    width:56px;
    height:56px;
    border-radius:28px;
    border:1px solid #555;
    background:#1a1a1a;
    color:#ddd;
    font-size:14px;
    font-weight:700;
    box-shadow:0 6px 18px rgba(0,0,0,.5);
    touch-action:manipulation;
  `;
  document.body.appendChild(fab);
  fab.addEventListener("click", () => {
    const isVisible = wrapper.style.display === "block";
    if (isVisible) {
      wrapper.style.display = "none";
    } else {
      showWrapperNearViewport();
      renderSuggest();
    }
  });

  function showWrapperNearViewport() {
    wrapper.style.display = "block";
    if (!wrapper.style.left) {
      const vv = window.visualViewport;
      const vw = vv ? vv.width : window.innerWidth;
      const vh = vv ? vv.height : window.innerHeight;
      const w = wrapper.offsetWidth || 480;
      wrapper.style.left = Math.max(8, (vw - w) / 2) + "px";
      wrapper.style.top = Math.max(8, vh * 0.08) + "px";
    }
  }

  /* =========================================================
   * ▼ buttons
   * ========================================================= */
  function updateButtons() {
    btnStrong.style.background = emphasizeMode === 1 ? "#ff6a6a" : "#222";
    btnWeak.style.background = emphasizeMode === -1 ? "#6b4eff" : "#222";
    pinBtn.style.background = pinEnabled ? "#3f7cff" : "#222";
    tabAll.style.background = activeTab === "all" ? "#3f7cff" : "#222";
    tabFav.style.background = activeTab === "fav" ? "#3f7cff" : "#222";
    tabFav.textContent = FAVORITES.size > 0 ? `★ ${FAVORITES.size}` : "★";
    btnSort.style.background = sortMode === "count" ? "#3f7cff" : "#222";
    btnSort.textContent = sortMode === "prefix" ? "前方優先" : "件数順";
    valueLabel.textContent = emphasizeValue.toFixed(1);
  }

  btnStrong.addEventListener("click", () => { emphasizeMode = emphasizeMode === 1 ? 0 : 1; updateButtons(); });
  btnWeak.addEventListener("click", () => { emphasizeMode = emphasizeMode === -1 ? 0 : -1; updateButtons(); });
  btnPlus.addEventListener("click", () => { emphasizeValue = Math.min(5, emphasizeValue + 0.5); updateButtons(); });
  btnMinus.addEventListener("click", () => { emphasizeValue = Math.max(0.5, emphasizeValue - 0.5); updateButtons(); });
  pinBtn.addEventListener("click", () => { pinEnabled = !pinEnabled; updateButtons(); });
  closeBtn.addEventListener("click", () => { wrapper.style.display = "none"; });

  tabAll.addEventListener("click", () => { activeTab = "all"; updateButtons(); renderSuggest(); });
  tabFav.addEventListener("click", () => { activeTab = "fav"; updateButtons(); renderSuggest(); });
  btnSort.addEventListener("click", () => { sortMode = sortMode === "prefix" ? "count" : "prefix"; updateButtons(); renderSuggest(); });

  jpBox.addEventListener("input", () => {
    jpSearchQuery = jpBox.value.trim();
    wrapper.style.display = "block";
    renderSuggest();
  });
  tagBox.addEventListener("input", () => {
    wrapper.style.display = "block";
    renderSuggest();
  });

  updateButtons();

  /* =========================================================
   * ▼ drag (pointer events)
   * ========================================================= */
  let dragState = null;
  header.addEventListener("pointerdown", e => {
    if (e.target.closest("button,input")) return;
    dragState = {
      id: e.pointerId,
      dx: e.clientX - wrapper.offsetLeft,
      dy: e.clientY - wrapper.offsetTop,
    };
    header.setPointerCapture(e.pointerId);
  });
  header.addEventListener("pointermove", e => {
    if (!dragState || dragState.id !== e.pointerId) return;
    const vv = window.visualViewport;
    const vw = vv ? vv.width : window.innerWidth;
    const vh = vv ? vv.height : window.innerHeight;
    const w = wrapper.offsetWidth, h = wrapper.offsetHeight;
    const x = Math.max(0, Math.min(vw - w, e.clientX - dragState.dx));
    const y = Math.max(0, Math.min(vh - 40, e.clientY - dragState.dy));
    wrapper.style.left = x + "px";
    wrapper.style.top = y + "px";
  });
  const endDrag = e => {
    if (!dragState) return;
    try { header.releasePointerCapture(dragState.id); } catch {}
    dragState = null;
  };
  header.addEventListener("pointerup", endDrag);
  header.addEventListener("pointercancel", endDrag);

  /* =========================================================
   * ▼ 挿入先の捕捉
   * - activeInput: 最後にフォーカスされた textarea/contenteditable
   * - savedRange: contenteditable の場合の Range スナップショット
   * ========================================================= */
  function captureRange() {
    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0) {
      const r = sel.getRangeAt(0);
      // activeInput の内部にある range だけ残す
      if (activeInput && activeInput.contains && activeInput.contains(r.startContainer)) {
        savedRange = r.cloneRange();
      }
    }
  }

  function insertToTarget(outText) {
    if (!activeInput) return false;
    // textarea / input 系
    if (activeInput.tagName === "TEXTAREA" || activeInput.tagName === "INPUT") {
      const el = activeInput;
      const s = el.selectionStart ?? el.value.length;
      const e = el.selectionEnd ?? el.value.length;
      const v = el.value;
      // 現在位置のトークンを置換
      const b = tokenBounds(v, s);
      const newVal = v.slice(0, b.s) + outText + "," + v.slice(b.e);
      el.value = newVal;
      const caret = b.s + outText.length + 1;
      try { el.setSelectionRange(caret, caret); } catch {}
      el.dispatchEvent(new Event("input", { bubbles: true }));
      return true;
    }
    // contenteditable
    if (savedRange) {
      const range = savedRange;
      const node = range.startContainer;
      if (node && node.nodeType === 3) {
        const text = node.textContent;
        const pos = range.startOffset;
        const b = tokenBounds(text, pos);
        const r = document.createRange();
        r.setStart(node, b.s);
        r.setEnd(node, b.e);
        r.deleteContents();
        const insert = document.createTextNode(outText + ",");
        r.insertNode(insert);
        const sel = window.getSelection();
        sel.removeAllRanges();
        const after = document.createRange();
        after.setStart(insert, insert.length);
        after.collapse(true);
        sel.addRange(after);
        savedRange = after.cloneRange();
        // 入力イベントを NovelAI 側に通知
        activeInput.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: outText + "," }));
        return true;
      }
    }
    return false;
  }

  /* =========================================================
   * ▼ suggest 本体
   * ========================================================= */
  function getQueryState() {
    // 内蔵 tagBox に入力があれば最優先
    const tbv = (tagBox.value || "").trim();
    if (tbv) {
      return { source: "box", raw: tbv, query: normalize(tbv) };
    }
    // jp box 優先
    if (jpSearchQuery) {
      return { source: "jp", raw: jpSearchQuery, query: jpSearchQuery, isJp: true };
    }
    // activeInput の現在位置
    if (activeInput) {
      if (activeInput.tagName === "TEXTAREA" || activeInput.tagName === "INPUT") {
        const v = activeInput.value || "";
        const s = activeInput.selectionStart ?? v.length;
        const b = tokenBounds(v, s);
        const raw = stripDecorations(v.slice(b.s, b.e)).trim();
        if (!raw) return { source: "none", raw: "", query: "" };
        const isJp = containsNonAscii(raw);
        return { source: "input", raw, query: isJp ? raw : normalize(raw), isJp };
      }
      if (savedRange) {
        const node = savedRange.startContainer;
        if (node && node.nodeType === 3) {
          const text = node.textContent;
          const pos = savedRange.startOffset;
          const b = tokenBounds(text, pos);
          const raw = stripDecorations(text.slice(b.s, b.e)).trim();
          if (raw) {
            const isJp = containsNonAscii(raw);
            return { source: "ce", raw, query: isJp ? raw : normalize(raw), isJp };
          }
        }
      }
    }
    return { source: "none", raw: "", query: "" };
  }

  function renderSuggest() {
    if (!TAGS.length) return;
    const qs = getQueryState();
    const query = qs.query;
    const isJp = !!qs.isJp;
    const queryCompact = !isJp && query ? compactForSearch(query) : "";

    const canShowWithoutQuery = activeTab === "fav";
    if (!query && !canShowWithoutQuery) {
      hitLabel.textContent = "";
      list.innerHTML = "";
      return;
    }

    const directStarts = [];
    const directIncludes = [];
    const aliasStartsByTag = new Map();
    const aliasIncludesByTag = new Map();

    if (query) {
      for (const t of TAGS) {
        if (!isJp) {
          const startsByKey = t.key.startsWith(query);
          const includesByKey = t.key.includes(query);
          const startsByCompact = !!queryCompact && (t.compactKey || "").startsWith(queryCompact);
          const includesByCompact = !!queryCompact && (t.compactKey || "").includes(queryCompact);
          if (startsByKey || startsByCompact) directStarts.push({ t, alias: "" });
          else if (includesByKey || includesByCompact) directIncludes.push({ t, alias: "" });
        } else {
          if (!t.jp) continue;
          if (t.jp.startsWith(query)) directStarts.push({ t, alias: "" });
          else if (t.jp.includes(query)) directIncludes.push({ t, alias: "" });
        }
      }
      if (!isJp) {
        for (const a of TAG_ALIAS_ENTRIES) {
          const startsByAliasKey = a.aliasKey.startsWith(query);
          const includesByAliasKey = a.aliasKey.includes(query);
          const startsByAliasCompact = !!queryCompact && (a.aliasCompact || "").startsWith(queryCompact);
          const includesByAliasCompact = !!queryCompact && (a.aliasCompact || "").includes(queryCompact);
          const isStarts = startsByAliasKey || startsByAliasCompact;
          const isIncludes = !isStarts && (includesByAliasKey || includesByAliasCompact);
          if (!isStarts && !isIncludes) continue;
          const t = TAGS_BY_KEY.get(a.tagKey);
          if (!t) continue;
          if (isStarts) {
            const prev = aliasStartsByTag.get(a.tagKey);
            if (!prev || betterAliasChoice(a.alias, prev.alias, query)) {
              aliasStartsByTag.set(a.tagKey, { t, alias: a.alias });
            }
            aliasIncludesByTag.delete(a.tagKey);
          } else if (!aliasStartsByTag.has(a.tagKey)) {
            const prev = aliasIncludesByTag.get(a.tagKey);
            if (!prev || betterAliasChoice(a.alias, prev.alias, query)) {
              aliasIncludesByTag.set(a.tagKey, { t, alias: a.alias });
            }
          }
        }
      }
    }

    directStarts.sort((a, b) => b.t.count - a.t.count);
    directIncludes.sort((a, b) => b.t.count - a.t.count);
    const aliasStarts = [...aliasStartsByTag.values()].sort((a, b) => b.t.count - a.t.count);
    const aliasIncludes = [...aliasIncludesByTag.values()].sort((a, b) => b.t.count - a.t.count);

    const orderedHits = sortMode === "prefix"
      ? [...directStarts, ...directIncludes, ...aliasStarts, ...aliasIncludes]
      : [...directStarts, ...directIncludes, ...aliasStarts, ...aliasIncludes].sort((a, b) => b.t.count - a.t.count);

    let hits = [];
    const usedHits = new Set();
    for (const hit of orderedHits) {
      if (!hit?.t?.key) continue;
      if (usedHits.has(hit.t.key)) continue;
      usedHits.add(hit.t.key);
      hits.push(hit);
      if (hits.length >= MAX_RESULTS) break;
    }

    const matchFn = t => {
      if (!query) return true;
      if (!isJp) {
        if (t.key.includes(query)) return true;
        if (!!queryCompact && (t.compactKey || "").includes(queryCompact)) return true;
        const aliases = TAG_ALIASES_BY_TAG_KEY.get(t.key) || [];
        for (const a of aliases) {
          if ((a.aliasKey || "").includes(query)) return true;
          if (!!queryCompact && (a.aliasCompact || "").includes(queryCompact)) return true;
        }
        return false;
      }
      return t.jp && t.jp.includes(query);
    };

    const favList = [...FAVORITES].map(k => TAGS_BY_KEY.get(k)).filter(Boolean).filter(matchFn);
    if (sortMode === "count") favList.sort((a, b) => b.count - a.count);
    const histList = HISTORY.map(k => TAGS_BY_KEY.get(k)).filter(Boolean).filter(matchFn);
    if (sortMode === "count") histList.sort((a, b) => b.count - a.count);

    const used = new Set();
    let merged = [];
    if (activeTab === "fav") {
      merged = favList.slice(0, MAX_RESULTS).map(t => ({ t, alias: "" }));
    } else {
      for (const t of favList) { if (used.has(t.key)) continue; used.add(t.key); merged.push({ t, alias: "" }); }
      for (const t of histList) { if (used.has(t.key)) continue; used.add(t.key); merged.push({ t, alias: "" }); }
      for (const hit of hits) {
        const t = hit.t;
        if (used.has(t.key)) continue;
        used.add(t.key);
        merged.push(hit);
        if (merged.length >= MAX_RESULTS) break;
      }
    }
    TAGS_HITS_CACHE = merged;
    hitLabel.textContent = merged.length >= MAX_RESULTS ? "500+" : `${merged.length}`;

    list.innerHTML = "";
    merged.forEach(hit => appendRow(hit, query, isJp));

    clearTimeout(hideTimer);
    if (!pinEnabled) {
      hideTimer = setTimeout(() => { wrapper.style.display = "none"; }, AUTO_HIDE_DELAY_MS);
    }
  }

  function appendRow(hit, query, isJp) {
    const t = hit.t;
    const row = document.createElement("div");
    row.style.cssText = `
      padding:12px 10px;
      display:flex;
      gap:10px;
      align-items:center;
      color:${colorByType(t.typeId)};
      border-bottom:1px solid #1e1e1e;
      touch-action:manipulation;
    `;

    const star = document.createElement("button");
    star.textContent = FAVORITES.has(t.key) ? "★" : "☆";
    star.style.cssText = `
      font-size:18px;
      min-width:40px;
      min-height:40px;
      padding:0;
      border-radius:8px;
      border:1px solid #333;
      background:#181818;
      color:#ffcc66;
      touch-action:manipulation;
    `;

    const textBox = document.createElement("div");
    textBox.style.cssText = "flex:1;min-width:0;";

    const main = document.createElement("div");
    main.style.cssText = "white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-size:15px;";
    const disp = displayTag(t.tag);
    if (!isJp && hit.alias) {
      main.innerHTML =
        highlightMatch(hit.alias, query) +
        ` <span style="color:#777;">→</span> ` +
        `<span style="font-weight:700;">${t.tag}</span>`;
    } else {
      main.innerHTML = !isJp ? highlightMatch(disp, query) : disp;
    }

    const sub = document.createElement("div");
    sub.style.cssText = "color:#aaa;font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;";
    sub.innerHTML = t.jp ? (isJp ? highlightMatch(t.jp, query) : t.jp) : "";
    if (Number(t.count) > 0) {
      sub.innerHTML += ` <span style="color:#666;">(${t.count})</span>`;
    }

    textBox.append(main, sub);

    const wikiBtn = document.createElement("button");
    wikiBtn.textContent = "?";
    wikiBtn.style.cssText = `
      font-size:14px;
      min-width:40px;
      min-height:40px;
      padding:0;
      border-radius:8px;
      border:1px solid #333;
      background:#181818;
      color:#bbb;
      touch-action:manipulation;
    `;

    row.append(star, textBox, wikiBtn);

    // iOS 重要: pointerdown を preventDefault してフォーカス奪取を防ぐ
    const preventFocusSteal = e => { e.preventDefault(); };
    row.addEventListener("pointerdown", preventFocusSteal);
    star.addEventListener("pointerdown", preventFocusSteal);
    wikiBtn.addEventListener("pointerdown", preventFocusSteal);

    star.addEventListener("click", e => {
      e.preventDefault();
      e.stopPropagation();
      if (FAVORITES.has(t.key)) FAVORITES.delete(t.key);
      else FAVORITES.add(t.key);
      saveFavorites();
      updateButtons();
      renderSuggest();
    });

    wikiBtn.addEventListener("click", e => {
      e.preventDefault();
      e.stopPropagation();
      window.open(wikiUrlForTag(t.tag), "_blank", "noopener,noreferrer");
    });

    row.addEventListener("click", e => {
      if (e.target === star || e.target === wikiBtn) return;
      e.preventDefault();
      let out = formatInsertTag(t.tag);
      if (emphasizeMode === 1) out = `${emphasizeValue}::${out}::`;
      if (emphasizeMode === -1) out = `-${emphasizeValue}::${out}::`;

      // tagBox 経由の場合: tagBox をクリアしてから activeInput に挿入
      const usedTagBox = (tagBox.value || "").trim().length > 0;
      if (insertToTarget(out)) {
        addToHistory(t.key);
        if (usedTagBox) tagBox.value = "";
        if (!pinEnabled) wrapper.style.display = "none";
      } else {
        // 挿入先がないときは tagBox に書き戻す (クリップボード用途)
        tagBox.value = out;
      }
    });

    list.appendChild(row);
  }

  /* =========================================================
   * ▼ activeInput の捕捉 + selection スナップショット
   * ========================================================= */
  function attach(el) {
    if (el.__naiIos) return;
    el.__naiIos = true;
    el.addEventListener("focus", () => { activeInput = el; });
    el.addEventListener("blur", () => { /* keep activeInput */ });
    el.addEventListener("input", () => {
      if (el.tagName === "TEXTAREA" || el.tagName === "INPUT") {
        renderSuggest();
      } else {
        captureRange();
        renderSuggest();
      }
    });
    el.addEventListener("keyup", () => { captureRange(); });
    el.addEventListener("mouseup", () => { captureRange(); });
    el.addEventListener("touchend", () => { setTimeout(captureRange, 0); });
  }

  // contenteditable 内での選択変更を広く拾う
  document.addEventListener("selectionchange", () => {
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;
    const anchor = sel.anchorNode;
    if (!anchor) return;
    // contenteditable の先祖を探して activeInput にする
    let el = anchor.nodeType === 1 ? anchor : anchor.parentNode;
    while (el && el !== document.body) {
      if (el.isContentEditable) {
        if (!el.__naiIos) attach(el);
        activeInput = el;
        savedRange = sel.getRangeAt(0).cloneRange();
        break;
      }
      el = el.parentNode;
    }
  });

  function scan() {
    document.querySelectorAll("textarea,[contenteditable='true']").forEach(attach);
  }

  /* =========================================================
   * ▼ visualViewport 追従 (ソフトキーボード対応)
   * ========================================================= */
  if (window.visualViewport) {
    window.visualViewport.addEventListener("resize", () => {
      if (wrapper.style.display !== "block") return;
      const vv = window.visualViewport;
      const bottomOverflow =
        (parseInt(wrapper.style.top || "0", 10) + wrapper.offsetHeight) - vv.height;
      if (bottomOverflow > 0) {
        const newTop = Math.max(8, parseInt(wrapper.style.top || "0", 10) - bottomOverflow - 8);
        wrapper.style.top = newTop + "px";
      }
    });
  }

  /* =========================================================
   * ▼ main
   * ========================================================= */
  async function main() {
    loadUserLists();
    updateButtons();
    // ローディング表示
    list.innerHTML = `<div style="padding:16px;color:#aaa;">CSV 読込中... (初回のみ通信、2回目以降はキャッシュ)</div>`;
    wrapper.style.display = "block";
    showWrapperNearViewport();
    try {
      await loadJPMaps();
      await loadAllCSV();
      list.innerHTML = `<div style="padding:16px;color:#7f7;">読込完了: ${TAGS.length} タグ</div>`;
    } catch (e) {
      list.innerHTML = `<div style="padding:16px;color:#f77;">読込失敗: ${String(e && e.message || e)}<br>CSV_BASE_URL を確認してください</div>`;
    }
    scan();
    new MutationObserver(scan).observe(document.body, { childList: true, subtree: true });
  }

  main();
})();
