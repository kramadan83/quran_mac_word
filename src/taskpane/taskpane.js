/*
 * Quran in Word - Office Add-in
 * Inserts Quranic verses with translations into Word documents.
 * Supports mushaf-style rendering with continuous Arabic text and verse markers.
 */

/* global document, Office, Word */

import surahList from "../data/surahList.json";
import pageToAyahs from "../data/pageToAyahs.json";
import { getAllLanguages, getLanguageById, getDefaultLanguageIds } from "./translationRegistry";
import { loadTranslation } from "./translationLoader";

// Data cache
const dataCache = {
  arabic: {},
  translations: {}, // { [langId]: { [surahNum]: { ayahs: [...] } } }
};

// Active translation languages (managed by UI)
let activeLanguages = [];

// Load Arabic data for a surah
async function loadArabicData(surahNumber) {
  if (dataCache.arabic[surahNumber]) return;
  const mod = await import(
    /* webpackChunkName: "arabic-[request]" */ `../data/arabic/${surahNumber}.json`
  );
  dataCache.arabic[surahNumber] = mod.default || mod;
}

// Load all active translations for a surah
async function loadSurahData(surahNumber) {
  const loads = [loadArabicData(surahNumber)];
  for (const langId of activeLanguages) {
    if (!dataCache.translations[langId]) {
      dataCache.translations[langId] = {};
    }
    if (!dataCache.translations[langId][surahNumber]) {
      loads.push(
        loadTranslation(langId, surahNumber)
          .then((data) => {
            dataCache.translations[langId][surahNumber] = data;
          })
          .catch(() => {
            // Graceful degradation: skip this language if fetch fails
            dataCache.translations[langId][surahNumber] = { ayahs: [] };
          })
      );
    }
  }
  await Promise.all(loads);
}

// --- Helpers ---

const ARABIC_INDIC_ZERO = 0x0660; // ٠

// Waqf marks (U+06D6-U+06DC) are Unicode combining marks.
// In Word, they show as dotted circles when standalone or are invisible when attached.
// The font has glyphs but Word's text engine doesn't render them well as combining marks.
// Solution: strip them from the text entirely for clean rendering in Word.
// Waqf marks are an editorial feature of printed mushafs, not part of the Quranic text.
const MUSHAF_MARKS_RE = /[\u06D6-\u06DC\u06DE\u06DF\u06E0\u06E9]/g;

function cleanArabicText(text) {
  return text.replace(MUSHAF_MARKS_RE, "").replace(/  +/g, " ").trim();
}

function toArabicIndic(num) {
  return String(num)
    .split("")
    .map((d) => String.fromCharCode(ARABIC_INDIC_ZERO + parseInt(d, 10)))
    .join("");
}

function buildVerseMarker(ayahNumber) {
  // Arabic-Indic digits in KFGQPC HAFS font render inside ornamental circles.
  // U+06DD is NOT used because Word on Mac renders it as a separate blank circle
  // alongside the digit's ornamental circle, resulting in a duplicate marker.
  return " " + toArabicIndic(ayahNumber) + " ";
}

// --- Init ---

// Register service worker for offline support
if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("service-worker.js").catch(() => {});
  });
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    initUI();
  }
});

function initUI() {
  initSurahSearch();
  initLanguageSelector();
  updateAyahLimits();

  // Mode toggle
  document.querySelectorAll('input[name="insert-mode"]').forEach((radio) => {
    radio.addEventListener("change", toggleMode);
  });

  // Arabic only toggle
  document.getElementById("chk-arabic-only").addEventListener("change", () => {
    toggleArabicOnly();
    saveArabicOnlyPreference();
  });
  restoreArabicOnlyPreference();

  // Single ayah input - clamp on blur so user can freely type
  document.getElementById("ayah-single").addEventListener("change", () => {
    clampSingleAyah();
  });

  // Range inputs - clamp on blur so user can freely type
  document.getElementById("ayah-from").addEventListener("change", () => {
    clampRangeInputs();
  });
  document.getElementById("ayah-to").addEventListener("change", () => {
    clampRangeInputs();
  });

  // Page input
  document.getElementById("mushaf-page").addEventListener("change", () => {
    clampPageInput();
    updatePageInfo();
  });

  document.getElementById("btn-insert").addEventListener("click", insertToWord);
}

function getInsertMode() {
  return document.querySelector('input[name="insert-mode"]:checked').value;
}

function getRangeLayout() {
  return document.querySelector('input[name="range-layout"]:checked').value;
}

function toggleMode() {
  const mode = getInsertMode();
  document.getElementById("single-mode").style.display = mode === "single" ? "" : "none";
  document.getElementById("range-mode").style.display = mode === "range" ? "" : "none";
  document.getElementById("page-mode").style.display = mode === "page" ? "" : "none";

  const surahWrap = document.getElementById("surah-selector-wrap");
  if (surahWrap) {
    surahWrap.style.display = mode === "page" ? "none" : "";
  }

  // Auto-enable Arabic-only for page mode
  if (mode === "page") {
    document.getElementById("chk-arabic-only").checked = true;
    toggleArabicOnly();
    updatePageInfo();
  } else {
    const arabicOnly = document.getElementById("chk-arabic-only").checked;
    if (!arabicOnly) {
      showTranslationsSection();
    }
  }
}

// --- Language Selector ---

const STORAGE_KEY = "quran-word-active-languages";
const ENABLED_KEY = "quran-word-enabled-languages";
const MAX_LANGUAGES = 3;

function loadLanguagePreferences() {
  try {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (stored) {
      const parsed = JSON.parse(stored);
      const valid = parsed.filter((id) => getLanguageById(id));
      if (valid.length >= 1 && valid.length <= MAX_LANGUAGES) {
        return valid;
      }
    }
  } catch (_) {
    // ignore parse errors
  }
  return getDefaultLanguageIds();
}

function loadEnabledLanguages() {
  try {
    const stored = localStorage.getItem(ENABLED_KEY);
    if (stored) {
      const parsed = JSON.parse(stored);
      return parsed.filter((id) => getLanguageById(id));
    }
  } catch (_) {
    // ignore parse errors
  }
  return getDefaultLanguageIds();
}

function saveLanguagePreferences() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(activeLanguages));
}

function saveEnabledLanguages() {
  const enabled = getEnabledLanguages();
  localStorage.setItem(ENABLED_KEY, JSON.stringify(enabled));
}

function initLanguageSelector() {
  activeLanguages = loadLanguagePreferences();
  const savedEnabled = loadEnabledLanguages();
  renderLanguageChips(savedEnabled);

  const addBtn = document.getElementById("btn-add-language");
  const dropdown = document.getElementById("language-dropdown");

  addBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    if (activeLanguages.length >= MAX_LANGUAGES) return;
    if (dropdown.classList.contains("open")) {
      closeLangDropdown();
    } else {
      renderLangDropdown();
      dropdown.classList.add("open");
    }
  });

  document.addEventListener("click", (e) => {
    if (!e.target.closest(".language-add-wrap")) {
      closeLangDropdown();
    }
  });
}

function renderLanguageChips(enabledIds) {
  const container = document.getElementById("active-languages");

  container.innerHTML = activeLanguages
    .map((id) => {
      const lang = getLanguageById(id);
      if (!lang) return "";
      const checked = enabledIds ? enabledIds.includes(id) : true;
      return (
        `<label class="language-chip" data-lang="${id}">` +
        `<input type="checkbox" class="language-chip__check" data-lang="${id}"${checked ? " checked" : ""} />` +
        `<span class="language-chip__name">${lang.name}</span>` +
        `<button class="language-chip__remove" data-lang="${id}" title="Remove ${lang.name}">\u00d7</button>` +
        `</label>`
      );
    })
    .join("");

  container.querySelectorAll(".language-chip__remove").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      removeLanguage(btn.dataset.lang);
    });
  });

  container.querySelectorAll(".language-chip__check").forEach((chk) => {
    chk.addEventListener("change", () => {
      saveEnabledLanguages();
    });
  });

  // Update add button state
  const addBtn = document.getElementById("btn-add-language");
  if (activeLanguages.length >= MAX_LANGUAGES) {
    addBtn.classList.add("language-add-btn--disabled");
    addBtn.title = "Maximum 3 translations";
  } else {
    addBtn.classList.remove("language-add-btn--disabled");
    addBtn.title = "";
  }
}

function renderLangDropdown() {
  const dropdown = document.getElementById("language-dropdown");
  const all = getAllLanguages();
  const available = all.filter((l) => !activeLanguages.includes(l.id));

  dropdown.innerHTML = available
    .map(
      (l) =>
        `<div class="language-dropdown-item" data-lang="${l.id}">` +
        `<span class="language-dropdown-item__name">${l.name}</span>` +
        `<span class="language-dropdown-item__native">${l.nativeName}</span>` +
        `</div>`
    )
    .join("");

  dropdown.querySelectorAll(".language-dropdown-item").forEach((el) => {
    el.addEventListener("mousedown", (e) => {
      e.preventDefault();
      addLanguage(el.dataset.lang);
      closeLangDropdown();
    });
  });
}

function closeLangDropdown() {
  document.getElementById("language-dropdown").classList.remove("open");
}

function addLanguage(id) {
  if (activeLanguages.length >= MAX_LANGUAGES) return;
  if (activeLanguages.includes(id)) return;
  activeLanguages.push(id);
  saveLanguagePreferences();
  // New language is checked by default
  const enabled = getEnabledLanguages();
  enabled.push(id);
  renderLanguageChips(enabled);
  saveEnabledLanguages();
}

function removeLanguage(id) {
  activeLanguages = activeLanguages.filter((l) => l !== id);
  if (activeLanguages.length === 0) {
    activeLanguages = getDefaultLanguageIds().slice(0, 1);
  }
  saveLanguagePreferences();
  renderLanguageChips();
  saveEnabledLanguages();
}

function getEnabledLanguages() {
  const checks = document.querySelectorAll(".language-chip__check:checked");
  return Array.from(checks).map((c) => c.dataset.lang);
}

function getActiveLanguages() {
  // Return only the checked (enabled) languages
  return getEnabledLanguages();
}

// --- Arabic Only Mode ---

const ARABIC_ONLY_KEY = "quran-word-arabic-only";

function toggleArabicOnly() {
  const checked = document.getElementById("chk-arabic-only").checked;
  if (checked) {
    hideTranslationsSection();
    activeLanguages = [];
  } else {
    activeLanguages = loadLanguagePreferences();
    showTranslationsSection();
    const savedEnabled = loadEnabledLanguages();
    renderLanguageChips(savedEnabled);
  }
}

function hideTranslationsSection() {
  const container = document.getElementById("active-languages");
  const addWrap = document.getElementById("btn-add-language").closest(".language-add-wrap");
  // Find parent form-group for each
  let containerGroup = container.parentElement;
  while (containerGroup && !containerGroup.classList.contains("form-group")) {
    containerGroup = containerGroup.parentElement;
  }
  let addWrapGroup = addWrap.parentElement;
  while (addWrapGroup && !addWrapGroup.classList.contains("form-group")) {
    addWrapGroup = addWrapGroup.parentElement;
  }
  if (containerGroup) containerGroup.style.display = "none";
  if (addWrapGroup) addWrapGroup.style.display = "none";
}

function showTranslationsSection() {
  const container = document.getElementById("active-languages");
  const addWrap = document.getElementById("btn-add-language").closest(".language-add-wrap");
  let containerGroup = container.parentElement;
  while (containerGroup && !containerGroup.classList.contains("form-group")) {
    containerGroup = containerGroup.parentElement;
  }
  let addWrapGroup = addWrap.parentElement;
  while (addWrapGroup && !addWrapGroup.classList.contains("form-group")) {
    addWrapGroup = addWrapGroup.parentElement;
  }
  if (containerGroup) containerGroup.style.display = "";
  if (addWrapGroup) addWrapGroup.style.display = "";
}

function saveArabicOnlyPreference() {
  localStorage.setItem(
    ARABIC_ONLY_KEY,
    document.getElementById("chk-arabic-only").checked ? "1" : "0"
  );
}

function restoreArabicOnlyPreference() {
  const stored = localStorage.getItem(ARABIC_ONLY_KEY);
  if (stored === "1") {
    document.getElementById("chk-arabic-only").checked = true;
    toggleArabicOnly();
  }
}

// --- Mushaf Page Mode ---

function clampPageInput() {
  const el = document.getElementById("mushaf-page");
  let val = parseInt(el.value, 10);
  if (isNaN(val) || val < 1) val = 1;
  if (val > 604) val = 604;
  el.value = val;
}

function getPageInfo(pageNumber) {
  const ayahs = pageToAyahs[String(pageNumber)];
  if (!ayahs || ayahs.length === 0) return null;

  const first = ayahs[0];
  const last = ayahs[ayahs.length - 1];
  const firstInfo = getSurahInfo(first.surah);
  const lastInfo = getSurahInfo(last.surah);

  let info;
  if (first.surah === last.surah) {
    info = `${firstInfo.name} (${first.surah}:${first.ayah}-${last.ayah})`;
  } else {
    info = `${firstInfo.name} (${first.surah}:${first.ayah}) - ${lastInfo.name} (${last.surah}:${last.ayah})`;
  }

  return {
    surahs: [...new Set(ayahs.map((a) => a.surah))],
    ayahs,
    info,
    count: ayahs.length,
  };
}

function updatePageInfo() {
  const pageNum = parseInt(document.getElementById("mushaf-page").value, 10);
  if (isNaN(pageNum) || pageNum < 1 || pageNum > 604) return;

  const info = getPageInfo(pageNum);
  const el = document.getElementById("page-info");
  if (info) {
    el.textContent = info.info;
  } else {
    el.textContent = "";
  }
}

async function loadPageData(pageNumber) {
  const pageEntry = pageToAyahs[String(pageNumber)];
  if (!pageEntry) return [];

  const uniqueSurahs = [...new Set(pageEntry.map((a) => a.surah))];
  await Promise.all(uniqueSurahs.map((s) => loadArabicData(s)));

  const results = [];
  for (const entry of pageEntry) {
    const surahData = dataCache.arabic[entry.surah];
    if (surahData && surahData.ayahs) {
      const ayah = surahData.ayahs.find((a) => a.number === entry.ayah);
      if (ayah) {
        results.push({
          surah: entry.surah,
          ayah: entry.ayah,
          arabic: cleanArabicText(ayah.text),
        });
      }
    }
  }
  return results;
}

// --- Surah search / Ayah helpers ---

let activeItemIndex = -1;

function initSurahSearch() {
  const input = document.getElementById("surah-search");
  const dropdown = document.getElementById("surah-dropdown");
  const hiddenVal = document.getElementById("surah-value");

  // Default to surah 1 (field left empty to show placeholder)
  const first = surahList[0];
  hiddenVal.value = first.number;

  // Build all items once
  renderDropdown(surahList, dropdown);

  input.addEventListener("focus", () => {
    input.select();
    filterAndShow();
  });

  input.addEventListener("input", () => {
    filterAndShow();
  });

  input.addEventListener("keydown", (e) => {
    const items = dropdown.querySelectorAll(".surah-item");
    if (e.key === "ArrowDown") {
      e.preventDefault();
      activeItemIndex = Math.min(activeItemIndex + 1, items.length - 1);
      highlightItem(items);
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      activeItemIndex = Math.max(activeItemIndex - 1, 0);
      highlightItem(items);
    } else if (e.key === "Enter") {
      e.preventDefault();
      if (activeItemIndex >= 0 && items[activeItemIndex]) {
        selectSurah(parseInt(items[activeItemIndex].dataset.number, 10));
      }
      closeDropdown();
    } else if (e.key === "Escape") {
      closeDropdown();
      input.blur();
    }
  });

  // Close dropdown on outside click
  document.addEventListener("click", (e) => {
    if (!e.target.closest(".surah-search-wrap")) {
      closeDropdown();
    }
  });
}

function filterAndShow() {
  const input = document.getElementById("surah-search");
  const dropdown = document.getElementById("surah-dropdown");
  const query = input.value.toLowerCase().trim();

  const filtered = query
    ? surahList.filter(
        (s) =>
          String(s.number).startsWith(query) ||
          s.name.toLowerCase().includes(query) ||
          s.arabic.includes(query)
      )
    : surahList;

  renderDropdown(filtered, dropdown);
  activeItemIndex = -1;
  dropdown.classList.add("open");
}

function renderDropdown(items, dropdown) {
  const selected = getSelectedSurah();
  dropdown.innerHTML = items
    .map(
      (s) =>
        `<div class="surah-item${s.number === selected ? " selected" : ""}" data-number="${s.number}">` +
        `<span class="surah-item__name">${s.number}. ${s.name}</span>` +
        `<span class="surah-item__arabic">${s.arabic}</span>` +
        `</div>`
    )
    .join("");

  dropdown.querySelectorAll(".surah-item").forEach((el) => {
    el.addEventListener("mousedown", (e) => {
      e.preventDefault();
      selectSurah(parseInt(el.dataset.number, 10));
      closeDropdown();
    });
  });
}

function selectSurah(number) {
  const input = document.getElementById("surah-search");
  const hiddenVal = document.getElementById("surah-value");
  const info = getSurahInfo(number);
  if (!info) return;
  hiddenVal.value = number;
  input.value = `${info.number}. ${info.name} (${info.arabic})`;
  updateAyahLimits();
  resetAyahInputs();
}

function closeDropdown() {
  document.getElementById("surah-dropdown").classList.remove("open");
  activeItemIndex = -1;
}

function highlightItem(items) {
  items.forEach((el, i) => {
    el.classList.toggle("active", i === activeItemIndex);
    if (i === activeItemIndex) el.scrollIntoView({ block: "nearest" });
  });
}

function getSelectedSurah() {
  return parseInt(document.getElementById("surah-value").value, 10);
}

function getSurahInfo(surahNumber) {
  return surahList.find((s) => s.number === surahNumber);
}

function getSelectedAyahRange() {
  if (getInsertMode() === "single") {
    let num = parseInt(document.getElementById("ayah-single").value, 10);
    if (isNaN(num) || num < 1) num = 1;
    return { from: num, to: num };
  }
  let from = parseInt(document.getElementById("ayah-from").value, 10);
  let to = parseInt(document.getElementById("ayah-to").value, 10);
  if (isNaN(from) || from < 1) from = 1;
  if (isNaN(to) || to < 1) to = 1;
  if (from > to) {
    const tmp = from;
    from = to;
    to = tmp;
  }
  return { from, to };
}

function updateAyahLimits() {
  const info = getSurahInfo(getSelectedSurah());
  if (!info) return;
  const singleEl = document.getElementById("ayah-single");
  singleEl.max = info.total_ayah;
  document.getElementById("ayah-total-single").textContent = `/ ${info.total_ayah}`;
  const fromEl = document.getElementById("ayah-from");
  const toEl = document.getElementById("ayah-to");
  fromEl.max = info.total_ayah;
  toEl.max = info.total_ayah;
  document.getElementById("ayah-range").textContent = `/ ${info.total_ayah}`;
}

function resetAyahInputs() {
  const info = getSurahInfo(getSelectedSurah());
  if (!info) return;
  document.getElementById("ayah-single").value = 1;
  document.getElementById("ayah-from").value = 1;
  document.getElementById("ayah-to").value = info.total_ayah;
}

function clampSingleAyah() {
  const info = getSurahInfo(getSelectedSurah());
  if (!info) return;
  const el = document.getElementById("ayah-single");
  let val = parseInt(el.value, 10);
  if (isNaN(val) || val < 1) val = 1;
  if (val > info.total_ayah) val = info.total_ayah;
  el.value = val;
}

function clampRangeInputs() {
  const info = getSurahInfo(getSelectedSurah());
  if (!info) return;
  const fromEl = document.getElementById("ayah-from");
  const toEl = document.getElementById("ayah-to");

  let from = parseInt(fromEl.value, 10);
  let to = parseInt(toEl.value, 10);

  if (isNaN(from) || from < 1) from = 1;
  if (from > info.total_ayah) from = info.total_ayah;

  if (isNaN(to) || to < 1) to = 1;
  if (to > info.total_ayah) to = info.total_ayah;

  if (to < from) to = from;

  fromEl.value = from;
  toEl.value = to;
}

// --- Data access ---

function getAyahData(surahNumber, ayahNumber) {
  const arabic = dataCache.arabic[surahNumber];
  if (!arabic) return null;

  const arabicAyah = arabic.ayahs.find((a) => a.number === ayahNumber);
  if (!arabicAyah) return null;

  // Build translations map for active languages
  const translations = {};
  for (const langId of activeLanguages) {
    const langData = dataCache.translations[langId] && dataCache.translations[langId][surahNumber];
    if (langData && langData.ayahs) {
      const ayah = langData.ayahs.find((a) => a.number === ayahNumber);
      if (ayah) {
        translations[langId] = ayah.text;
      }
    }
  }

  return {
    number: ayahNumber,
    arabic: cleanArabicText(arabicAyah.text),
    translations,
  };
}

function getAyahRangeData(surahNumber, fromAyah, toAyah) {
  const results = [];
  for (let i = fromAyah; i <= toAyah; i++) {
    const data = getAyahData(surahNumber, i);
    if (data) results.push(data);
  }
  return results;
}

// --- Word insertion ---

function setStatus(message, isError) {
  const el = document.getElementById("status");
  el.textContent = message;
  el.className = "ms-font-s status " + (isError ? "status--error" : "status--success");
  if (message) {
    setTimeout(() => {
      el.textContent = "";
      el.className = "ms-font-s status";
    }, 3000);
  }
}

function buildTranslationLines(surahNum, fromAyah, toAyah, langIds) {
  const ayahs = getAyahRangeData(surahNum, fromAyah, toAyah);
  const info = getSurahInfo(surahNum);
  const surahName = info ? info.name : `Surah ${surahNum}`;
  const lines = [];

  for (const langId of langIds) {
    const lang = getLanguageById(langId);
    if (!lang) continue;

    // Check if we have any translation data for this language
    const hasData = ayahs.some((a) => a.translations[langId]);
    if (!hasData) continue;

    ayahs.forEach((a) => {
      const text = a.translations[langId];
      if (text) {
        lines.push({ text: `${a.number}. ${text}`, langId });
      }
    });
  }

  const rangeStr = fromAyah === toAyah ? `${fromAyah}` : `${fromAyah}-${toAyah}`;
  lines.push({ text: `(QS. ${surahName}: ${rangeStr})`, isReference: true });

  return lines;
}

export async function insertToWord() {
  const mode = getInsertMode();

  // Page mode
  if (mode === "page") {
    await insertPageToWord();
    return;
  }

  // Single or Range mode
  const surahNum = getSelectedSurah();
  const { from, to } = getSelectedAyahRange();
  const langs = getActiveLanguages();
  const isSingleMode = mode === "single";
  const showAyahNumber = isSingleMode
    ? document.getElementById("chk-show-ayah-number").checked
    : document.getElementById("chk-show-ayah-number-range").checked;

  // Load surah data if not yet cached
  setStatus("Loading data...", false);
  try {
    await loadSurahData(surahNum);
  } catch (err) {
    setStatus("Failed to load surah data: " + err.message, true);
    return;
  }

  const ayahs = getAyahRangeData(surahNum, from, to);
  if (ayahs.length === 0) {
    setStatus("No data available for this ayah range.", true);
    return;
  }

  const translationLines = buildTranslationLines(surahNum, from, to, langs);

  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const isPerLine = !isSingleMode && getRangeLayout() === "per-line";

      if (isPerLine) {
        // Per-line layout: each ayah gets its own paragraph
        for (let i = 0; i < ayahs.length; i++) {
          const a = ayahs[i];
          const para = body.insertParagraph("", Word.InsertLocation.end);
          para.font.name = "KFGQPC HAFS Uthmanic Script";
          para.font.size = 18;
          para.font.color = "#000000";
          para.alignment = Word.Alignment.right;
          para.lineSpacing = 16;
          para.spaceAfter = 0;
          para.spaceBefore = 0;
          para.rightIndent = 0;
          para.leftIndent = 0;
          para.firstLineIndent = 0;

          await context.sync();

          const textRange = para.getRange(Word.RangeLocation.end);
          const textRun = textRange.insertText(a.arabic, Word.InsertLocation.end);
          textRun.font.name = "KFGQPC HAFS Uthmanic Script";
          textRun.font.size = 18;
          textRun.font.color = "#000000";

          if (showAyahNumber) {
            const markerRange = para.getRange(Word.RangeLocation.end);
            const markerRun = markerRange.insertText(
              buildVerseMarker(a.number),
              Word.InsertLocation.end
            );
            markerRun.font.name = "KFGQPC HAFS Uthmanic Script";
            markerRun.font.size = 20;
            markerRun.font.color = "#000000";
          }
        }
      } else {
        // Continuous layout: all ayahs in one paragraph
        const arabicPara = body.insertParagraph("", Word.InsertLocation.end);
        arabicPara.font.name = "KFGQPC HAFS Uthmanic Script";
        arabicPara.font.size = 18;
        arabicPara.font.color = "#000000";
        arabicPara.alignment = Word.Alignment.right;
        arabicPara.lineSpacing = 16;
        arabicPara.spaceAfter = 0;
        arabicPara.spaceBefore = 0;
        arabicPara.rightIndent = 0;
        arabicPara.leftIndent = 0;
        arabicPara.firstLineIndent = 0;

        await context.sync();

        for (let i = 0; i < ayahs.length; i++) {
          const a = ayahs[i];

          if (!showAyahNumber && i > 0) {
            const spaceRange = arabicPara.getRange(Word.RangeLocation.end);
            const spaceRun = spaceRange.insertText("    ", Word.InsertLocation.end);
            spaceRun.font.name = "KFGQPC HAFS Uthmanic Script";
            spaceRun.font.size = 18;
            spaceRun.font.color = "#000000";
          }

          const textRange = arabicPara.getRange(Word.RangeLocation.end);
          const textRun = textRange.insertText(a.arabic, Word.InsertLocation.end);
          textRun.font.name = "KFGQPC HAFS Uthmanic Script";
          textRun.font.size = 18;
          textRun.font.color = "#000000";

          if (showAyahNumber) {
            const markerRange = arabicPara.getRange(Word.RangeLocation.end);
            const markerRun = markerRange.insertText(
              buildVerseMarker(a.number),
              Word.InsertLocation.end
            );
            markerRun.font.name = "KFGQPC HAFS Uthmanic Script";
            markerRun.font.size = 20;
            markerRun.font.color = "#000000";
          }
        }
      }

      await context.sync();

      // Insert translation lines
      for (let i = 0; i < translationLines.length; i++) {
        const line = translationLines[i];
        const para = body.insertParagraph(line.text, Word.InsertLocation.end);

        if (line.isReference) {
          para.font.name = "Calibri";
          para.font.size = 9;
          para.font.color = "#888888";
          para.alignment = Word.Alignment.left;
          para.spaceAfter = 12;
        } else {
          const lang = line.langId ? getLanguageById(line.langId) : null;
          const fontName = lang && lang.fontName ? lang.fontName : null;
          const isRtl = lang && lang.dir === "rtl";

          if (fontName) {
            para.font.name = fontName;
          }

          para.font.size = 11;
          para.font.italic = true;
          para.font.color = "#444444";
          para.alignment = isRtl ? Word.Alignment.right : Word.Alignment.left;
          para.spaceAfter = 2;
        }
      }

      await context.sync();
    });

    const info = getSurahInfo(surahNum);
    const rangeStr = from === to ? `${from}` : `${from}-${to}`;
    setStatus(`Inserted QS. ${info.name}: ${rangeStr}`, false);
  } catch (error) {
    setStatus("Error: " + error.message, true);
  }
}

async function insertPageToWord() {
  const pageNum = parseInt(document.getElementById("mushaf-page").value, 10);
  if (isNaN(pageNum) || pageNum < 1 || pageNum > 604) {
    setStatus("Please enter a valid page number (1-604).", true);
    return;
  }

  setStatus("Loading page data...", false);
  try {
    const pageData = await loadPageData(pageNum);
    if (pageData.length === 0) {
      setStatus("No data available for this page.", true);
      return;
    }

    const pageInfo = getPageInfo(pageNum);

    await Word.run(async (context) => {
      const body = context.document.body;

      const arabicPara = body.insertParagraph("", Word.InsertLocation.end);
      arabicPara.font.name = "KFGQPC HAFS Uthmanic Script";
      arabicPara.font.size = 18;
      arabicPara.font.color = "#000000";
      arabicPara.alignment = Word.Alignment.right;
      arabicPara.lineSpacing = 16;
      arabicPara.spaceAfter = 0;
      arabicPara.spaceBefore = 0;
      arabicPara.rightIndent = 0;
      arabicPara.leftIndent = 0;
      arabicPara.firstLineIndent = 0;

      await context.sync();

      for (let i = 0; i < pageData.length; i++) {
        const a = pageData[i];

        if (i > 0) {
          const spaceRange = arabicPara.getRange(Word.RangeLocation.end);
          const spaceRun = spaceRange.insertText("    ", Word.InsertLocation.end);
          spaceRun.font.name = "KFGQPC HAFS Uthmanic Script";
          spaceRun.font.size = 18;
          spaceRun.font.color = "#000000";
        }

        const textRange = arabicPara.getRange(Word.RangeLocation.end);
        const textRun = textRange.insertText(a.arabic, Word.InsertLocation.end);
        textRun.font.name = "KFGQPC HAFS Uthmanic Script";
        textRun.font.size = 18;
        textRun.font.color = "#000000";

        const markerRange = arabicPara.getRange(Word.RangeLocation.end);
        const markerRun = markerRange.insertText(buildVerseMarker(a.ayah), Word.InsertLocation.end);
        markerRun.font.name = "KFGQPC HAFS Uthmanic Script";
        markerRun.font.size = 20;
        markerRun.font.color = "#000000";
      }

      await context.sync();

      // Add page break after mushaf page content
      body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
      await context.sync();
    });

    setStatus(`Inserted Page ${pageNum}: ${pageInfo.info}`, false);
  } catch (error) {
    setStatus("Error: " + error.message, true);
  }
}
