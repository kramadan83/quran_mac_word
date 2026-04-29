/*
 * Quran in Word - Office Add-in
 * Inserts Quranic verses with translations into Word documents.
 * Supports mushaf-style rendering with continuous Arabic text and verse markers.
 */

/* global document, Office, Word */

import surahList from "../data/surahList.json";
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
}

// --- Language Selector ---

const STORAGE_KEY = "quran-word-active-languages";
const MAX_LANGUAGES = 3;
const MIN_LANGUAGES = 1;

function loadLanguagePreferences() {
  try {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (stored) {
      const parsed = JSON.parse(stored);
      // Validate: filter to known language IDs
      const valid = parsed.filter((id) => getLanguageById(id));
      if (valid.length >= MIN_LANGUAGES && valid.length <= MAX_LANGUAGES) {
        return valid;
      }
    }
  } catch (_) {
    // ignore parse errors
  }
  return getDefaultLanguageIds();
}

function saveLanguagePreferences() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(activeLanguages));
}

function initLanguageSelector() {
  activeLanguages = loadLanguagePreferences();
  renderLanguageChips();

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

function renderLanguageChips() {
  const container = document.getElementById("active-languages");
  const canRemove = activeLanguages.length > MIN_LANGUAGES;

  container.innerHTML = activeLanguages
    .map((id) => {
      const lang = getLanguageById(id);
      if (!lang) return "";
      return (
        `<span class="language-chip" data-lang="${id}">` +
        `<span class="language-chip__name">${lang.name}</span>` +
        (canRemove
          ? `<button class="language-chip__remove" data-lang="${id}" title="Remove ${lang.name}">\u00d7</button>`
          : "") +
        `</span>`
      );
    })
    .join("");

  container.querySelectorAll(".language-chip__remove").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      removeLanguage(btn.dataset.lang);
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
  renderLanguageChips();
}

function removeLanguage(id) {
  if (activeLanguages.length <= MIN_LANGUAGES) return;
  activeLanguages = activeLanguages.filter((l) => l !== id);
  saveLanguagePreferences();
  renderLanguageChips();
}

function getActiveLanguages() {
  return activeLanguages;
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

    lines.push({ text: lang.sectionLabel, isLabel: true, langId });
    ayahs.forEach((a) => {
      const text = a.translations[langId];
      if (text) {
        lines.push({ text: `${a.number}. ${text}`, isLabel: false, langId });
      }
    });
  }

  const rangeStr = fromAyah === toAyah ? `${fromAyah}` : `${fromAyah}-${toAyah}`;
  lines.push({ text: `(QS. ${surahName}: ${rangeStr})`, isReference: true });

  return lines;
}

export async function insertToWord() {
  const surahNum = getSelectedSurah();
  const { from, to } = getSelectedAyahRange();
  const langs = getActiveLanguages();
  const isSingleMode = getInsertMode() === "single";
  const showAyahNumber = isSingleMode
    ? document.getElementById("chk-show-ayah-number").checked
    : true; // always show in range mode

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

          const markerRange = para.getRange(Word.RangeLocation.end);
          const markerRun = markerRange.insertText(buildVerseMarker(a.number), Word.InsertLocation.end);
          markerRun.font.name = "KFGQPC HAFS Uthmanic Script";
          markerRun.font.size = 20;
          markerRun.font.color = "#000000";
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

          const textRange = arabicPara.getRange(Word.RangeLocation.end);
          const textRun = textRange.insertText(a.arabic, Word.InsertLocation.end);
          textRun.font.name = "KFGQPC HAFS Uthmanic Script";
          textRun.font.size = 18;
          textRun.font.color = "#000000";

          if (showAyahNumber) {
            const markerRange = arabicPara.getRange(Word.RangeLocation.end);
            const markerRun = markerRange.insertText(buildVerseMarker(a.number), Word.InsertLocation.end);
            markerRun.font.name = "KFGQPC HAFS Uthmanic Script";
            markerRun.font.size = 20;
            markerRun.font.color = "#000000";
          }
        }
      }

      // Sync to commit Arabic text
      await context.sync();

      // Insert translation lines with per-language font handling
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
          // Look up language config for font and direction
          const lang = line.langId ? getLanguageById(line.langId) : null;
          const fontName = lang && lang.fontName ? lang.fontName : null;
          const isRtl = lang && lang.dir === "rtl";

          if (fontName) {
            para.font.name = fontName;
          }
          // If fontName is null, don't set font - let Word use system fallback

          if (line.isLabel) {
            para.font.size = 10;
            para.font.color = "#333333";
            para.alignment = isRtl ? Word.Alignment.right : Word.Alignment.left;
            para.spaceAfter = 2;
            para.spaceBefore = 6;
          } else {
            para.font.size = 11;
            para.font.italic = true;
            para.font.color = "#444444";
            para.alignment = isRtl ? Word.Alignment.right : Word.Alignment.left;
            para.spaceAfter = 2;
          }
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
