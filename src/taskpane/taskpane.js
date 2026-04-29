/*
 * Quran in Word - Office Add-in
 * Inserts Quranic verses with translations into Word documents.
 * Supports mushaf-style rendering with continuous Arabic text and verse markers.
 */

/* global document, Office, Word */

import surahList from "../data/surahList.json";

// Data cache
const dataCache = {
  arabic: {},
  english: {},
  indonesian: {},
};

// Dynamically load surah data on demand
async function loadSurahData(surahNumber) {
  if (dataCache.arabic[surahNumber]) return; // Already loaded

  const [arabic, english, indonesian] = await Promise.all([
    import(/* webpackChunkName: "arabic-[request]" */ `../data/arabic/${surahNumber}.json`),
    import(/* webpackChunkName: "english-[request]" */ `../data/english/${surahNumber}.json`),
    import(/* webpackChunkName: "indonesian-[request]" */ `../data/indonesian/${surahNumber}.json`),
  ]);

  dataCache.arabic[surahNumber] = arabic.default || arabic;
  dataCache.english[surahNumber] = english.default || english;
  dataCache.indonesian[surahNumber] = indonesian.default || indonesian;
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

// Strip footnote reference numbers (e.g. "7)", "8)") from translation text
const FOOTNOTE_RE = /\d+\)/g;
function cleanTranslation(text) {
  return text.replace(FOOTNOTE_RE, "").replace(/  +/g, " ").trim();
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
  const english = dataCache.english[surahNumber];
  const indonesian = dataCache.indonesian[surahNumber];

  if (!arabic) return null;

  const arabicAyah = arabic.ayahs.find((a) => a.number === ayahNumber);
  const englishAyah = english ? english.ayahs.find((a) => a.number === ayahNumber) : null;
  const indonesianAyah = indonesian ? indonesian.ayahs.find((a) => a.number === ayahNumber) : null;

  return {
    number: ayahNumber,
    arabic: arabicAyah ? cleanArabicText(arabicAyah.text) : null,
    english: englishAyah ? englishAyah.text : null,
    indonesian: indonesianAyah ? cleanTranslation(indonesianAyah.text) : null,
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

function buildArabicText(surahNum, fromAyah, toAyah) {
  const ayahs = getAyahRangeData(surahNum, fromAyah, toAyah);
  if (ayahs.length === 0) return null;
  return ayahs.map((a) => a.arabic + buildVerseMarker(a.number)).join("");
}

function buildTranslationLines(surahNum, fromAyah, toAyah, showEnglish, showIndonesian) {
  const ayahs = getAyahRangeData(surahNum, fromAyah, toAyah);
  const info = getSurahInfo(surahNum);
  const surahName = info ? info.name : `Surah ${surahNum}`;
  const lines = [];

  if (showEnglish) {
    lines.push({ text: "English Translation (Sahih International)", isLabel: true });
    ayahs.forEach((a) => {
      if (a.english) lines.push({ text: `${a.number}. ${a.english}`, isLabel: false });
    });
  }

  if (showIndonesian) {
    lines.push({ text: "Terjemahan Bahasa Indonesia", isLabel: true });
    ayahs.forEach((a) => {
      if (a.indonesian) lines.push({ text: `${a.number}. ${a.indonesian}`, isLabel: false });
    });
  }

  const rangeStr = fromAyah === toAyah ? `${fromAyah}` : `${fromAyah}-${toAyah}`;
  lines.push({ text: `(QS. ${surahName}: ${rangeStr})`, isReference: true });

  return lines;
}

export async function insertToWord() {
  const surahNum = getSelectedSurah();
  const { from, to } = getSelectedAyahRange();
  const showEnglish = document.getElementById("chk-english").checked;
  const showIndonesian = document.getElementById("chk-indonesian").checked;
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

  const translationLines = buildTranslationLines(surahNum, from, to, showEnglish, showIndonesian);

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
        } else if (line.isLabel) {
          para.font.name = "Calibri";
          para.font.size = 10;
          para.font.color = "#333333";
          para.alignment = Word.Alignment.left;
          para.spaceAfter = 2;
          para.spaceBefore = 6;
        } else {
          para.font.name = "Calibri";
          para.font.size = 11;
          para.font.italic = true;
          para.font.color = "#444444";
          para.alignment = Word.Alignment.left;
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
