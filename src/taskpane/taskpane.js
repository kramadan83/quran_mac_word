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

// Pre-import available data (Surah 1)
import arabic1 from "../data/arabic/1.json";
import english1 from "../data/english/1.json";
import indonesian1 from "../data/indonesian/1.json";

dataCache.arabic[1] = arabic1;
dataCache.english[1] = english1;
dataCache.indonesian[1] = indonesian1;

// --- Helpers ---

const ARABIC_INDIC_ZERO = 0x0660; // ٠

function toArabicIndic(num) {
  return String(num)
    .split("")
    .map((d) => String.fromCharCode(ARABIC_INDIC_ZERO + parseInt(d, 10)))
    .join("");
}

function buildVerseMarker(ayahNumber) {
  return " \uFD3F" + toArabicIndic(ayahNumber) + "\uFD3E ";
}

// --- Init ---

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    initUI();
  }
});

function initUI() {
  populateSurahDropdown();
  updateAyahRange();
  updatePreview();

  document.getElementById("surah-select").addEventListener("change", () => {
    updateAyahRange();
    resetAyahInputs();
    updatePreview();
  });

  document.getElementById("ayah-from").addEventListener("input", () => {
    clampAyahInputs();
    updatePreview();
  });

  document.getElementById("ayah-to").addEventListener("input", () => {
    clampAyahInputs();
    updatePreview();
  });

  document.getElementById("chk-english").addEventListener("change", updatePreview);
  document.getElementById("chk-indonesian").addEventListener("change", updatePreview);
  document.getElementById("btn-insert").addEventListener("click", insertToWord);
}

// --- Surah / Ayah helpers ---

function populateSurahDropdown() {
  const select = document.getElementById("surah-select");
  surahList.forEach((s) => {
    const option = document.createElement("option");
    option.value = s.number;
    option.textContent = `${s.number}. ${s.name} (${s.arabic})`;
    select.appendChild(option);
  });
}

function getSelectedSurah() {
  return parseInt(document.getElementById("surah-select").value, 10);
}

function getSurahInfo(surahNumber) {
  return surahList.find((s) => s.number === surahNumber);
}

function getSelectedAyahRange() {
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

function updateAyahRange() {
  const info = getSurahInfo(getSelectedSurah());
  if (!info) return;
  const fromEl = document.getElementById("ayah-from");
  const toEl = document.getElementById("ayah-to");
  fromEl.max = info.total_ayah;
  toEl.max = info.total_ayah;
  document.getElementById("ayah-range").textContent = `/ ${info.total_ayah}`;
}

function resetAyahInputs() {
  const info = getSurahInfo(getSelectedSurah());
  if (!info) return;
  document.getElementById("ayah-from").value = 1;
  document.getElementById("ayah-to").value = info.total_ayah;
}

function clampAyahInputs() {
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
    arabic: arabicAyah ? arabicAyah.text : null,
    english: englishAyah ? englishAyah.text : null,
    indonesian: indonesianAyah ? indonesianAyah.text : null,
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

// --- Preview ---

function updatePreview() {
  const previewEl = document.getElementById("preview");
  const surahNum = getSelectedSurah();
  const { from, to } = getSelectedAyahRange();
  const showEnglish = document.getElementById("chk-english").checked;
  const showIndonesian = document.getElementById("chk-indonesian").checked;

  const ayahs = getAyahRangeData(surahNum, from, to);

  if (ayahs.length === 0) {
    const info = getSurahInfo(surahNum);
    previewEl.innerHTML = `<p style="color:#605e5c; font-style:italic;">Data for ${info ? info.name : "this surah"} is not yet loaded. Only Surah 1 (Al-Fatihah) has dummy data.</p>`;
    return;
  }

  // Mushaf-style Arabic block
  let arabicParts = ayahs.map(
    (a) => a.arabic + '<span class="preview-verse-marker">' + buildVerseMarker(a.number) + "</span>"
  );
  let html = '<div class="preview-arabic">' + arabicParts.join("") + "</div>";

  // English translations
  if (showEnglish) {
    html += '<div class="preview-translation-section">';
    html += '<div class="preview-label">English (Sahih International)</div>';
    ayahs.forEach((a) => {
      if (a.english) {
        html += `<div class="preview-translation-item">${a.number}. ${a.english}</div>`;
      }
    });
    html += "</div>";
  }

  // Indonesian translations
  if (showIndonesian) {
    html += '<div class="preview-translation-section">';
    html += '<div class="preview-label">Bahasa Indonesia</div>';
    ayahs.forEach((a) => {
      if (a.indonesian) {
        html += `<div class="preview-translation-item">${a.number}. ${a.indonesian}</div>`;
      }
    });
    html += "</div>";
  }

  previewEl.innerHTML = html;
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

function buildInsertHtml(surahNum, fromAyah, toAyah, showEnglish, showIndonesian) {
  const ayahs = getAyahRangeData(surahNum, fromAyah, toAyah);
  if (ayahs.length === 0) return null;

  const info = getSurahInfo(surahNum);
  const surahName = info ? info.name : `Surah ${surahNum}`;

  let html = "";

  // Arabic mushaf block - one continuous RTL paragraph with verse markers
  const arabicContent = ayahs
    .map((a) => a.arabic + buildVerseMarker(a.number))
    .join("");

  html += `<p dir="rtl" style="font-family: 'Geeza Pro', 'Traditional Arabic', 'Arabic Typesetting', serif; font-size: 22pt; line-height: 1.8; text-align: right; color: #1a1a1a; margin-bottom: 8pt;">${arabicContent}</p>`;

  // English translations - per ayah
  if (showEnglish) {
    html += `<p dir="ltr" style="font-family: 'Calibri', 'Segoe UI', sans-serif; font-size: 10pt; font-weight: bold; color: #333333; margin-bottom: 2pt;">English Translation (Sahih International)</p>`;
    ayahs.forEach((a) => {
      if (a.english) {
        html += `<p dir="ltr" style="font-family: 'Calibri', 'Segoe UI', sans-serif; font-size: 11pt; line-height: 1.5; color: #444444; margin-bottom: 2pt; font-style: italic;">${a.number}. ${a.english}</p>`;
      }
    });
  }

  // Indonesian translations - per ayah
  if (showIndonesian) {
    html += `<p dir="ltr" style="font-family: 'Calibri', 'Segoe UI', sans-serif; font-size: 10pt; font-weight: bold; color: #333333; margin-bottom: 2pt;">Terjemahan Bahasa Indonesia</p>`;
    ayahs.forEach((a) => {
      if (a.indonesian) {
        html += `<p dir="ltr" style="font-family: 'Calibri', 'Segoe UI', sans-serif; font-size: 11pt; line-height: 1.5; color: #444444; margin-bottom: 2pt; font-style: italic;">${a.number}. ${a.indonesian}</p>`;
      }
    });
  }

  // Reference line
  const rangeStr = fromAyah === toAyah ? `${fromAyah}` : `${fromAyah}-${toAyah}`;
  html += `<p dir="ltr" style="font-family: 'Calibri', 'Segoe UI', sans-serif; font-size: 9pt; color: #888888; margin-bottom: 8pt;">(QS. ${surahName}: ${rangeStr})</p>`;

  return html;
}

export async function insertToWord() {
  const surahNum = getSelectedSurah();
  const { from, to } = getSelectedAyahRange();
  const showEnglish = document.getElementById("chk-english").checked;
  const showIndonesian = document.getElementById("chk-indonesian").checked;

  const html = buildInsertHtml(surahNum, from, to, showEnglish, showIndonesian);

  if (!html) {
    setStatus("No data available for this ayah range.", true);
    return;
  }

  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const range = body.getRange(Word.RangeLocation.end);
      range.insertHtml(html, Word.InsertLocation.after);
      await context.sync();
    });

    const info = getSurahInfo(surahNum);
    const rangeStr = from === to ? `${from}` : `${from}-${to}`;
    setStatus(`Inserted QS. ${info.name}: ${rangeStr}`, false);
  } catch (error) {
    setStatus("Error: " + error.message, true);
  }
}
