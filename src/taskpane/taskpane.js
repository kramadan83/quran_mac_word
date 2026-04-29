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
  // U+06DD = End of Ayah mark (Medina mushaf style)
  // Digits following it render inside the decorative marker
  return " \u06DD" + toArabicIndic(ayahNumber) + " ";
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

  document.getElementById("surah-select").addEventListener("change", () => {
    updateAyahRange();
    resetAyahInputs();
  });

  document.getElementById("ayah-from").addEventListener("input", () => {
    clampAyahInputs();
  });

  document.getElementById("ayah-to").addEventListener("input", () => {
    clampAyahInputs();
  });

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

  const ayahs = getAyahRangeData(surahNum, from, to);
  if (ayahs.length === 0) {
    setStatus("No data available for this ayah range.", true);
    return;
  }

  const translationLines = buildTranslationLines(surahNum, from, to, showEnglish, showIndonesian);

  try {
    await Word.run(async (context) => {
      const body = context.document.body;

      // Insert empty paragraph and set its default font to KFGQPC
      const arabicPara = body.insertParagraph("", Word.InsertLocation.end);
      arabicPara.font.name = "KFGQPC HAFS Uthmanic Script";
      arabicPara.font.size = 18;
      arabicPara.font.color = "#000000";
      arabicPara.alignment = Word.Alignment.right;
      arabicPara.lineSpacing = 30;
      arabicPara.spaceAfter = 6;
      arabicPara.spaceBefore = 0;
      arabicPara.rightIndent = 0;
      arabicPara.leftIndent = 0;
      arabicPara.firstLineIndent = 0;

      // Sync to apply default font before inserting text
      await context.sync();

      // Insert each ayah text + marker as separate runs with different sizes
      for (let i = 0; i < ayahs.length; i++) {
        const a = ayahs[i];

        // Ayah text - full size
        const textRange = arabicPara.getRange(Word.RangeLocation.end);
        const textRun = textRange.insertText(a.arabic, Word.InsertLocation.end);
        textRun.font.name = "KFGQPC HAFS Uthmanic Script";
        textRun.font.size = 18;
        textRun.font.color = "#000000";

        // Verse marker - smaller size
        const markerRange = arabicPara.getRange(Word.RangeLocation.end);
        const markerRun = markerRange.insertText(buildVerseMarker(a.number), Word.InsertLocation.end);
        markerRun.font.name = "KFGQPC HAFS Uthmanic Script";
        markerRun.font.size = 11;
        markerRun.font.color = "#000000";
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
