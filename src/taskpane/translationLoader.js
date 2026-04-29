/**
 * Translation Loader
 * Unified data fetcher for bundled and API-sourced translations.
 */

import { getLanguageById } from "./translationRegistry";

const QURAN_API_BASE = "https://api.quran.com/api/v4/quran/translations";
const HTML_TAG_RE = /<[^>]+>/g;
const FOOTNOTE_RE = /\d+\)/g;

/**
 * Load a single translation for a surah.
 * @param {string} langId - Language ID from registry (e.g. "en", "fr")
 * @param {number} surahNumber - Surah number (1-114)
 * @returns {Promise<{ayahs: Array<{number: number, text: string}>}>}
 */
export async function loadTranslation(langId, surahNumber) {
  const lang = getLanguageById(langId);
  if (!lang) throw new Error(`Unknown language: ${langId}`);

  if (lang.source === "bundled") {
    return loadBundled(lang, surahNumber);
  }
  return loadFromApi(lang, surahNumber);
}

async function loadBundled(lang, surahNumber) {
  let mod;
  if (lang.folder === "english") {
    mod = await import(
      /* webpackChunkName: "english-[request]" */ `../data/english/${surahNumber}.json`
    );
  } else if (lang.folder === "indonesian") {
    mod = await import(
      /* webpackChunkName: "indonesian-[request]" */ `../data/indonesian/${surahNumber}.json`
    );
  } else {
    throw new Error(`No bundled folder for language: ${lang.id}`);
  }
  const data = mod.default || mod;
  return {
    ayahs: data.ayahs.map((a) => ({
      number: a.number,
      text: cleanText(a.text, lang.cleanFn),
    })),
  };
}

async function loadFromApi(lang, surahNumber) {
  const url = `${QURAN_API_BASE}/${lang.editionId}?chapter_number=${surahNumber}`;
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`API error ${response.status} for ${lang.name}`);
  }
  const json = await response.json();
  const translations = json.translations || [];

  return {
    ayahs: translations.map((t) => {
      const verseKey = t.verse_key || "";
      const ayahNum = parseInt(verseKey.split(":")[1], 10) || 0;
      return {
        number: ayahNum,
        text: cleanText(t.text, lang.cleanFn),
      };
    }),
  };
}

function cleanText(text, cleanFn) {
  // Always strip HTML tags (API responses may contain <sup>, <a>, etc.)
  let cleaned = text.replace(HTML_TAG_RE, "");
  // Apply language-specific cleaning
  if (cleanFn === "stripFootnotes") {
    cleaned = cleaned.replace(FOOTNOTE_RE, "");
  }
  // Normalize whitespace
  return cleaned.replace(/  +/g, " ").trim();
}
