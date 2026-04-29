#!/usr/bin/env node

/**
 * Downloads Quran data from risan/quran-json and converts
 * into per-surah JSON files for Arabic, English, Indonesian.
 */

const https = require("https");
const fs = require("fs");
const path = require("path");

const BASE_URL =
  "https://raw.githubusercontent.com/risan/quran-json/main/data";

const SOURCES = {
  arabic: `${BASE_URL}/quran.json`,
  english: `${BASE_URL}/editions/en.json`,
  indonesian: `${BASE_URL}/editions/id.json`,
};

// Surah metadata (name + arabic name) from surahList.json
const surahListPath = path.join(__dirname, "../src/data/surahList.json");
const surahList = JSON.parse(fs.readFileSync(surahListPath, "utf8"));

function fetch(url) {
  return new Promise((resolve, reject) => {
    https
      .get(url, { headers: { "User-Agent": "quran-addin-data-downloader" } }, (res) => {
        if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
          return fetch(res.headers.location).then(resolve).catch(reject);
        }
        if (res.statusCode !== 200) {
          return reject(new Error(`HTTP ${res.statusCode} for ${url}`));
        }
        const chunks = [];
        res.on("data", (chunk) => chunks.push(chunk));
        res.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
        res.on("error", reject);
      })
      .on("error", reject);
  });
}

function ensureDir(dir) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

async function main() {
  console.log("Downloading Quran data from risan/quran-json...\n");

  // Download all three sources in parallel
  const [arabicRaw, englishRaw, indonesianRaw] = await Promise.all([
    fetch(SOURCES.arabic),
    fetch(SOURCES.english),
    fetch(SOURCES.indonesian),
  ]);

  console.log("Downloaded all sources. Parsing...");

  const arabicData = JSON.parse(arabicRaw);
  const englishData = JSON.parse(englishRaw);
  const indonesianData = JSON.parse(indonesianRaw);

  const dataDir = path.join(__dirname, "../src/data");
  ensureDir(path.join(dataDir, "arabic"));
  ensureDir(path.join(dataDir, "english"));
  ensureDir(path.join(dataDir, "indonesian"));

  let totalFiles = 0;

  for (const surah of surahList) {
    const num = surah.number;
    const key = String(num);

    // Arabic
    const arabicAyahs = arabicData[key];
    if (arabicAyahs) {
      const arabicOut = {
        surah: num,
        name: surah.arabic,
        transliteration: surah.name,
        total_ayah: surah.total_ayah,
        ayahs: arabicAyahs.map((a) => ({
          number: a.verse,
          text: a.text,
        })),
      };
      fs.writeFileSync(
        path.join(dataDir, "arabic", `${num}.json`),
        JSON.stringify(arabicOut, null, 2) + "\n"
      );
      totalFiles++;
    } else {
      console.warn(`  WARNING: No Arabic data for surah ${num}`);
    }

    // English
    const englishAyahs = englishData[key];
    if (englishAyahs) {
      const englishOut = {
        surah: num,
        name: surah.name,
        translator: "Sahih International",
        total_ayah: surah.total_ayah,
        ayahs: englishAyahs.map((a) => ({
          number: a.verse,
          text: a.text,
        })),
      };
      fs.writeFileSync(
        path.join(dataDir, "english", `${num}.json`),
        JSON.stringify(englishOut, null, 2) + "\n"
      );
      totalFiles++;
    } else {
      console.warn(`  WARNING: No English data for surah ${num}`);
    }

    // Indonesian
    const indonesianAyahs = indonesianData[key];
    if (indonesianAyahs) {
      const indonesianOut = {
        surah: num,
        name: surah.name,
        translator: "Kementerian Agama RI",
        total_ayah: surah.total_ayah,
        ayahs: indonesianAyahs.map((a) => ({
          number: a.verse,
          text: a.text,
        })),
      };
      fs.writeFileSync(
        path.join(dataDir, "indonesian", `${num}.json`),
        JSON.stringify(indonesianOut, null, 2) + "\n"
      );
      totalFiles++;
    } else {
      console.warn(`  WARNING: No Indonesian data for surah ${num}`);
    }
  }

  console.log(`\nDone! Generated ${totalFiles} files (114 surahs x 3 languages = 342 expected).`);
}

main().catch((err) => {
  console.error("Error:", err);
  process.exit(1);
});
