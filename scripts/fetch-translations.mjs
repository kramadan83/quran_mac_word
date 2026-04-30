/**
 * Fetch English and Indonesian translations from quran.com API v4
 * and save them in the bundled JSON format.
 *
 * Usage: node scripts/fetch-translations.mjs
 */

import { readFileSync, writeFileSync, mkdirSync } from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = join(__dirname, "..");

const SURAH_LIST = JSON.parse(
  readFileSync(join(ROOT, "src/data/surahList.json"), "utf-8")
);

const EDITIONS = [
  {
    id: 20,
    lang: "english",
    translator: "Sahih International",
  },
  {
    id: 33,
    lang: "indonesian",
    translator: "Kementerian Agama RI",
  },
];

const API_BASE = "https://api.quran.com/api/v4/quran/translations";
const DELAY_MS = 300; // polite rate-limiting

function stripHtml(text) {
  return text
    .replace(/<sup[^>]*>.*?<\/sup>/gi, "") // remove footnote superscripts
    .replace(/<[^>]+>/g, "") // remove remaining HTML tags
    .replace(/\s+/g, " ") // collapse whitespace
    .trim();
}

async function fetchSurah(editionId, surahNumber) {
  const url = `${API_BASE}/${editionId}?chapter_number=${surahNumber}`;
  const res = await fetch(url);
  if (!res.ok) {
    throw new Error(`HTTP ${res.status} for ${url}`);
  }
  return res.json();
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function main() {
  for (const edition of EDITIONS) {
    const outDir = join(ROOT, "src/data", edition.lang);
    mkdirSync(outDir, { recursive: true });

    console.log(`\nFetching ${edition.lang} (edition ${edition.id})...`);

    for (const surah of SURAH_LIST) {
      const data = await fetchSurah(edition.id, surah.number);

      const ayahs = data.translations.map((t, idx) => ({
        number: idx + 1,
        text: stripHtml(t.text),
      }));

      const output = {
        surah: surah.number,
        name: surah.name,
        translator: edition.translator,
        total_ayah: surah.total_ayah,
        ayahs,
      };

      const filePath = join(outDir, `${surah.number}.json`);
      writeFileSync(filePath, JSON.stringify(output, null, 2) + "\n");

      process.stdout.write(
        `  ${surah.number}/114 ${surah.name} (${ayahs.length} ayahs)\n`
      );

      await sleep(DELAY_MS);
    }

    console.log(`Done: ${edition.lang}`);
  }

  console.log("\nAll translations fetched and saved.");
}

main().catch((err) => {
  console.error("Error:", err);
  process.exit(1);
});
