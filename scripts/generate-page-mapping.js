/**
 * Generate Madinah Mushaf page-to-ayah mapping.
 * Fetches data from quran.com API v4 and creates src/data/pageToAyahs.json
 *
 * Usage: node scripts/generate-page-mapping.js
 */

const fs = require("fs");
const path = require("path");

const API_BASE = "https://api.quran.com/api/v4";
const TOTAL_PAGES = 604;
const DELAY_MS = 250;
const OUTPUT_PATH = path.join(__dirname, "..", "src", "data", "pageToAyahs.json");

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function fetchPage(pageNum) {
  const url = `${API_BASE}/verses/by_page/${pageNum}?language=en&words=false`;
  const res = await fetch(url);
  if (!res.ok) {
    throw new Error(`Failed to fetch page ${pageNum}: ${res.status} ${res.statusText}`);
  }
  const data = await res.json();
  return data.verses || [];
}

function parseVerseKey(verseKey) {
  const [surah, ayah] = verseKey.split(":").map(Number);
  return { surah, ayah };
}

async function main() {
  console.log("Generating Madinah Mushaf page-to-ayah mapping...\n");

  const pageMapping = {};

  for (let page = 1; page <= TOTAL_PAGES; page++) {
    try {
      const verses = await fetchPage(page);
      pageMapping[String(page)] = verses.map((v) => {
        const { surah, ayah } = parseVerseKey(v.verse_key);
        return { surah, ayah };
      });

      const count = pageMapping[String(page)].length;
      process.stdout.write(`\rPage ${page}/${TOTAL_PAGES} - ${count} ayahs`);

      if (page < TOTAL_PAGES) {
        await sleep(DELAY_MS);
      }
    } catch (err) {
      console.error(`\nError on page ${page}: ${err.message}`);
      console.log("Retrying in 2 seconds...");
      await sleep(2000);
      // Retry once
      const verses = await fetchPage(page);
      pageMapping[String(page)] = verses.map((v) => {
        const { surah, ayah } = parseVerseKey(v.verse_key);
        return { surah, ayah };
      });
      process.stdout.write(`\rPage ${page}/${TOTAL_PAGES} - ${pageMapping[String(page)].length} ayahs (retry)`);
    }
  }

  console.log("\n\nWriting to", OUTPUT_PATH);
  fs.mkdirSync(path.dirname(OUTPUT_PATH), { recursive: true });
  fs.writeFileSync(OUTPUT_PATH, JSON.stringify(pageMapping, null, 2));

  // Stats
  const totalAyahs = Object.values(pageMapping).reduce((sum, arr) => sum + arr.length, 0);
  const fileSize = fs.statSync(OUTPUT_PATH).size;

  console.log(`\nDone! Generated mapping for ${TOTAL_PAGES} pages with ${totalAyahs} total ayah entries.`);
  console.log(`File size: ${(fileSize / 1024).toFixed(1)} KB`);

  // Spot-check some known pages
  const checks = [
    { page: 1, expected: "Al-Fatihah" },
    { page: 3, expected: "Al-Baqarah 6-16" },
    { page: 604, expected: "An-Nas" },
  ];

  console.log("\nSpot checks:");
  for (const { page, expected } of checks) {
    const ayahs = pageMapping[String(page)];
    const first = ayahs[0];
    const last = ayahs[ayahs.length - 1];
    console.log(
      `  Page ${page}: Surah ${first.surah}:${first.ayah} to Surah ${last.surah}:${last.ayah} (${ayahs.length} ayahs) [${expected}]`
    );
  }
}

main().catch((err) => {
  console.error("Fatal error:", err);
  process.exit(1);
});
