#!/usr/bin/env node

/**
 * Downloads clean Uthmani Arabic text from quran.com API v4.
 * Replaces corrupted Arabic data from risan/quran-json.
 * Includes waqf marks and proper Uthmani encoding.
 */

const https = require("https");
const fs = require("fs");
const path = require("path");

const surahListPath = path.join(__dirname, "../src/data/surahList.json");
const surahList = JSON.parse(fs.readFileSync(surahListPath, "utf8"));
const arabicDir = path.join(__dirname, "../src/data/arabic");

function fetch(url) {
  return new Promise((resolve, reject) => {
    https
      .get(url, { headers: { "User-Agent": "quran-addin-data-downloader" } }, (res) => {
        if (res.statusCode !== 200) {
          return reject(new Error(`HTTP ${res.statusCode} for ${url}`));
        }
        // Accumulate raw buffers to avoid splitting multi-byte UTF-8 chars
        const chunks = [];
        res.on("data", (chunk) => chunks.push(chunk));
        res.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
        res.on("error", reject);
      })
      .on("error", reject);
  });
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function main() {
  console.log("Downloading clean Uthmani Arabic text from quran.com API v4...\n");

  let successCount = 0;
  let errorCount = 0;

  for (const surah of surahList) {
    const num = surah.number;
    const url = `https://api.quran.com/api/v4/quran/verses/uthmani?chapter_number=${num}`;

    try {
      const raw = await fetch(url);
      const data = JSON.parse(raw);

      if (!data.verses || data.verses.length === 0) {
        console.error(`  ERROR: No verses returned for surah ${num}`);
        errorCount++;
        continue;
      }

      const output = {
        surah: num,
        name: surah.arabic,
        transliteration: surah.name,
        total_ayah: surah.total_ayah,
        ayahs: data.verses.map((v) => {
          // verse_key is "chapter:verse"
          const verseNum = parseInt(v.verse_key.split(":")[1], 10);
          return {
            number: verseNum,
            text: v.text_uthmani.trim(),
          };
        }),
      };

      // Verify count matches
      if (output.ayahs.length !== surah.total_ayah) {
        console.warn(
          `  WARNING: Surah ${num} expected ${surah.total_ayah} ayahs but got ${output.ayahs.length}`
        );
      }

      // Check for any U+FFFD corruption
      let corrupted = 0;
      for (const a of output.ayahs) {
        corrupted += (a.text.match(/\uFFFD/g) || []).length;
      }
      if (corrupted > 0) {
        console.warn(`  WARNING: Surah ${num} still has ${corrupted} corrupted chars`);
      }

      fs.writeFileSync(
        path.join(arabicDir, `${num}.json`),
        JSON.stringify(output, null, 2) + "\n"
      );

      successCount++;
      process.stdout.write(`  Surah ${num}/114 (${surah.name}) - ${output.ayahs.length} ayahs ✓\n`);

      // Rate limit: small delay between requests
      if (num < 114) await sleep(200);
    } catch (err) {
      console.error(`  ERROR: Surah ${num}: ${err.message}`);
      errorCount++;
    }
  }

  console.log(`\nDone! Success: ${successCount}, Errors: ${errorCount}`);
}

main().catch((err) => {
  console.error("Fatal error:", err);
  process.exit(1);
});
