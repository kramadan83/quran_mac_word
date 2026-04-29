# Quran in Word (for Mac)

A free, open-source Microsoft Word add-in for macOS that lets you insert Quranic verses with Arabic text and translations (English & Indonesian) directly into your Word documents.

## Features

- **Full Quran** - All 114 surahs, 6,236 ayahs
- **Arabic Uthmani text** - Sourced from quran.com (text_uthmani) with KFGQPC HAFS Uthmanic Script font
- **Translations** - English (Sahih International) and Indonesian (Kemenag RI)
- **Two insert modes** - Single ayah or ayah range
- **Range layout options** - Continuous (mushaf-style) or one ayah per line
- **Searchable surah selector** - Filter by surah name, number, or Arabic name
- **Verse numbering** - Optional Arabic-Indic ayah markers in mushaf style
- **Offline support** - Service worker caches all assets after first load
- **No server required** - Hosted on GitHub Pages, no localhost needed

## Compatibility

| Requirement | Minimum |
|---|---|
| **macOS** | 10.15 (Catalina) or later |
| **Microsoft Word** | Version 16.9 or later |

### Compatible Word versions

| Edition | Compatible |
|---|---|
| Microsoft 365 for Mac | Yes |
| Office 2024 for Mac | Yes |
| Office 2021 for Mac | Yes |
| Office 2019 for Mac (v16.9+) | Yes |
| Office 2016 for Mac | No |

## Installation

### One-line install

Open Terminal and run:

```bash
curl -fsSL https://kramadan83.github.io/quran_mac_word/install.sh | bash
```

Then:
1. Quit Word completely (Cmd+Q)
2. Open Word
3. Go to **Insert** > **My Add-ins** > **Shared Folder** tab
4. Select **"Quran in Word Mac"** and click **Add**
5. The button appears in the **Home** tab ribbon

### Manual install

**Step 1** - Install the Arabic font:
```bash
curl -L -o ~/Library/Fonts/UthmanicHafs1Ver18.ttf \
  https://kramadan83.github.io/quran_mac_word/fonts/UthmanicHafs1Ver18.ttf
```

**Step 2** - Install the add-in manifest:
```bash
mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef
curl -o ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/manifest.xml \
  https://kramadan83.github.io/quran_mac_word/manifest.xml
```

**Step 3** - Restart Word and load the add-in from Insert > My Add-ins > Shared Folder.

### Uninstall

```bash
curl -fsSL https://kramadan83.github.io/quran_mac_word/install.sh | bash -s -- --uninstall
```

## Architecture

```
quran_addins_word/
+-- src/
|   +-- taskpane/
|   |   +-- taskpane.html      # Add-in UI (taskpane panel)
|   |   +-- taskpane.js        # Core logic: data loading, Word API, search
|   |   +-- taskpane.css       # Styles (Fluent UI based)
|   +-- commands/
|   |   +-- commands.html      # Ribbon command handler
|   |   +-- commands.js
|   +-- data/
|   |   +-- surahList.json     # Surah metadata (114 entries)
|   |   +-- arabic/*.json      # Arabic text per surah (from quran.com API v4)
|   |   +-- english/*.json     # English translation per surah (Sahih International)
|   |   +-- indonesian/*.json  # Indonesian translation per surah (Kemenag RI)
|   +-- fonts/
|   |   +-- UthmanicHafs1Ver18.ttf  # KFGQPC HAFS Uthmanic Script font
|   +-- service-worker.js      # Offline caching (stale-while-revalidate)
+-- assets/                    # Add-in icons (16, 32, 64, 80, 128px)
+-- manifest.xml               # Office add-in manifest (dev, localhost)
+-- install.sh                 # One-line installer/uninstaller
+-- webpack.config.js          # Build config (dev + production)
+-- package.json
```

### How it works

```
+-------------------+     +--------------------+     +------------------+
|   Word for Mac    |     |  GitHub Pages      |     |  quran.com API   |
|                   |     |  (Static hosting)  |     |  (Data source)   |
|  +-------------+  |     |                    |     |                  |
|  | Ribbon Btn  |--+---->| taskpane.html/js   |     |  text_uthmani    |
|  +-------------+  |     | (webpack bundles)  |     |  (Uthmani text)  |
|                   |     |                    |     +------------------+
|  +-------------+  |     | arabic/*.json      |
|  | Taskpane    |<-+-----| english/*.json     |     +------------------+
|  | (WebView)   |  |     | indonesian/*.json  |     |  risan/quran-json|
|  +-------------+  |     |                    |     |  (Translations)  |
|        |          |     | service-worker.js  |     +------------------+
|        v          |     | (offline cache)    |
|  +-------------+  |     |                    |
|  | Word Doc    |  |     | fonts/             |
|  | (Insert API)|  |     | UthmanicHafs.ttf   |
|  +-------------+  |     +--------------------+
+-------------------+
```

### Data flow

1. User selects a surah and ayah range in the taskpane
2. Surah data is lazy-loaded via webpack dynamic imports (`import()`)
3. Arabic text is cleaned (waqf marks stripped for Word compatibility)
4. Text is inserted into Word via the Office JS API (`Word.run()`)
5. In continuous layout, all ayahs are joined in one paragraph; in per-line layout, each ayah gets its own paragraph
6. Arabic text and verse markers are inserted as separate runs with different font sizes
7. Translations are inserted as separate paragraphs below

### Key technical decisions

| Decision | Reason |
|---|---|
| **Dynamic imports per surah** | Lazy-load only the surahs needed, not all 4.8MB at once |
| **Waqf marks stripped** | Word on Mac renders combining marks (U+06D6-U+06DC) as dotted circles |
| **No U+06DD for verse numbers** | Word on Mac renders it as a separate blank circle alongside the digit |
| **Arabic-Indic digits for markers** | KFGQPC font renders these inside ornamental circles natively |
| **Stale-while-revalidate SW** | Serves cached assets instantly, updates in background |
| **Font URL relative in CSS** | Required for GitHub Pages subdirectory deployment (`/quran_mac_word/`) |
| **Footnote refs stripped** | Indonesian source data contains `7)`, `8)` artifacts from print edition |

### Data sources

| Data | Source | Notes |
|---|---|---|
| Arabic (Uthmani) | [quran.com API v4](https://api.quran.com/api/v4/quran/verses/uthmani) | Official Uthmani text |
| English | [risan/quran-json](https://github.com/risan/quran-json) | Sahih International translation |
| Indonesian | [risan/quran-json](https://github.com/risan/quran-json) | Kemenag RI translation |
| Font | [quran.com GitHub](https://github.com/nickvdyck/quran.com-frontend) | UthmanicHafs v18 (KFGQPC HAFS) |

## Development

### Prerequisites

- Node.js 16+
- npm

### Setup

```bash
npm install
```

### Run locally (dev mode)

```bash
npm start
```

This starts a webpack dev server at `https://localhost:3000` with the dev manifest.

### Production build

```bash
npx webpack --mode production
```

Output goes to `dist/`. All `localhost` URLs in the manifest are replaced with the GitHub Pages URL.

### Deploy to GitHub Pages

```bash
npx gh-pages -d dist
```

## License

Free and open source.
