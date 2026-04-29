#!/bin/bash
# Quran in Word (for Mac) - Install / Uninstall script
# Usage:
#   Install:    bash install.sh
#   Uninstall:  bash install.sh --uninstall

set -e

BASE_URL="https://kramadan83.github.io/quran_mac_word"
FONT_DIR="$HOME/Library/Fonts"
WEF_DIR="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"
FONT_FILE="UthmanicHafs1Ver18.ttf"
MANIFEST_FILE="manifest.xml"

# Colors
GREEN='\033[0;32m'
RED='\033[0;31m'
BLUE='\033[0;34m'
BOLD='\033[1m'
NC='\033[0m'

print_header() {
  echo ""
  echo -e "${BLUE}${BOLD}  Quran in Word (for Mac)${NC}"
  echo -e "  Insert Quranic verses into your document"
  echo ""
}

# --- UNINSTALL ---
if [ "$1" = "--uninstall" ] || [ "$1" = "uninstall" ]; then
  print_header
  echo -e "${BOLD}Uninstalling...${NC}"
  echo ""

  # Remove manifest
  if [ -f "$WEF_DIR/$MANIFEST_FILE" ]; then
    rm -f "$WEF_DIR/$MANIFEST_FILE"
    echo -e "  ${GREEN}Removed add-in manifest${NC}"
  else
    echo "  Manifest not found (already removed)"
  fi

  # Remove font
  if [ -f "$FONT_DIR/$FONT_FILE" ]; then
    rm -f "$FONT_DIR/$FONT_FILE"
    echo -e "  ${GREEN}Removed Arabic font${NC}"
  else
    echo "  Font not found (already removed)"
  fi

  # Clear Office cache
  rm -rf "$HOME/Library/Containers/com.microsoft.Word/Data/Library/Caches/Microsoft/Office/16.0/Wef/" 2>/dev/null
  echo -e "  ${GREEN}Cleared Office add-in cache${NC}"

  echo ""
  echo -e "${GREEN}${BOLD}  Uninstall complete.${NC}"
  echo "  Please restart Word for changes to take effect."
  echo ""
  exit 0
fi

# --- INSTALL ---
print_header
echo -e "${BOLD}Installing...${NC}"
echo ""

# Check if Word is installed
if [ ! -d "/Applications/Microsoft Word.app" ]; then
  echo -e "  ${RED}Microsoft Word for Mac is not installed.${NC}"
  echo "  Please install Word first, then run this script again."
  exit 1
fi

# Check Word version
WORD_VER=$(defaults read "/Applications/Microsoft Word.app/Contents/Info.plist" CFBundleShortVersionString 2>/dev/null || echo "unknown")
echo -e "  Word version: ${BOLD}$WORD_VER${NC}"

# Step 1: Install font
echo -n "  Installing Arabic font... "
mkdir -p "$FONT_DIR"
if curl -sfL -o "$FONT_DIR/$FONT_FILE" "$BASE_URL/fonts/$FONT_FILE"; then
  echo -e "${GREEN}done${NC}"
else
  echo -e "${RED}failed${NC}"
  echo "  Could not download font. Check your internet connection."
  exit 1
fi

# Step 2: Install manifest
echo -n "  Installing add-in manifest... "
mkdir -p "$WEF_DIR"
if curl -sfL -o "$WEF_DIR/$MANIFEST_FILE" "$BASE_URL/$MANIFEST_FILE"; then
  echo -e "${GREEN}done${NC}"
else
  echo -e "${RED}failed${NC}"
  echo "  Could not download manifest. Check your internet connection."
  exit 1
fi

# Step 3: Clear old cache
rm -rf "$HOME/Library/Containers/com.microsoft.Word/Data/Library/Caches/Microsoft/Office/16.0/Wef/" 2>/dev/null
echo -e "  Cleared Office cache... ${GREEN}done${NC}"

# Done
echo ""
echo -e "${GREEN}${BOLD}  Installation complete!${NC}"
echo ""
echo "  Next steps:"
echo "    1. Quit Word completely (Cmd+Q)"
echo "    2. Open Word"
echo "    3. Go to Insert > My Add-ins > Shared Folder"
echo "    4. Select \"Quran in Word Mac\" and click Add"
echo ""
echo "  To uninstall later:"
echo "    bash install.sh --uninstall"
echo ""
