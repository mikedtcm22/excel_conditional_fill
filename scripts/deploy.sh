#!/bin/bash

# deploy.sh - Production build and deployment script for FillGaps Excel Add-in
# Builds assets to /docs folder for GitHub Pages deployment

set -e  # Exit on any error

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"

echo "=== FillGaps Production Build ==="
echo "Project root: $PROJECT_ROOT"
echo ""

# Step 1: Navigate to project root
cd "$PROJECT_ROOT"

# Step 2: Clean and rebuild
echo "Step 1: Cleaning /docs folder..."
rm -rf docs
mkdir -p docs

echo "Step 2: Running production build..."
npm run build:prod

# Step 3: Verify required files exist
echo "Step 3: Verifying build output..."

REQUIRED_FILES=(
  "docs/taskpane.html"
  "docs/commands.html"
  "docs/taskpane.js"
  "docs/commands.js"
  "docs/shortcuts.json"
  "docs/manifest-production.xml"
  "docs/assets/icon-16.png"
  "docs/assets/icon-32.png"
  "docs/assets/icon-64.png"
  "docs/assets/icon-80.png"
)

MISSING_FILES=()
for file in "${REQUIRED_FILES[@]}"; do
  if [ ! -f "$file" ]; then
    MISSING_FILES+=("$file")
  fi
done

if [ ${#MISSING_FILES[@]} -ne 0 ]; then
  echo "ERROR: Missing required files:"
  for file in "${MISSING_FILES[@]}"; do
    echo "  - $file"
  done
  exit 1
fi

echo "All required files present."
echo ""

# Step 4: Display build summary
echo "=== Build Summary ==="
echo "Files in /docs folder:"
find docs -type f | sort | while read -r file; do
  SIZE=$(ls -lh "$file" | awk '{print $5}')
  echo "  $file ($SIZE)"
done

echo ""
echo "=== Build Complete ==="
echo ""
echo "To deploy to GitHub Pages:"
echo "  1. Commit the /docs folder: git add docs && git commit -m 'Production build'"
echo "  2. Push to main branch: git push origin main"
echo "  3. Ensure GitHub Pages is configured to serve from /docs on main branch"
echo ""
echo "GitHub Pages URL: https://mikedtcm22.github.io/excel_conditional_fill/"
