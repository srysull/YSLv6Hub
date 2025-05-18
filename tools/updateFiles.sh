#!/bin/bash
# Script to update files in Google Apps Script project

# Exit on error
set -e

echo "Updating Google Apps Script project with local files..."

# Navigate to project root
cd "$(dirname "$0")/.."

# Build the project
echo "Building TypeScript files..."
npm run build
echo "Build completed."

# Copy appsscript.json to build directory
cp appsscript.json build/

# Create a directory to prepare files for CLASP
mkdir -p dist

# Copy all built files to a single directory without subdirectories
echo "Preparing files for CLASP..."
cp build/appsscript.json dist/
cp build/00_System.js dist/
cp build/01_Core.js dist/
cp build/10_SystemLog.js dist/
cp build/interfaces/index.js dist/interfaces_index.js
cp build/utils/constants.js dist/utils_constants.js

# Update .clasp.json to point to dist directory
echo '{
  "scriptId": "1jvHLWHyckIleHMWuNiCYLOrfQqBiaBogdP-4rCh1QEa5023Hoal-j1r_",
  "rootDir": "dist"
}' > .clasp.json

# Push to Google Apps Script
echo "Pushing to Google Apps Script..."
clasp push --force

echo "Script updated successfully!"
echo "Open the script in the browser with:"
echo "  clasp open-script"