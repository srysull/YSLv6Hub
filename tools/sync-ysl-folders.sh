#!/bin/bash

# YSLv6Hub Folder Synchronization Script
# This script synchronizes the YSLv6Hub project files with Google Drive

# Configuration
SOURCE_DIR="/Users/galagrove/yslv6hub"
TARGET_DIR="/Users/galagrove/Library/CloudStorage/GoogleDrive-ssullivan@penbayymca.org/My Drive/SRS YSLv6Hub"

# Validate TypeScript compilation
echo "Validating TypeScript..."
cd "$SOURCE_DIR"
npx tsc --noEmit
if [ $? -ne 0 ]; then
  echo "TypeScript validation failed, aborting sync"
  exit 1
fi

# Sync files
echo "Syncing files from $SOURCE_DIR to $TARGET_DIR"
rsync -av --delete \
  --exclude='.git' \
  --exclude='node_modules' \
  --exclude='.clasp.json' \
  --exclude='build' \
  --exclude='tests' \
  "$SOURCE_DIR/" "$TARGET_DIR/"

# Create timestamp file
echo "Last synced: $(date)" > "$TARGET_DIR/LAST_UPDATED.txt"

echo "Sync completed successfully."
echo "Synced with Google Drive folder: $TARGET_DIR"