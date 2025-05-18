#!/bin/bash
# YSLv6Hub Setup Script
# This script sets up the development environment for YSLv6Hub

# Exit on error
set -e

echo "Setting up YSLv6Hub development environment..."

# Check if Node.js is installed
if ! command -v node &> /dev/null; then
  echo "Node.js is required but not installed. Please install Node.js first."
  exit 1
fi

# Check if npm is installed
if ! command -v npm &> /dev/null; then
  echo "npm is required but not installed. Please install npm first."
  exit 1
fi

# Check if git is installed
if ! command -v git &> /dev/null; then
  echo "git is required but not installed. Please install git first."
  exit 1
fi

# Install dependencies
echo "Installing project dependencies..."
npm install

# Check if clasp is installed globally
if ! command -v clasp &> /dev/null; then
  echo "Installing clasp globally..."
  npm install -g @google/clasp
fi

# Login to clasp if not already logged in
echo "Checking clasp login status..."
clasp login

# Create necessary directories if they don't exist
echo "Creating project directory structure..."
mkdir -p build src/utils src/interfaces src/models src/services src/ui
mkdir -p tests/mocks tests/utils tests/services
mkdir -p types docs

# Make sync script executable
chmod +x tools/sync-ysl-folders.sh

# Verify TypeScript compilation
echo "Verifying TypeScript compilation..."
npm run typecheck

# Run linting
echo "Running linting..."
npm run lint

# Run tests
echo "Running tests..."
npm test

echo "Setup complete. You're ready to start developing YSLv6Hub!"
echo "Run 'npm run build' to compile TypeScript."
echo "Run 'npm run deploy' to deploy to Google Apps Script."