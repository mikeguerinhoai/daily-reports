#!/usr/bin/env bash
# Deploy the latest daily dashboard HTML to GitHub Pages (gh-pages branch).
# Uses git worktree to avoid switching branches in the main working directory.
#
# Usage:
#   bash daily-reports/deploy-to-pages.sh --date 2026-05-10
#   bash daily-reports/deploy-to-pages.sh                     # finds latest

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_DIR="$SCRIPT_DIR"
OUTPUT_DIR="$SCRIPT_DIR/output"
WORKTREE_DIR="$SCRIPT_DIR/.gh-pages-worktree"

# Parse --date arg
DATE=""
while [[ $# -gt 0 ]]; do
  case $1 in
    --date) DATE="$2"; shift 2 ;;
    *) DATE="$1"; shift ;;
  esac
done

# Find the dashboard HTML
if [[ -n "$DATE" ]]; then
  HTML_FILE="$OUTPUT_DIR/HOAi_Daily_Dashboard_${DATE}.html"
else
  # Find most recent dashboard HTML
  HTML_FILE=$(ls -t "$OUTPUT_DIR"/HOAi_Daily_Dashboard_*.html 2>/dev/null | head -1)
  if [[ -n "$HTML_FILE" ]]; then
    DATE=$(echo "$HTML_FILE" | grep -oP '\d{4}-\d{2}-\d{2}')
  fi
fi

if [[ ! -f "$HTML_FILE" ]]; then
  echo "ERROR: Dashboard HTML not found: $HTML_FILE"
  echo "Run generate-daily-dashboard.js first."
  exit 1
fi

echo "Deploying $HTML_FILE to gh-pages..."

# Clean up any leftover worktree
if [[ -d "$WORKTREE_DIR" ]]; then
  cd "$REPO_DIR"
  git worktree remove "$WORKTREE_DIR" --force 2>/dev/null || rm -rf "$WORKTREE_DIR"
fi

# Create worktree for gh-pages
cd "$REPO_DIR"
git worktree add "$WORKTREE_DIR" gh-pages

# Copy dashboard as both latest.html and dated archive
cp "$HTML_FILE" "$WORKTREE_DIR/latest.html"
cp "$HTML_FILE" "$WORKTREE_DIR/HOAi_Daily_Dashboard_${DATE}.html"

# Commit and push
cd "$WORKTREE_DIR"
git add latest.html "HOAi_Daily_Dashboard_${DATE}.html"

if git diff --cached --quiet; then
  echo "No changes to deploy."
else
  git commit -m "deploy: dashboard ${DATE}"
  git push origin gh-pages
  echo "Deployed to gh-pages: latest.html + HOAi_Daily_Dashboard_${DATE}.html"
fi

# Clean up worktree
cd "$REPO_DIR"
git worktree remove "$WORKTREE_DIR" --force

echo "Done. URL: https://mikeguerinhoai.github.io/daily-reports/latest.html"
