#!/bin/bash
set -e
pip install -r requirements.txt -q
# Map API_KEY to ANTHROPIC_API_KEY if not already set
export ANTHROPIC_API_KEY="${ANTHROPIC_API_KEY:-$API_KEY}"
python flatfinder_scraper.py
echo ""
echo "✓ flatfinder_toronto.xlsx"
echo "✓ flatfinder_toronto_latest.csv"
