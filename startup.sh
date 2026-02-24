#!/bin/bash
set -e
pip install -r requirements.txt -q
python flatfinder_scraper.py
echo ""
echo "✓ flatfinder_toronto.xlsx"
echo "✓ flatfinder_toronto_latest.csv"
