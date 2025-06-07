#!/bin/bash
cd "$(dirname "$0")"
echo "🔄 Running FilteringDividendChampionsExcel.py..."
python3 FilteringDividendChampionsExcel.py
echo "✅ Done. Press any key to close."
read -n 1
