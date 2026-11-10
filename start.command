#!/usr/bin/env bash
cd "$(dirname "$0")"

echo "=================================================="
echo "    Table to Sheets AI - แปลงภาพตารางเป็น Google Sheets"
echo "=================================================="

# Check venv
if [ -d "../.venv" ]; then
    PYTHON="../.venv/bin/python"
elif [ -d ".venv" ]; then
    PYTHON=".venv/bin/python"
else
    PYTHON="python3"
fi

# Ensure output directory exists and initial files are generated
$PYTHON export_data.py >/dev/null 2>&1

echo "กำลังเปิดเบราว์เซอร์ที่ http://127.0.0.1:8765 ..."
sleep 1 && open "http://127.0.0.1:8765" &

# Start FastAPI server
$PYTHON -m uvicorn server:app --host 127.0.0.1 --port 8765
