#!/usr/bin/env bash
# Quick launcher script for E-Commerce RFM Segmentation Dashboard

cd "$(dirname "$0")"

echo "======================================================"
echo " Starting Customer Lifecycle & RFM Segmentation Dashboard"
echo "======================================================"

if [ -f "./venv/bin/python" ]; then
    ./venv/bin/python run_dashboard.py
elif command -v python3 &>/dev/null; then
    python3 run_dashboard.py
else
    python run_dashboard.py
fi
