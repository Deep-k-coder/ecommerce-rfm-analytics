"""
run_dashboard.py
Single executable launcher for the E-Commerce RFM Customer Lifecycle & Segmentation Dashboard.

Usage:
    python3 run_dashboard.py
    # or
    ./venv/bin/python run_dashboard.py
"""

import sys
import os
import time
import socket
import webbrowser
import threading
from pathlib import Path
from http.server import HTTPServer, SimpleHTTPRequestHandler

PROJECT_ROOT = Path(__file__).resolve().parent
DASHBOARD_DIR = PROJECT_ROOT / "dashboard"
INDEX_HTML = DASHBOARD_DIR / "index.html"
SUMMARY_JSON = DASHBOARD_DIR / "summary_data.json"
DIM_CUST_CSV = PROJECT_ROOT / "data" / "processed" / "dim_customer_rfm.csv"


def ensure_pipeline_executed():
    """Ensures data and dashboard files exist; runs pipeline if missing."""
    if not (INDEX_HTML.exists() and SUMMARY_JSON.exists() and DIM_CUST_CSV.exists()):
        print("\n[INFO] Analytics artifacts missing. Executing ML pipeline & dashboard build...")
        import main
        main.run_pipeline()
        import build_dashboard
        build_dashboard.build_html()
        print("[SUCCESS] Pipeline initialized.\n")


def find_free_port(start_port=8000, max_attempts=50):
    """Finds an open TCP port starting from start_port."""
    for port in range(start_port, start_port + max_attempts):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", port)) != 0:
                return port
    return 8000


class CustomDashboardHandler(SimpleHTTPRequestHandler):
    """Serves the dashboard directory cleanly."""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(DASHBOARD_DIR), **kwargs)

    def log_message(self, format, *args):
        # Keep terminal clean; suppress standard GET/POST request noise
        pass


def launch_server():
    ensure_pipeline_executed()

    port = find_free_port(8000)
    server_address = ("127.0.0.1", port)
    url = f"http://127.0.0.1:{port}"

    server = HTTPServer(server_address, CustomDashboardHandler)

    banner = f"""
======================================================================
  ⚡ E-COMMERCE RFM CUSTOMER SEGMENTATION DASHBOARD ⚡
======================================================================
  • Student Candidate   : Mr. Deep Koshiya (B.Sc. AI & Data Science)
  • Academic Program    : Semester V (5th Semester) | Project - 1 (CE-502)
  • Machine Learning    : Vectorized RFM + K-Means Clustering + Quantile CRM
  • Local Server URL    : {url}
  • Direct File Path    : file://{INDEX_HTML}
======================================================================
  [STATUS] Local server active! Opening dashboard in your default browser...
  [HINT]   Press Ctrl+C in this terminal to stop the server anytime.
======================================================================
"""
    print(banner)

    # Open browser automatically after a short delay
    def open_browser():
        time.sleep(1.0)
        webbrowser.open(url)

    browser_thread = threading.Thread(target=open_browser, daemon=True)
    browser_thread.start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n[SHUTDOWN] Dashboard server stopped gracefully. Have a great day!\n")
        server.server_close()
        sys.exit(0)


if __name__ == "__main__":
    launch_server()
