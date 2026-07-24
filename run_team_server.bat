@echo off
setlocal
cd /d "%~dp0"

echo Starting MSDS Translation App for team access...
echo.
echo Local URL   : http://localhost:8501
echo Network URL : http://<this-pc-ip>:8501
echo.
echo Password required in app: 5050
echo Press Ctrl+C to stop server.
echo.

streamlit run "번역앱.py" --server.address 0.0.0.0 --server.port 8501 --server.headless true
