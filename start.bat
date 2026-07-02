@echo off
title Liberlive AI Station v26.0

echo.
echo  ========================================
echo   Liberlive AI Station v26.0  by Brett
echo  ========================================
echo.
echo  Starting server...
echo  Browser: http://localhost:8501
echo  Mobile (same WiFi): http://[your-PC-IP]:8501
echo.

cd /d "%~dp0"

pip install -r requirements.txt -q

streamlit run app.py --server.address=0.0.0.0 --server.port=8501

pause
