@echo off
title Attendance Report Server
cd C:\Users\91876\Desktop\attendance_web
call venv\Scripts\activate
echo.
echo =========================================
echo   Attendance Report Server Starting...
echo =========================================
echo.
echo Server running at: http://127.0.0.1:10000
echo.
echo Press Ctrl+C to stop the server
echo.
python app.py
pause