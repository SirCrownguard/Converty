@echo off
chcp 65001 >nul
cls
cd /d %~dp0
python pdf_to_pptx.py
echo Kapatmak için herhangi bir tuşa basın...
pause >nul
