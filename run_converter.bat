@echo off
echo Installing required dependencies...
pip install -r requirements.txt
echo.
echo Starting PDF to Word Converter...
python pdf_to_word.py
pause
