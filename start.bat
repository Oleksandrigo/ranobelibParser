@echo off

call "%~dp0.venv\Scripts\activate"

cd "%~dp0"

python main.py

pause