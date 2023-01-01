@echo off

call "%~dp0venv\Scripts\activate"

cd "%~dp0"

python main.py

pause