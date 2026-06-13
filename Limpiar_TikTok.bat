@echo off
cd /d "%~dp0"
if exist ".venv\Scripts\pythonw.exe" (
    start "" ".venv\Scripts\pythonw.exe" clean_tiktok.py
) else (
    echo No se encontro el entorno .venv en esta carpeta.
    pause
)
