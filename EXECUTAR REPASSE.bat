@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Instalando/verificando bibliotecas (so demora na primeira vez)...
python -m pip install -q -r requirements.txt
python main.py
echo.
pause
