@echo off
title Monitoramento OptimusBI
cd /d "%~dp0"
echo Verificando bibliotecas Python (openpyxl e selenium)...
pip install openpyxl selenium >nul 2>&1
echo Bibliotecas instaladas.
:loop
echo [%time%] Iniciando Monitoramento OptimusBI...
python monitor_automacoes.py
echo O monitor parou. Reiniciando em 30 segundos...
timeout /t 30 >nul
goto loop