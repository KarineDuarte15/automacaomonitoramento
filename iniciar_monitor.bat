@echo off
title Monitoramento OptimusBI v6.0
cd /d "%~dp0"

echo [1/3] Liberando porta 8080 de processos antigos...
powershell -NoProfile -Command "Get-NetTCPConnection -LocalPort 8080 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess -Unique | ForEach-Object { Stop-Process -Id $_ -Force -ErrorAction SilentlyContinue }" >nul 2>&1

echo [2/3] Verificando bibliotecas Python...
pip install openpyxl selenium pyautogui pyperclip pywinauto >nul 2>&1

:loop
echo [%time%] Iniciando Monitoramento OptimusBI v6.0...
python monitor_automacoes.py

if %errorlevel%==10 (
  echo.
  echo [%time%] Monitoramento encerrado automaticamente apos o ultimo horario do dia.
  timeout /t 10 >nul
  exit /b 0
)

if %errorlevel%==3 (
  echo.
  echo Ja existe outro monitor rodando. Feche a outra janela e execute este arquivo novamente.
  pause
  exit /b
)

echo O monitor parou inesperadamente. Reiniciando em 30 segundos...
timeout /t 30 >nul
goto loop