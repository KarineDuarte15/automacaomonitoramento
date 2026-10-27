iniciar_monitor (2).bat
 
@echo off

title Monitoramento OptimusBI

cd /d "%~dp0"

echo Verificando bibliotecas Python (openpyxl e selenium)...

pip install openpyxl selenium >nul 2>&1

:loop

echo [%time%] Iniciando Monitoramento OptimusBI...

python monitor_automacoes.py

if %errorlevel%==3 (

  echo.

  echo Ja existe outro monitor rodando. Feche a outra janela e execute este arquivo novamente.

  pause

  exit /b

)

echo O monitor parou. Reiniciando em 30 segundos...

timeout /t 30 >nul

goto loop
 