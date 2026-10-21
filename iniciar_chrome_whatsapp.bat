@echo off
REM Abre um Chrome exclusivo do monitor. Faca login no WhatsApp Web UMA vez (QR code) e deixe este Chrome aberto.
start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\chrome_whatsapp_monitor" https://web.whatsapp.com
