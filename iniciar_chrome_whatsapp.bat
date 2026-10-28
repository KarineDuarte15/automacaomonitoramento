@echo off
REM Abre um Edge exclusivo do monitor. Faca login no WhatsApp Web UMA vez (QR code) e deixe este Edge aberto.
start "" "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222 --user-data-dir="C:\edge_whatsapp_monitor" https://web.whatsapp.com
