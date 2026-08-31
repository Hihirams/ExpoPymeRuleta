@echo off
REM ============================================================
REM   Expo PyME - Configurar red WiFi local (Opcion 2: Loopback)
REM   Doble clic. Aparecera un aviso de permisos: elige "Si".
REM ============================================================
chcp 65001 >nul
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0CONFIGURAR_LOOPBACK.ps1"
