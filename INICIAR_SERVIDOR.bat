@echo off
REM ============================================================
REM   EXPO PYME - Iniciar servidor OFFLINE (doble clic)
REM   Levanta el servidor y abre el panel de control.
REM ============================================================
chcp 65001 >nul
cd /d "%~dp0"

REM Abre el panel en el navegador 3 segundos despues (cuando ya arranco).
start "" cmd /c "timeout /t 3 >nul & start http://localhost:8080/panel"

echo.
echo   Iniciando servidor Expo PyME en el puerto 8080...
echo   (El panel se abrira solo en el navegador)
echo.

REM Intenta 'python'; si no existe, prueba 'py'.
python servidor.py --port 8080
if errorlevel 1 py servidor.py --port 8080

echo.
echo   El servidor se detuvo. Presiona una tecla para cerrar.
pause >nul
