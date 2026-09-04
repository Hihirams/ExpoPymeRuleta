@echo off
REM ============================================================
REM   EXPO PYME - Limpiar todos los datos (dejar todo en 0)
REM   Borra TODAS las respuestas (encuestas contestadas), todos
REM   los premios entregados, el cupo por hora y la configuracion
REM   (premios, PIN, "Premios por hora"...) restableciendola a
REM   sus valores por defecto. Util para hacer varias pruebas.
REM
REM   Solo limpia el SERVIDOR de la laptop (respuestas.db).
REM   Lo guardado en los CELULARES se borra por dispositivo:
REM   engrane -> "Borrar historico", o ventana de incognito.
REM ============================================================
chcp 65001 >nul
cd /d "%~dp0"

echo.
echo   ==========================================
echo     EXPO PYME - Limpiar todos los datos
echo   ==========================================
echo   Se borraran TODAS las respuestas, premios
echo   entregados y la configuracion quedara en 0.
echo.

REM Intenta 'python'; si no existe, prueba 'py'.
python limpiar_datos.py
if errorlevel 1 py limpiar_datos.py

echo.
echo   Proceso terminado. Todo quedo en cero.
pause >nul