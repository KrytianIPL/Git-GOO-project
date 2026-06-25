@echo off
REM ============================================================
REM  Generador de Dashboard de Reabastecimiento - Cliente GOO
REM  Doble clic para ejecutar.
REM ============================================================
cd /d "%~dp0"
echo Iniciando generador...
echo.

REM Intenta con "python"; si no, prueba con "py"
python generar_dashboard.py
if %errorlevel% neq 0 (
    py generar_dashboard.py
)

REM Si ambos fallan, avisa
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] No se pudo ejecutar Python.
    echo Verifica que Python este instalado y agregado al PATH.
    echo.
    pause
)
