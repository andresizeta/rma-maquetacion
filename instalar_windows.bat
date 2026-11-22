@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"

echo [1/4] Verificando Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python no esta instalado o no figura en el PATH.
    pause
    exit /b 1
)

echo [2/4] Creando entorno virtual .venv...
python -m venv .venv
if %errorlevel% neq 0 goto :error

echo [3/4] Instalando dependencias en .venv...
.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install -r requirements.txt
if %errorlevel% neq 0 goto :error

echo [4/4] Verificando motor grafico de WeasyPrint (GTK3/Pango)...
.venv\Scripts\python.exe -c "import weasyprint" >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo [AVISO] Python y los paquetes se instalaron, pero falta el runtime GTK3/Pango en Windows.
    echo Descarga e instala GTK3 Runtime para Windows y vuelve a ejecutar este script.
    pause
    exit /b 1
)

echo.
echo [OK] Instalacion completada correctamente.
echo Para probar el articulo 4812, ejecuta:
echo   .venv\Scripts\python maquetar_rma.py ejemplo_ojs 4812
echo Para probar la paginacion consecutiva de un fasciculo completo:
echo   .venv\Scripts\python maquetar_fasciculo.py ejemplo_fasciculo
pause
exit /b 0

:error
echo.
echo [ERROR] Se produjo un fallo durante la instalacion.
pause
exit /b 1