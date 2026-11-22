@echo off
chcp 65001 > nul
set "PYTHONIOENCODING=utf-8"

if not defined WEASYPRINT_DLL_DIRECTORIES (
    if exist "C:\Program Files\GTK3-Runtime Win64\bin\libgobject-2.0-0.dll" (
        set "WEASYPRINT_DLL_DIRECTORIES=C:\Program Files\GTK3-Runtime Win64\bin"
    ) else if exist "C:\msys64\mingw64\bin\libpango-1.0-0.dll" (
        set "WEASYPRINT_DLL_DIRECTORIES=C:\msys64\mingw64\bin"
    )
)

if not exist ".venv\Scripts\python.exe" (
    echo [ERROR] No se encontro el entorno virtual .venv. Ejecuta primero instalar_windows.bat.
    pause
    exit /b 1
)

set "VOL=%~1"
set "NUM=%~2"
set "YEAR=%~3"

if "%VOL%"=="" set /p "VOL=Volumen (Enter para 19): "
if "%VOL%"=="" set "VOL=19"

if "%NUM%"=="" set /p "NUM=Numero (Enter para 3): "
if "%NUM%"=="" set "NUM=3"

if "%YEAR%"=="" set /p "YEAR=Anio (Enter para 2026): "
if "%YEAR%"=="" set "YEAR=2026"

set "OUT_PDF=tapa_v%VOL%n%NUM%.pdf"
set "OUT_PNG=tapa_v%VOL%n%NUM%.png"

echo [1/3] Generando trama SVG ^(Vol. %VOL%, Num. %NUM%, %YEAR%^)...
.venv\Scripts\python.exe generar_trama.py --vol %VOL% --num %NUM% --year %YEAR% --out trama.svg
if errorlevel 1 (
    echo [ERROR] Fallo la generacion de trama.svg.
    pause
    exit /b 1
)

echo [2/3] Renderizando PDF de tapa ^(%OUT_PDF%^)...
.venv\Scripts\python.exe -m weasyprint tapa.html "%OUT_PDF%" --presentational-hints
if errorlevel 1 (
    echo [ERROR] Fallo la generacion de %OUT_PDF% con WeasyPrint.
    pause
    exit /b 1
)

echo [3/3] Generando vista previa PNG ^(%OUT_PNG%^)...
.venv\Scripts\python.exe -c "import fitz; doc=fitz.open('%OUT_PDF%'); doc[0].get_pixmap(dpi=150).save('%OUT_PNG%')" > nul 2>&1

echo.
echo [OK] Tapa generada: %OUT_PDF%
pause
exit /b 0
