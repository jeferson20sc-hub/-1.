@echo off
REM ===========================================================================
REM build_all.bat - Atalho Windows para gerar EXAUSTAO_360_ENTERPRISE_PRO.xlsm
REM ===========================================================================
setlocal

echo.
echo === EXAUSTAO 360 ENTERPRISE PRO - build_all ===
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [ERR] Python nao encontrado. Instale Python 3.10+ e adicione ao PATH.
    exit /b 1
)

echo [1/3] Instalando dependencias Python...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERR] Falha ao instalar dependencias.
    exit /b 1
)

echo.
echo [2/3] Gerando esqueleto .xlsx...
python build.py
if errorlevel 1 (
    echo [ERR] Falha em build.py.
    exit /b 1
)

echo.
echo [3/3] Injetando VBA e salvando .xlsm...
powershell -ExecutionPolicy Bypass -File assemble.ps1
if errorlevel 1 (
    echo [ERR] Falha em assemble.ps1.
    exit /b 1
)

echo.
echo === Build concluido. Arquivo final em dist\EXAUSTAO_360_ENTERPRISE_PRO.xlsm ===
endlocal
