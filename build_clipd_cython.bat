@echo off
setlocal enabledelayedexpansion

rem Determine which python executable to use
set "PYTHON_CMD="
if not "%~1"=="" (
    set "PYTHON_CMD=%~1"
) else if defined PYTHON_EXE (
    set "PYTHON_CMD=%PYTHON_EXE%"
) else (
    set "PYTHON_CMD=python"
)

%PYTHON_CMD% -c "import sys" >nul 2>nul
if errorlevel 1 (
    py -c "import sys" >nul 2>nul
    if errorlevel 1 (
        echo Konnte Python Interpreter nicht finden. Bitte Pfad per Parameter oder Variablen PYTHON_EXE angeben.
        exit /b 1
    )
    set "PYTHON_CMD=py"
)

for /f "delims=" %%I in ('%PYTHON_CMD% -c "import sys;print(sys.executable)"') do set "PYTHON_PATH=%%~I"
echo Verwende Python Interpreter: %PYTHON_PATH%

set "SCRIPT_DIR=%~dp0"
set "SCRIPT_DIR_NOSEP=%SCRIPT_DIR:~0,-1%"
set "SCRIPT_PATH=%SCRIPT_DIR%clipboard_guardian.py"
set "BUILD_DIR=%SCRIPT_DIR%build_cython"

if not exist "%SCRIPT_PATH%" (
    echo clipboard_guardian.py nicht gefunden.
    exit /b 1
)

echo Cleaning previous build artifacts...
if exist "%BUILD_DIR%" rd /s /q "%BUILD_DIR%"
if exist "%SCRIPT_DIR%__pycache__" rd /s /q "%SCRIPT_DIR%__pycache__"
del /q "%SCRIPT_DIR%clipd_core.cp*.pyd" 2>nul

echo Building ClipD core with Cython...
"%PYTHON_CMD%" setup_cython.py build_ext --inplace --build-temp build_cython
if errorlevel 1 (
    echo Cython build failed.
    exit /b %errorlevel%
)

echo Building ClipD executable mit PyInstaller...

"%PYTHON_CMD%" -m PyInstaller ^
    --name ClipD ^
    --onefile ^
    --noconsole ^
    --icon "%SCRIPT_DIR%logo.ico" ^
    --additional-hooks-dir "%SCRIPT_DIR_NOSEP%" ^
    --add-data "%SCRIPT_DIR%logo.png;." ^
    --add-data "%SCRIPT_DIR%logo.ico;." ^
    --hidden-import PySide6.QtSvg ^
    --hidden-import PySide6.QtGui ^
    --hidden-import PySide6.QtWidgets ^
    --hidden-import PySide6.QtCore ^
    --clean ^
    --distpath distCython ^
    --workpath build_temp ^
    run_clipd.py

if errorlevel 1 (
    echo Build failed.
    exit /b %errorlevel%
)

echo Build completed. Output befindet sich im dist-Verzeichnis.
endlocal
exit /b 0
