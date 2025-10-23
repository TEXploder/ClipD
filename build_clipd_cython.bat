@echo off
setlocal enabledelayedexpansion

set SCRIPT_DIR=%~dp0
set SCRIPT_PATH=%SCRIPT_DIR%clipboard_guardian.py
set BUILD_DIR=%SCRIPT_DIR%build_cython

if not exist "%SCRIPT_PATH%" (
    echo clipboard_guardian.py nicht gefunden.
    exit /b 1
)

echo Cleaning previous build artifacts...
if exist "%BUILD_DIR%" rd /s /q "%BUILD_DIR%"
if exist "%SCRIPT_DIR%__pycache__" rd /s /q "%SCRIPT_DIR%__pycache__"

echo Building ClipD core with Cython...
python setup_cython.py build_ext --inplace --build-temp build_cython
if errorlevel 1 (
    echo Cython build failed.
    exit /b %errorlevel%
)

echo Building ClipD executable with PyInstaller...

pyinstaller ^
    --name ClipD ^
    --onefile ^
    --noconsole ^
    --icon "%SCRIPT_DIR%logo.ico" ^
    --add-data "%SCRIPT_DIR%logo.png;." ^
    --add-data "%SCRIPT_DIR%logo.ico;." ^
    --hidden-import PySide6.QtSvg ^
    --hidden-import PySide6.QtGui ^
    --hidden-import PySide6.QtWidgets ^
    --hidden-import PySide6.QtCore ^
    --clean ^
    --distpath dist ^
    --workpath build_temp ^
    run_clipd.py

if errorlevel 1 (
    echo Build failed.
    exit /b %errorlevel%
)

echo Build completed. Output befindet sich im dist-Verzeichnis.
endlocal
