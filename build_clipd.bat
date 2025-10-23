@echo off
setlocal enabledelayedexpansion

set SCRIPT_DIR=%~dp0
set SCRIPT_PATH=%SCRIPT_DIR%clipboard_guardian.py

if not exist "%SCRIPT_PATH%" (
    echo clipboard_guardian.py nicht gefunden.
    exit /b 1
)

echo Building ClipD with Nuitka...

set USE_CLANG=
set CLANG_PATH=C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Tools\Llvm\x64\bin\clang.exe
if exist "%CLANG_PATH%" (
    set USE_CLANG=--clang
) else (
    echo LLVM/Clang nicht gefunden. Verwende MinGW64 Toolchain.
    set USE_CLANG=--mingw64
)

python -m nuitka ^
    --onefile ^
    --standalone ^
    %USE_CLANG% ^
    --enable-plugin=pyside6 ^
    --windows-console-mode=disable ^
    --windows-icon-from-ico=logo.ico ^
    --include-data-files=logo.png=logo.png ^
    --include-data-files=logo.ico=logo.ico ^
    --output-filename=ClipD ^
    --output-dir=dist ^
    clipboard_guardian.py

if errorlevel 1 (
    echo Build failed.
    exit /b %errorlevel%
)

echo Build completed. Output befindet sich im dist-Verzeichnis.
endlocal
