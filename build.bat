@echo off
echo ============================================================
echo  Build: Highway Maintenance Management EXE
echo ============================================================
echo.

python --version 2>nul
if errorlevel 1 (
    echo [ERROR] Python not found.
    echo         Install Python 3.11+ from https://www.python.org
    echo         Check "Add Python to PATH" during install.
    pause
    exit /b 1
)

echo [1/4] Installing packages...
pip install --quiet customtkinter CTkMessagebox darkdetect packaging Pillow reportlab openpyxl pywin32 pyinstaller
if errorlevel 1 (
    echo [ERROR] Package install failed.
    pause
    exit /b 1
)

echo [2/4] Cleaning previous build...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [3/4] Running PyInstaller...
pyinstaller highway.spec --noconfirm
if errorlevel 1 (
    echo [ERROR] Build failed.
    pause
    exit /b 1
)

echo [4/4] Creating ZIP...
powershell -Command "Get-ChildItem dist | ForEach-Object { Compress-Archive -Path ($_.FullName + '\*') -DestinationPath ('dist\' + $_.Name + '.zip') -Force }"

echo.
echo ============================================================
echo  BUILD COMPLETE!
echo  EXE folder: dist\
echo ============================================================
echo.
pause
