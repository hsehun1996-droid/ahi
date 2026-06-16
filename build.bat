@echo off
chcp 65001 > /dev/null
echo ============================================================
echo  Build: Highway Maintenance Management EXE
echo ============================================================
echo.

python --version >/dev/null 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Install Python 3.11+ from https://www.python.org
    echo         Check "Add Python to PATH" during install.
    pause
    exit /b 1
)

echo [1/4] Installing packages...
pip install --upgrade pip --quiet
pip install customtkinter CTkMessagebox darkdetect packaging Pillow reportlab openpyxl pywin32 pyinstaller
if errorlevel 1 (
    echo [ERROR] Package install failed.
    pause
    exit /b 1
)

echo.
echo [2/4] Cleaning previous build...
if exist "dist" (
    for /d %%i in ("dist\*") do rmdir /s /q "%%i"
)
if exist "build" rmdir /s /q "build"

echo.
echo [3/4] Running PyInstaller...
pyinstaller highway.spec --noconfirm
if errorlevel 1 (
    echo [ERROR] Build failed. Check errors above.
    pause
    exit /b 1
)

echo.
echo [4/4] Creating ZIP archive...
for /d %%d in ("dist\*") do (
    powershell -Command "Compress-Archive -Path '%%d\*' -DestinationPath 'dist\highway_release.zip' -Force"
)

echo.
echo ============================================================
echo  BUILD COMPLETE
echo  EXE folder : dist\
echo  ZIP file   : dist\highway_release.zip
echo ============================================================
echo.
pause
