@echo off
chcp 65001 >nul
echo ============================================
echo  고속도로 포장유지보수 이력 관리 프로그램
echo  EXE 빌드 스크립트
echo ============================================
echo.

cd /d "%~dp0"

:: pyinstaller 설치 확인
python -m pyinstaller --version >nul 2>&1
if errorlevel 1 (
    echo [오류] PyInstaller가 설치되어 있지 않습니다.
    echo 다음 명령으로 설치하세요: pip install pyinstaller
    pause
    exit /b 1
)

echo [1/3] 이전 빌드 정리 중...
if exist "dist\고속도로_포장유지보수_관리" rmdir /s /q "dist\고속도로_포장유지보수_관리"
if exist "build\highway" rmdir /s /q "build\highway"

echo [2/3] 빌드 시작...
python -m pyinstaller highway.spec --noconfirm

if errorlevel 1 (
    echo.
    echo [오류] 빌드 실패. 위 오류 메시지를 확인하세요.
    pause
    exit /b 1
)

echo.
echo [3/3] 빌드 완료!
echo.
echo 배포 폴더: dist\고속도로_포장유지보수_관리\
echo 실행 파일: dist\고속도로_포장유지보수_관리\고속도로_포장유지보수_관리.exe
echo.
echo ※ 다른 PC에 배포할 때는 dist\고속도로_포장유지보수_관리\ 폴더 전체를 복사하세요.
echo ※ CSV 세이브 파일이 폴더 안에 함께 포함되어 있습니다.
echo.
pause
