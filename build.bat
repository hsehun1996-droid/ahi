@echo off
REM 고속도로 포장유지보수 이력 관리 프로그램 — Windows EXE 빌드 스크립트
REM 사용법: build.bat 을 프로젝트 루트에서 더블클릭하거나 cmd에서 실행

echo ============================================================
echo  고속도로 포장유지보수 이력 관리 프로그램 EXE 빌드
echo ============================================================
echo.

REM Python 설치 확인
python --version >/dev/null 2>&1
if errorlevel 1 (
    echo [오류] Python이 설치되지 않았거나 PATH에 없습니다.
    echo        https://www.python.org 에서 Python 3.11+ 을 설치하세요.
    pause
    exit /b 1
)

echo [1/4] 의존성 패키지 설치 중...
pip install --upgrade pip >/dev/null
pip install customtkinter CTkMessagebox darkdetect packaging Pillow reportlab openpyxl pywin32 pyinstaller
if errorlevel 1 (
    echo [오류] 패키지 설치 실패
    pause
    exit /b 1
)

echo.
echo [2/4] 이전 빌드 결과 삭제 중...
if exist "dist\고속도로_포장유지보수_관리" rmdir /s /q "dist\고속도로_포장유지보수_관리"
if exist build rmdir /s /q build

echo.
echo [3/4] PyInstaller 빌드 시작...
pyinstaller highway.spec --noconfirm
if errorlevel 1 (
    echo [오류] 빌드 실패. 위 오류 메시지를 확인하세요.
    pause
    exit /b 1
)

echo.
echo [4/4] ZIP 압축 중...
powershell -Command "Compress-Archive -Path 'dist\고속도로_포장유지보수_관리\*' -DestinationPath 'dist\고속도로_포장유지보수_관리.zip' -Force"

echo.
echo ============================================================
echo  빌드 완료!
echo  실행 파일: dist\고속도로_포장유지보수_관리\고속도로_포장유지보수_관리.exe
echo  ZIP 파일:  dist\고속도로_포장유지보수_관리.zip
echo ============================================================
echo.
pause
