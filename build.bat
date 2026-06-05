@echo off
chcp 65001 >nul
echo ============================================
echo  고속도로 포장유지보수 이력 관리 프로그램
echo  EXE 빌드 스크립트
echo ============================================
echo.

cd /d "%~dp0"

:: 0) 파이썬 확인
python --version >nul 2>&1
if errorlevel 1 (
    echo [오류] Python이 설치되어 있지 않거나 PATH에 없습니다.
    echo https://www.python.org 에서 Python 3.10+ 를 설치하세요.
    pause
    exit /b 1
)

:: 1) 의존성 설치 (PyInstaller 포함)
echo [1/4] 의존성 설치/확인 중... (requirements.txt)
python -m pip install --upgrade pip >nul 2>&1
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [오류] 의존성 설치 실패. 인터넷 연결 또는 pip 설정을 확인하세요.
    pause
    exit /b 1
)

:: 2) 이전 빌드 정리
echo [2/4] 이전 빌드 정리 중...
if exist "dist\고속도로_포장유지보수_관리" rmdir /s /q "dist\고속도로_포장유지보수_관리"
if exist "build\고속도로_포장유지보수_관리" rmdir /s /q "build\고속도로_포장유지보수_관리"
if exist "build\highway" rmdir /s /q "build\highway"

:: 3) 빌드
echo [3/4] 빌드 시작...
python -m pyinstaller highway.spec --noconfirm
if errorlevel 1 (
    echo.
    echo [오류] 빌드 실패. 위 오류 메시지를 확인하세요.
    pause
    exit /b 1
)

echo.
echo [4/4] 빌드 완료!
echo.
echo 배포 폴더: dist\고속도로_포장유지보수_관리\
echo 실행 파일: dist\고속도로_포장유지보수_관리\고속도로_포장유지보수_관리.exe
echo.
echo ※ 다른 PC에 배포할 때는 dist\고속도로_포장유지보수_관리\ 폴더 전체를 복사하세요.
echo ※ PDF·Excel·한글 내보내기, 한글 폰트, 모식도 아이콘이 모두 포함됩니다.
echo ※ 한글(HWP) 양식 자동 작성은 실행 PC에 '한글' 프로그램이 설치된 경우에만 동작합니다.
echo.
pause
