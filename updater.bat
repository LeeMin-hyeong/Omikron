@echo off
REM Step 1: Git 및 Python 3.10 설치
echo Step 1: Installing Git and Python 3.10...
choco install git -y
IF %ERRORLEVEL% NEQ 0 exit /b %ERRORLEVEL%
choco install python3 -y
IF %ERRORLEVEL% NEQ 0 exit /b %ERRORLEVEL%
python --version

REM Step 2: 필요한 Python 패키지 설치
echo Step 2: Installing required Python packages...
pip install --upgrade openpyxl selenium webdriver-manager pyinstaller pywin32 python-dateutil
IF %ERRORLEVEL% NEQ 0 exit /b %ERRORLEVEL%

REM Step 3: Omikron 레포지토리 클론 및 빌드
echo Step 3: Cloning Omikron repository and building omikrondb.py...
git clone https://github.com/LeeMin-hyeong/Omikron.git
IF %ERRORLEVEL% NEQ 0 exit /b %ERRORLEVEL%
cd Omikron
pyinstaller omikrondb.py
IF %ERRORLEVEL% NEQ 0 exit /b %ERRORLEVEL%

REM 이동 및 삭제 작업 추가
cd ..
move Omikron\dist\* .
rd /s /q Omikron

REM 결과물 확인
echo.
echo Build completed. Check the current directory for the executable.
pause