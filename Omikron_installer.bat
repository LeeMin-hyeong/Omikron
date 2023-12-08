@echo off
@chcp 65001 1> NUL 2> NUL

python --version
IF %ERRORLEVEL% NEQ 0 (
    echo 프로그램 설치를 진행하기 위해서 python이 필요합니다.
    pause
    exit
)
git --version
IF %ERRORLEVEL% NEQ 0 (
    echo 프로그램 설치를 진행하기 위해서 git이 필요합니다.
    pause
    exit
)

echo 필요 파일 다운로드
pip install upgrade pip
pip install --upgrade openpyxl selenium webdriver-manager pyinstaller pywin32 python-dateutil
git clone https://github.com/LeeMin-hyeong/Omikron.git
cd Omikron
pyinstaller -F omikrondb.py

echo 어플리케이션 빌드 완료
cd ..
move Omikron\dist\* .
rd /s /q Omikron

echo 프로그램 설치가 완료되었습니다.
pause