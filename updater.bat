@echo off
@chcp 65001 1> NUL 2> NUL

python --version
IF %ERRORLEVEL% NEQ 0 (
    echo 수동 업데이트를 진행하기 위해서 python이 필요합니다.
    pause
    exit
)
git --version
IF %ERRORLEVEL% NEQ 0 (
    echo 수동 업데이트를 진행하기 위해서 git이 필요합니다.
    pause
    exit
)

pip install --upgrade openpyxl selenium webdriver-manager pyinstaller pywin32 python-dateutil
git clone https://github.com/LeeMin-hyeong/Omikron.git
cd Omikron
pyinstaller -F omikrondb.py

cd ..
move Omikron\dist\* .
rd /s /q Omikron

echo 수동 업데이트가 완료되었습니다.
pause