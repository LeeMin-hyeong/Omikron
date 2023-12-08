@echo off

echo.
echo Checking requirements...
echo.
python --version
IF %ERRORLEVEL% NEQ 0 (
    echo The installation of this program requires Python3.10+ to be executed.
    pause
    exit
)
git --version
IF %ERRORLEVEL% NEQ 0 (
    echo The installation of this program requires git to be executed.
    pause
    exit
)

echo.
echo Upgrade python packages...
echo.
pip install --upgrade pip
pip install --upgrade openpyxl selenium webdriver-manager pyinstaller pywin32 python-dateutil

echo.
echo Fetching Source Code...
echo.
git clone https://github.com/LeeMin-hyeong/Omikron.git

echo.
echo Build executable file...
echo.
cd Omikron
C:\Users\%USERNAME%\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.10_qbz5n2kfra8p0\LocalCache\local-packages\Python310\Scripts\pyinstaller.exe -F omikrondb.py

cd ..
move Omikron\dist\* .
rd /s /q Omikron

echo.
echo Installation completed
pause