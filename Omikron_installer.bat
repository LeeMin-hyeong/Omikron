@echo off

echo.
echo Checking requirements...
echo.
python --version
IF %ERRORLEVEL% NEQ 0 (
    echo The installation of Omikron requires Python3.10 to be executed.
    pause
    exit
)
git --version
IF %ERRORLEVEL% NEQ 0 (
    echo The installation of Omikron requires git to be executed.
    pause
    exit
)

echo.
echo Upgrade python packages...
echo.
pip install --upgrade pip
pip install --upgrade openpyxl selenium webdriver-manager pyinstaller pywin32 python-dateutil
pip install -U pyinstaller

echo.
echo Fetching Source Code...
echo.
git clone https://github.com/LeeMin-hyeong/Omikron.git OmikronTemp

echo.
echo Build executable file...
echo.
cd OmikronTemp
C:\Users\%USERNAME%\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.10_qbz5n2kfra8p0\LocalCache\local-packages\Python310\Scripts\pyinstaller.exe -F --exclude numpy -n Omikron.exe omikron.py

cd ..
move OmikronTemp\dist\* .
rd /s /q OmikronTemp

echo.
echo Installation completed
pause