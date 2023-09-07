@echo off
echo Omikron Program Update Progress

mkdir temp

git clone -b v1.2.0 https://github.com/LeeMin-hyeong/Omikron.git temp
pip install --upgrade webdriver-manager
pip install --upgrade pyinstaller
pip install --upgrade openpyxl
pip install --upgrade selenium

cd temp
pyinstaller -F omikrondb.py
copy ./dist/omikrondb.exe ../
pause