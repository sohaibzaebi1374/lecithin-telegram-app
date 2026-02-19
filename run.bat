@echo off
cd /d %~dp0

echo Installing dependencies...
py -m pip install --upgrade pip
py -m pip install -r requirements.txt

echo Starting bot...
py bot.py

pause
