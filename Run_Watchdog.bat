@echo off
cd /d %~dp0
call .venv\Scripts\activate.bat
python Synapsen_Normalisierer\Synapsen_Watchdog.py
pause