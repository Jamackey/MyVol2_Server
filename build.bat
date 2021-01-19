@echo off
CALL .\venv\Scripts\activate
cxfreeze -c -O --target-name=MyVol2_Server --target-dir=MyVol2_Server --icon=favicon.ico main.py
pause