@echo off
CALL .\venv\Scripts\activate
cxfreeze -c -s -O --target-name=MyVol2_Server --include-files=server_data.json --target-dir=MyVol2_Server --icon=favicon.ico main.py
pause