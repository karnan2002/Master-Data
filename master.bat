@echo off
cd /E E:\MASTERDATA

start /B cmd /C "python try.py"

REM Waiting for a short delay to ensure the Flask application is up before opening the browser
timeout /nobreak /t 3

start http://localhost:5002    

echo DON'T CLOSE CMD..!! PLEASE MINIMIZE SIR.

REM Pause the script so that the Command Prompt window remains open
pause
