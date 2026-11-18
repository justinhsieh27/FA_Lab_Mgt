@echo off
setlocal EnableExtensions

set "SCRIPT_DIR=%~dp0"
set "PID_FILE=%SCRIPT_DIR%fa-lab.pid"
set "LOG_DIR=%SCRIPT_DIR%logs"
set "LOG_FILE=%LOG_DIR%\fa-lab.log"
set "ERROR_LOG_FILE=%LOG_DIR%\fa-lab.error.log"

if not defined PORT set "PORT=8080"
cd /d "%SCRIPT_DIR%"

if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

if exist "%PID_FILE%" (
  set /p EXISTING_PID=<"%PID_FILE%"
  tasklist /FI "PID eq %EXISTING_PID%" /NH | findstr /I "%EXISTING_PID%" >nul
  if not errorlevel 1 (
    echo FA Lab Case Management is already running ^(PID %EXISTING_PID%^).
    echo Open: http://localhost:%PORT%
    exit /b 0
  )
  del /q "%PID_FILE%"
)

netstat -ano | findstr /R /C:":%PORT% .*LISTENING" >nul
if not errorlevel 1 (
  echo Port %PORT% is already in use. Stop the process using that port or choose another port.
  echo   set PORT=4001 ^&^& run.bat
  exit /b 1
)

for /f %%P in ('powershell -NoProfile -Command "$env:PORT = '%PORT%'; $process = Start-Process -FilePath 'npm.cmd' -ArgumentList 'run start' -WorkingDirectory '%SCRIPT_DIR%' -RedirectStandardOutput '%LOG_FILE%' -RedirectStandardError '%ERROR_LOG_FILE%' -PassThru -WindowStyle Hidden; $process.Id"') do set "SERVER_PID=%%P"

if not defined SERVER_PID (
  echo FA Lab Case Management did not start. Check: %ERROR_LOG_FILE%
  exit /b 1
)

>"%PID_FILE%" echo %SERVER_PID%
timeout /t 1 /nobreak >nul

tasklist /FI "PID eq %SERVER_PID%" /NH | findstr /I "%SERVER_PID%" >nul
if errorlevel 1 (
  echo FA Lab Case Management did not start. Check: %ERROR_LOG_FILE%
  del /q "%PID_FILE%"
  exit /b 1
)

echo FA Lab Case Management started in the background ^(PID %SERVER_PID%^).
echo Open: http://localhost:%PORT%
echo Log: %LOG_FILE%
