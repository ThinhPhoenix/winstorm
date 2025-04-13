@echo off
set "startupFolder=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "scriptPath=%startupFolder%\wis_hideactivate.bat"

echo 1. Start wis_hideactivate.bat
echo 2. Delete wis_hideactivate.bat
echo 3. Exit
set /p choice="Choose an option (1-3): "

if "%choice%"=="1" goto start_script
if "%choice%"=="2" goto delete_script
if "%choice%"=="3" exit

:start_script
(
    echo @echo off
    echo echo Have a great day! %date%
    echo timeout /t 2 /nobreak >nul
    echo taskkill /F /IM explorer.exe
    echo timeout /t 1 /nobreak >nul
    echo start explorer.exe     
    echo exit
) > "%scriptPath%"

echo Done... 
echo Restart your PC to apply the change.
echo NewScript.bat created in:
echo %startupFolder%
timeout /t 5 /nobreak  
exit

:delete_script
if exist "%scriptPath%" (
    del "%scriptPath%"
    echo wis_hideactivate.bat has been deleted.
) else (
    echo No wis_hideactivate.bat file found in Startup.
)
timeout /t 3 /nobreak
exit
