@echo off
set "startupFolder=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
(
    echo @echo off
    echo echo Have a great day! %date%
    echo timeout /t 2 /nobreak >nul
    echo taskkill /F /IM explorer.exe
    echo timeout /t 1 /nobreak >nul
    echo start explorer.exe     
    echo exit
) > "%startupFolder%\removewindowatermark.bat"

echo Done... 
echo Restart your PC to apply the change.
echo NewScript.bat created in:
echo %startupFolder%
timeout /t 5 /nobreak  
exit
