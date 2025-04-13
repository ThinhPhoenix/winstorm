# Define the path to the startup folder and script
$startupFolder = [System.IO.Path]::Combine($env:APPDATA, "Microsoft\Windows\Start Menu\Programs\Startup")
$scriptPath = [System.IO.Path]::Combine($startupFolder, "wis_hideactivate.bat")

# Present the user with options
$choice = Read-Host "Choose an option (1- Start, 2- Delete, 3- Exit)"

switch ($choice) {
    1 {
        # Create the batch file in the Startup folder
        $batchContent = @"
@echo off
echo Have a great day! $(Get-Date)
timeout /t 2 /nobreak
taskkill /F /IM explorer.exe
timeout /t 1 /nobreak
start explorer.exe
exit
"@
        Set-Content -Path $scriptPath -Value $batchContent

        Write-Host "Done... Restart your PC to apply the change."
        Write-Host "NewScript.bat created in: $startupFolder"
        break
    }

    2 {
        # Delete the batch file if it exists
        if (Test-Path -Path $scriptPath) {
            Remove-Item -Path $scriptPath
            Write-Host "wis_hideactivate.bat has been deleted."
        } else {
            Write-Host "No wis_hideactivate.bat file found in Startup."
        }
        break
    }

    3 {
        # Exit the script
        Write-Host "Exiting..."
        break
    }

    default {
        Write-Host "Invalid choice. Exiting..."
        break
    }
}
