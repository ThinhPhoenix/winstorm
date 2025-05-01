# Load required assemblies
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

# Function to reset taskbar
function Reset-Taskbar {
    try {
        # Get Explorer process
        $explorer = Get-Process -Name explorer -ErrorAction SilentlyContinue
        
        if ($explorer) {
            # Kill and restart Explorer process
            $explorer | Stop-Process -Force
            Start-Sleep -Seconds 1
            Start-Process explorer
            return $true
        }
        else {
            # Just start Explorer if it wasn't running
            Start-Process explorer
            return $true
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error resetting taskbar: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
}

# Function to set desktop wallpaper
function Set-WallPaper {
    param (
        [string]$ImagePath
    )
    
    Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    
    public class Wallpaper {
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
    }
"@
    
    $SPI_SETDESKWALLPAPER = 0x0014
    $SPIF_UPDATEINIFILE = 0x01
    $SPIF_SENDCHANGE = 0x02
    
    $ret = [Wallpaper]::SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, $ImagePath, $SPIF_UPDATEINIFILE -bor $SPIF_SENDCHANGE)
    return $ret
}

# Function to set theme (Light/Dark)
function Set-WindowsTheme {
    param (
        [string]$Theme
    )
    
    try {
        if ($Theme -eq "Dark") {
            # Set Dark theme
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Value 0
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "SystemUsesLightTheme" -Value 0
            
            # Reset taskbar
            Reset-Taskbar
            return $true
        }
        elseif ($Theme -eq "Light") {
            # Set Light theme
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Value 1
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "SystemUsesLightTheme" -Value 1
            
            # Reset taskbar
            Reset-Taskbar
            return $true
        }
        else {
            return $false
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error setting theme: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
}

# Function to set lock screen wallpaper
function Set-LockScreenWallpaper {
    param (
        [string]$ImagePath
    )
    
    try {
        # Windows 10/11 lock screen registry path
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"
        
        # Create registry path if it doesn't exist
        if (!(Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }
        
        # Set the lock screen image
        New-ItemProperty -Path $regPath -Name "LockScreenImagePath" -Value $ImagePath -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $regPath -Name "LockScreenImageStatus" -Value 1 -PropertyType DWORD -Force | Out-Null
        
        # Alternative method for some Windows versions
        $destinationPath = "$env:WINDIR\System32\oobe\info\backgrounds"
        if (!(Test-Path $destinationPath)) {
            New-Item -Path $destinationPath -ItemType Directory -Force | Out-Null
        }
        
        Copy-Item -Path $ImagePath -Destination "$destinationPath\backgroundDefault.jpg" -Force
        
        return $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error setting lock screen wallpaper: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
}

# Function to hide Windows activation notification
function Hide-WindowsActivation {
    try {
        $startupFolder = [System.IO.Path]::Combine($env:APPDATA, "Microsoft\Windows\Start Menu\Programs\Startup")
        $scriptPath = [System.IO.Path]::Combine($startupFolder, "wis_hideactivate.bat")
        
        $hideActivateForm = New-Object System.Windows.Forms.Form
        $hideActivateForm.Text = "Hide Windows Activation"
        $hideActivateForm.Size = New-Object System.Drawing.Size(400, 250)
        $hideActivateForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $hideActivateForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
        $hideActivateForm.MaximizeBox = $false
        $hideActivateForm.MinimizeBox = $false
        
        $descriptionLabel = New-Object System.Windows.Forms.Label
        $descriptionLabel.Text = "This will hide the Windows activation watermark`nby restarting the explorer process at startup."
        $descriptionLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $descriptionLabel.Size = New-Object System.Drawing.Size(380, 50)
        $descriptionLabel.Location = New-Object System.Drawing.Point(10, 20)
        $hideActivateForm.Controls.Add($descriptionLabel)
        
        $startButton = New-Object System.Windows.Forms.Button
        $startButton.Text = "Enable Hiding"
        $startButton.Location = New-Object System.Drawing.Point(50, 90)
        $startButton.Size = New-Object System.Drawing.Size(120, 40)
        
        $deleteButton = New-Object System.Windows.Forms.Button
        $deleteButton.Text = "Disable Hiding"
        $deleteButton.Location = New-Object System.Drawing.Point(220, 90)
        $deleteButton.Size = New-Object System.Drawing.Size(120, 40)
        
        $closeButton = New-Object System.Windows.Forms.Button
        $closeButton.Text = "Cancel"
        $closeButton.Location = New-Object System.Drawing.Point(150, 160)
        $closeButton.Size = New-Object System.Drawing.Size(100, 35)
        
        $startButton.Add_Click({
            # Create the batch file in the Startup folder
            $batchContent = @"
@echo off
echo Hide Windows Activation - $(Get-Date)
timeout /t 2 /nobreak
taskkill /F /IM explorer.exe
timeout /t 1 /nobreak
start explorer.exe
exit
"@
            Set-Content -Path $scriptPath -Value $batchContent
            [System.Windows.Forms.MessageBox]::Show("Done! Restart your PC to apply the change.`nScript created at: $scriptPath", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $statusLabel.Text = "Hiding Windows activation enabled!"
            $statusLabel.ForeColor = [System.Drawing.Color]::Green
            $hideActivateForm.Close()
        })
        
        $deleteButton.Add_Click({
            # Delete the batch file if it exists
            if (Test-Path -Path $scriptPath) {
                Remove-Item -Path $scriptPath
                [System.Windows.Forms.MessageBox]::Show("Hiding script has been removed from startup.`nChanges will take effect after restarting your PC.", "Script Removed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                $statusLabel.Text = "Hiding Windows activation disabled"
                $statusLabel.ForeColor = [System.Drawing.Color]::Blue
            } else {
                [System.Windows.Forms.MessageBox]::Show("No activation hiding script found in Startup.", "No Script Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            $hideActivateForm.Close()
        })
        
        $closeButton.Add_Click({
            $hideActivateForm.Close()
        })
        
        $hideActivateForm.Controls.Add($startButton)
        $hideActivateForm.Controls.Add($deleteButton)
        $hideActivateForm.Controls.Add($closeButton)
        
        $hideActivateForm.ShowDialog() | Out-Null
        return $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error with Windows activation hiding: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
}

# Function to Activate Windows/Office
function Activate-WindowsOffice {
    try {
        $activateForm = New-Object System.Windows.Forms.Form
        $activateForm.Text = "Activate Windows/Office"
        $activateForm.Size = New-Object System.Drawing.Size(400, 300)
        $activateForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $activateForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
        $activateForm.MaximizeBox = $false
        $activateForm.MinimizeBox = $false
        
        $warningLabel = New-Object System.Windows.Forms.Label
        $warningLabel.Text = "WARNING:`nThis will attempt to activate Windows and/or Office`nusing an online activation script.`n`nOnly proceed if you understand the implications."
        $warningLabel.Font = New-Object System.Drawing.Font("Arial", 10)
        $warningLabel.Size = New-Object System.Drawing.Size(380, 100)
        $warningLabel.Location = New-Object System.Drawing.Point(10, 20)
        $warningLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $warningLabel.ForeColor = [System.Drawing.Color]::Red
        $activateForm.Controls.Add($warningLabel)
        
        $activateButton = New-Object System.Windows.Forms.Button
        $activateButton.Text = "Run Activation"
        $activateButton.Location = New-Object System.Drawing.Point(120, 130)
        $activateButton.Size = New-Object System.Drawing.Size(150, 40)
        
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Text = "Cancel"
        $cancelButton.Location = New-Object System.Drawing.Point(120, 190)
        $cancelButton.Size = New-Object System.Drawing.Size(150, 40)
        
        $activateButton.Add_Click({
            $statusLabel.Text = "Running activation script..."
            $statusLabel.ForeColor = [System.Drawing.Color]::Blue
            $form.Refresh()
            
            try {
                # Create a temporary PowerShell script that runs the activation command
                $tempScriptPath = [System.IO.Path]::Combine($env:TEMP, "activate_script.ps1")
                Set-Content -Path $tempScriptPath -Value 'irm https://get.activated.win | iex'
                
                # Run the script with elevated privileges
                $startInfo = New-Object System.Diagnostics.ProcessStartInfo
                $startInfo.FileName = "powershell.exe"
                $startInfo.Arguments = "-ExecutionPolicy Bypass -File `"$tempScriptPath`""
                $startInfo.Verb = "runas" # Run as administrator
                
                [System.Diagnostics.Process]::Start($startInfo)
                
                $statusLabel.Text = "Activation script launched!"
                $statusLabel.ForeColor = [System.Drawing.Color]::Green
            }
            catch {
                $statusLabel.Text = "Activation failed: $_"
                $statusLabel.ForeColor = [System.Drawing.Color]::Red
                [System.Windows.Forms.MessageBox]::Show("Failed to run activation script: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
            
            $activateForm.Close()
        })
        
        $cancelButton.Add_Click({
            $activateForm.Close()
        })
        
        $activateForm.Controls.Add($activateButton)
        $activateForm.Controls.Add($cancelButton)
        
        $activateForm.ShowDialog() | Out-Null
        return $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error with activation process: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
}

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Win🌩️"
$form.Size = New-Object System.Drawing.Size(400, 450) # Increased height for new buttons
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

# Title Label
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Tweaks"
$titleLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
$titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$titleLabel.Size = New-Object System.Drawing.Size(380, 30)
$titleLabel.Location = New-Object System.Drawing.Point(10, 15)
$form.Controls.Add($titleLabel)

# Theme Button
$themeButton = New-Object System.Windows.Forms.Button
$themeButton.Text = "Change Theme (Light/Dark)"
$themeButton.Location = New-Object System.Drawing.Point(100, 70)
$themeButton.Size = New-Object System.Drawing.Size(200, 40)
$form.Controls.Add($themeButton)

# Theme Button Click Event
$themeButton.Add_Click({
    $themeForm = New-Object System.Windows.Forms.Form
    $themeForm.Text = "Theme Selection"
    $themeForm.Size = New-Object System.Drawing.Size(300, 200)
    $themeForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $themeForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $themeForm.MaximizeBox = $false
    $themeForm.MinimizeBox = $false
    
    $themeLabel = New-Object System.Windows.Forms.Label
    $themeLabel.Text = "Select a system theme:"
    $themeLabel.Font = New-Object System.Drawing.Font("Arial", 10)
    $themeLabel.Size = New-Object System.Drawing.Size(280, 25)
    $themeLabel.Location = New-Object System.Drawing.Point(10, 20)
    $themeForm.Controls.Add($themeLabel)
    
    $lightButton = New-Object System.Windows.Forms.Button
    $lightButton.Text = "Light"
    $lightButton.Location = New-Object System.Drawing.Point(50, 60)
    $lightButton.Size = New-Object System.Drawing.Size(80, 35)
    
    $darkButton = New-Object System.Windows.Forms.Button
    $darkButton.Text = "Dark"
    $darkButton.Location = New-Object System.Drawing.Point(150, 60)
    $darkButton.Size = New-Object System.Drawing.Size(80, 35)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(100, 120)
    $cancelButton.Size = New-Object System.Drawing.Size(80, 30)
    $lightButton.Add_Click({
        $statusLabel.Text = "Setting Light theme and resetting taskbar..."
        $statusLabel.ForeColor = [System.Drawing.Color]::Blue
        $form.Refresh()
        
        $success = Set-WindowsTheme -Theme "Light"
        if ($success) {
            $statusLabel.Text = "Theme changed to Light successfully!"
            $statusLabel.ForeColor = [System.Drawing.Color]::Green
            [System.Windows.Forms.MessageBox]::Show("Successfully set theme to Light and reset taskbar", "Theme Change", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            $statusLabel.Text = "Error changing theme"
            $statusLabel.ForeColor = [System.Drawing.Color]::Red
        }
        $themeForm.Close()
    })
    
    $darkButton.Add_Click({
        $statusLabel.Text = "Setting Dark theme and resetting taskbar..."
        $statusLabel.ForeColor = [System.Drawing.Color]::Blue
        $form.Refresh()
        
        $success = Set-WindowsTheme -Theme "Dark"
        if ($success) {
            $statusLabel.Text = "Theme changed to Dark successfully!"
            $statusLabel.ForeColor = [System.Drawing.Color]::Green
            [System.Windows.Forms.MessageBox]::Show("Successfully set theme to Dark and reset taskbar", "Theme Change", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            $statusLabel.Text = "Error changing theme"
            $statusLabel.ForeColor = [System.Drawing.Color]::Red
        }
        $themeForm.Close()
    })
    
    $cancelButton.Add_Click({
        $themeForm.Close()
    })
    
    $themeForm.Controls.Add($lightButton)
    $themeForm.Controls.Add($darkButton)
    $themeForm.Controls.Add($cancelButton)
    
    $themeForm.ShowDialog() | Out-Null
})

# Wallpaper Button
$wallpaperButton = New-Object System.Windows.Forms.Button
$wallpaperButton.Text = "Set Desktop Wallpaper"
$wallpaperButton.Location = New-Object System.Drawing.Point(100, 120)
$wallpaperButton.Size = New-Object System.Drawing.Size(200, 40)
$form.Controls.Add($wallpaperButton)

# Wallpaper Button Click Event
$wallpaperButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Image Files (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp"
    $openFileDialog.Title = "Select Desktop Wallpaper Image"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $wallpaperPath = $openFileDialog.FileName
        $success = Set-WallPaper -ImagePath $wallpaperPath
        
        if ($success) {
            [System.Windows.Forms.MessageBox]::Show("Desktop wallpaper set successfully!", "Wallpaper Change", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Failed to set desktop wallpaper.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# Lock Screen Button
$lockScreenButton = New-Object System.Windows.Forms.Button
$lockScreenButton.Text = "Set Lock Screen Wallpaper"
$lockScreenButton.Location = New-Object System.Drawing.Point(100, 170)
$lockScreenButton.Size = New-Object System.Drawing.Size(200, 40)
$form.Controls.Add($lockScreenButton)

# Lock Screen Button Click Event
$lockScreenButton.Add_Click({
    # Notify about Administrator rights
    [System.Windows.Forms.MessageBox]::Show("Note: Setting lock screen wallpaper requires Administrator privileges. The script will attempt to set it, but may fail without proper permissions.", "Administrator Rights Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Image Files (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp"
    $openFileDialog.Title = "Select Lock Screen Wallpaper Image"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $lockWallpaperPath = $openFileDialog.FileName
        $success = Set-LockScreenWallpaper -ImagePath $lockWallpaperPath
        
        if ($success) {
            [System.Windows.Forms.MessageBox]::Show("Lock screen wallpaper set successfully!", "Lock Screen Change", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Failed to set lock screen wallpaper. Try running the script as Administrator.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# NEW BUTTON: Hide Windows Activation
$hideActivateButton = New-Object System.Windows.Forms.Button
$hideActivateButton.Text = "Hide Windows Activation"
$hideActivateButton.Location = New-Object System.Drawing.Point(100, 220)
$hideActivateButton.Size = New-Object System.Drawing.Size(200, 40)
$form.Controls.Add($hideActivateButton)

# Hide Windows Activation Click Event
$hideActivateButton.Add_Click({
    $statusLabel.Text = "Managing Windows activation notification..."
    $statusLabel.ForeColor = [System.Drawing.Color]::Blue
    $form.Refresh()
    
    $success = Hide-WindowsActivation
    if (-not $success) {
        $statusLabel.Text = "Error with hide activation function"
        $statusLabel.ForeColor = [System.Drawing.Color]::Red
    }
})

# NEW BUTTON: Activate Windows/Office
$activateButton = New-Object System.Windows.Forms.Button
$activateButton.Text = "Activate Windows/Office"
$activateButton.Location = New-Object System.Drawing.Point(100, 270)
$activateButton.Size = New-Object System.Drawing.Size(200, 40)
$form.Controls.Add($activateButton)

# Activate Windows/Office Click Event
$activateButton.Add_Click({
    $statusLabel.Text = "Preparing activation options..."
    $statusLabel.ForeColor = [System.Drawing.Color]::Blue
    $form.Refresh()
    
    $success = Activate-WindowsOffice
    if (-not $success) {
        $statusLabel.Text = "Error with activation function"
        $statusLabel.ForeColor = [System.Drawing.Color]::Red
    }
})

# Status Label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Ready"
$statusLabel.Size = New-Object System.Drawing.Size(380, 20)
$statusLabel.Location = New-Object System.Drawing.Point(10, 385)
$statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$form.Controls.Add($statusLabel)

# Quit Button
$quitButton = New-Object System.Windows.Forms.Button
$quitButton.Text = "Quit"
$quitButton.Location = New-Object System.Drawing.Point(100, 330)
$quitButton.Size = New-Object System.Drawing.Size(200, 40)
$form.Controls.Add($quitButton)

# Quit Button Click Event
$quitButton.Add_Click({
    $form.Close()
})

# Show the form
$form.ShowDialog() | Out-Null