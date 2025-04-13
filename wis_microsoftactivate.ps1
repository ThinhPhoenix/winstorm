$troubleshoot = 'https://massgrave.dev/troubleshoot'

# Ensure PowerShell is running in Full Language Mode
if ($ExecutionContext.SessionState.LanguageMode.value__ -ne 0) {
    Write-Host "Windows PowerShell is not running in Full Language Mode."
    Write-Host "Help - https://gravesoft.dev/fix_powershell" -ForegroundColor White -BackgroundColor Blue
    return
}

# Function to check for third-party antivirus software
function Check3rdAV {
    $avList = Get-CimInstance -Namespace root\SecurityCenter2 -Class AntiVirusProduct | 
              Where-Object { $_.displayName -notlike '*windows*' } | 
              Select-Object -ExpandProperty displayName
    if ($avList) {
        Write-Host "3rd party Antivirus might be blocking the script - " -ForegroundColor White -BackgroundColor Blue -NoNewline
        Write-Host " $($avList -join ', ')" -ForegroundColor DarkRed -BackgroundColor White
    }
}

# Function to verify the file creation in the temp folder
function CheckFile { 
    param ([string]$FilePath)
    if (-not (Test-Path $FilePath)) { 
        Check3rdAV
        Write-Host "Failed to create MAS file in temp folder, aborting!"
        Write-Host "Help - $troubleshoot" -ForegroundColor White -BackgroundColor Blue
        throw
    }
}

# Set security protocol for HTTPS connections
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# URLs to retrieve the script
$URLs = @(
    'https://raw.githubusercontent.com/ThinhPhoenix/winstorm/refs/heads/main/wis_microsoftactivate.cmd'
)

# Fetch the script from one of the URLs
$response = $null
foreach ($URL in $URLs | Sort-Object { Get-Random }) {
    try { $response = Invoke-WebRequest -Uri $URL -UseBasicParsing; break } catch {}
}

# If the script couldn't be fetched
if (-not $response) {
    Check3rdAV
    Write-Host "Failed to retrieve the script from the repository, aborting!"
    Write-Host "Help - $troubleshoot" -ForegroundColor White -BackgroundColor Blue
    return
}

# Hash verification to ensure script integrity
$releaseHash = '919F17B46BF62169E8811201F75EFDF1D5C1504321B78A7B0FB47C335ECBC1B0'
$hash = [BitConverter]::ToString([Security.Cryptography.SHA256]::Create().ComputeHash([System.Text.Encoding]::UTF8.GetBytes($response.Content))) -replace '-'
if ($hash -ne $releaseHash) {
    Write-Warning "Hash ($hash) mismatch, aborting! Report this issue at $troubleshoot"
    return
}

# Check for any autorun registry entries
$paths = "HKCU:\SOFTWARE\Microsoft\Command Processor", "HKLM:\SOFTWARE\Microsoft\Command Processor"
foreach ($path in $paths) { 
    if (Get-ItemProperty -Path $path -Name "Autorun" -ErrorAction SilentlyContinue) { 
        Write-Warning "Autorun registry found, CMD may crash! Remove the entry: Remove-ItemProperty -Path '$path' -Name 'Autorun'"
    }
}

# Prepare the file path and write the script content
$rand = [Guid]::NewGuid().Guid
$isAdmin = [bool]([Security.Principal.WindowsIdentity]::GetCurrent().Groups -match 'S-1-5-32-544')
$FilePath = if ($isAdmin) { "$env:SystemRoot\Temp\MAS_$rand.cmd" } else { "$env:USERPROFILE\AppData\Local\Temp\MAS_$rand.cmd" }
Set-Content -Path $FilePath -Value "@::: $rand `r`n$response.Content"

# Ensure the file is created
CheckFile $FilePath

# Verify cmd.exe is working
$chkcmd = & "$env:SystemRoot\system32\cmd.exe" /c "echo CMD is working"
if ($chkcmd -notcontains "CMD is working") {
    Write-Warning "cmd.exe is not working. Report this issue at $troubleshoot"
    return
}

# Run the script and wait for completion
Start-Process -FilePath $env:SystemRoot\system32\cmd.exe -ArgumentList "/c `"$FilePath`" $args" -Wait

# Clean up temporary files
Get-Item "$env:SystemRoot\Temp\MAS*.cmd", "$env:USERPROFILE\AppData\Local\Temp\MAS*.cmd" | Remove-Item
