# Alexandre Cop
# 2021.09.14

Function Invoke-Toast
{
    Param (
        [string]$Title = "Microsoft Endpoint Manager",
        [string]$Message = "This is a test notification",
        [System.Windows.Forms.ToolTipIcon]$MessageType = "Info"
    )

    $timeout = 5

    $path = Get-Process -id $pid | Select-Object -ExpandProperty Path                   
    $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)           
    $notify = New-Object System.Windows.Forms.NotifyIcon
    $notify.Icon = $icon
    $notify.BalloonTipTitle = $Title
    $notify.BalloonTipText = $Message
    $notify.BalloonTipIcon = $MessageType
    $notify.Visible = $true
    Try
    {
        Write-Host $Message
        $notify.ShowBalloonTip($timeout)
        Start-Sleep $timeout
    }
    Catch
    {}
    $notify.Dispose()
}

# init
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# variables
$DownloadDirectory = "$env:Temp"
$DownloadFileName = "Expedite_packages.zip"
$LogDirectory = "$env:Temp"
$LogFile = "UpdateTools.log"

# Get the download URL
$ProgressPreference = 'SilentlyContinue'
$URL = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=103324"
$Request = Invoke-WebRequest -Uri $URL -UseBasicParsing
$DownloadURL = ($Request.Links | Where-Object { $_.outerHTML -match "click here to download manually" }).href

# Download and extract the ZIP package
Invoke-WebRequest -Uri $DownloadURL -OutFile "$DownloadDirectory\$DownloadFileName" -UseBasicParsing
If (Test-Path "$DownloadDirectory\$DownloadFileName")
{
    Expand-Archive -Path "$DownloadDirectory\$DownloadFileName" -DestinationPath $DownloadDirectory -Force
}
Else 
{
    Invoke-Toast -Message "Update tools not downloaded" -MessageType Error
    Exit 1
}

# Determine which cab to use
[int]$CurrentBuild = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name "CurrentBuildNumber" | Select-Object -ExpandProperty CurrentBuildNumber
Switch ($CurrentBuild)
{
    17763 { $dir = "1809" }
    18363 { $dir = "1909" }
    default { $dir = "2004 and above" }
}
[string]$Bitness = Get-CimInstance Win32_OperatingSystem -Property OSArchitecture | Select-Object -ExpandProperty OSArchitecture
Switch ($Bitness)
{
    "64-bit" { $arch = "x64" }
    "32-bit" { $arch = "x86" }
    default { Write-Host "Unable to determine OS architecture"; Exit 1 }
}
$CabLocation = "$DownloadDirectory\$($DownloadFileName.Split('.')[0])\$dir\$arch"
$CabName = (Get-ChildItem $CabLocation -Name *.cab).pschildname

# Expand the cab and get the MSI
expand.exe /r "$CabLocation\$CabName" /F:* $DownloadDirectory
$File = Get-Childitem -Path $DownloadDirectory\*.msi -File | Where-Object { ((Get-Date).ToUniversalTime() - $_.CreationTimeUTC).TotalSeconds -lt 10 }

# Install the MSI
$Process = Start-Process -FilePath msiexec.exe -ArgumentList "/i $($File.FullName) /qn REBOOT=ReallySuppress /L*V ""$LogDirectory\$LogFile""" -Wait -PassThru
If ($Process.ExitCode -eq 0)
{
    Invoke-Toast -Message "Microsoft Update Health tools successfully installed"
    Exit 0
}
else 
{
    Invoke-Toast -Message "Microsoft Update Health tools installation failed with exit code $($Process.ExitCode)" -MessageType Error
    Exit 1
} 
