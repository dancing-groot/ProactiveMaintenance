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

# Check minimum OS version requirement (1809 or later)
[int]$CurrentBuild = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name "CurrentBuildNumber" | Select-Object -ExpandProperty CurrentBuildNumber
If ($CurrentBuild -notin (17763, 18363, 19041, 19042) -and $CurrentBuild -lt 19043)
{
    Write-Host "Minimum OS version requirement not met"
    Exit 0
}

# Check if Update tools installed
$MSUHT = Get-WmiObject -Class Win32_Product -Filter "Name = 'Microsoft Update Health Tools'" | Select-Object -Property Name, Version, InstallDate, IdentifyingNumber
If ($null -eq $MSUHT)
{
    # Application is missing
    Invoke-Toast -Message "Update Tools not found" -MessageType Warning
    Exit 1
}
Else
{
    # Application is installed - Assuming there's only one entry matching
    $MSUHT = $MSUHT[0]
    $result = [PSCustomObject]@{
        DisplayName    = $MSUHT.Name
        DisplayVersion = $MSUHT.Version
        InstallDate    = ([System.DateTime]::ParseExact($MSUHT.InstallDate, 'yyyyMMdd', $null)).ToString("dd/MM/yyyy")
        GUID           = $MSUHT.IdentifyingNumber
    }
    Invoke-Toast -Message ($result | ConvertTo-Json -Compress)
    Exit 0
}