<#
.SYNOPSIS
  Microsoft Update Health Tools - Detection
.DESCRIPTION
  Proactive Remediation
.NOTES
  Version:        2021.09.14
  Author:         Alexandre Cop
#>

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
    Write-Host "Update Health Tools not found"
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
    Write-Host "Update Health Tools installed"
    Write-Host ($result | ConvertTo-Json -Compress)
    Exit 0
}