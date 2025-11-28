<#
.SYNOPSIS
    Script Name:    Exchange GetDefaultSMTP
    Script Info:    Retrieves all Exchange servers and their default SMTP certificate details, including expiry warnings and export option.
    Created on:     28.11.2025
    Changed on:     
    Created by:     Eidolf (with help of Copilot AI)
    Changed by:
    Company:        ER-Netz
    Version:        1.0.0
.DESCRIPTION
    Detects Exchange servers, reads the active SMTP certificate per server (latest valid),
    shows details in GridView with an export option. DaysUntilExpiry is an integer (no decimals).
    Export goes to the script folder.
.EXAMPLE
    .\Exchange-GetDefaultSMTP.ps1 
    Skript wird ohne Attribute ausgeführt.
.NOTES
    Exchange Module is needed
#>

# Check if Exchange cmdlets are available
if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) {
    Write-Host "Exchange cmdlets not available, creating remote session..." -ForegroundColor Yellow

    $TargetExchangeServer = "YourExchangeServer"
    try {
        $TempSession = New-PSSession -ConfigurationName Microsoft.Exchange `
                                     -ConnectionUri "http://$TargetExchangeServer/PowerShell/" `
                                     -Authentication Kerberos
        Import-PSSession $TempSession -DisableNameChecking
    }
    catch {
        Write-Error "Failed to create remote session. Please verify server name and connectivity."
        return
    }
} else {
    Write-Host "Exchange cmdlets already available, using local session." -ForegroundColor Green
}

# Retrieve all Exchange servers with Mailbox role
$ExchangeServers = Get-ExchangeServer | Where-Object { $_.ServerRole -like "*Mailbox*" }

if (-not $ExchangeServers) {
    Write-Warning "No Exchange servers detected."
    return
}

$Results = @()

foreach ($Server in $ExchangeServers) {
    # Get all valid SMTP certificates (not expired)
    $Certs = Get-ExchangeCertificate -Server $Server.Name | Where-Object {
        $_.Services -like "*SMTP*" -and $_.NotAfter -gt (Get-Date)
    }

    if ($Certs) {
        # Pick the one with the latest expiry
        $DefaultCert = $Certs | Sort-Object NotAfter -Descending | Select-Object -First 1

        # Integer days (no decimals)
        $DaysUntilExpiry = [int][math]::Floor( ($DefaultCert.NotAfter - (Get-Date)).TotalDays )

        # Optional: string version to avoid any thousand separators in GridView
        $DaysUntilExpiryStr = $DaysUntilExpiry.ToString()   # uncomment to use string instead of int

        if ($DaysUntilExpiry -lt 30) {
            Write-Warning "Certificate on $($Server.Name) expires in $DaysUntilExpiry days!"
        }

        $Results += [PSCustomObject]@{
            ServerName           = $Server.Name
            Edition              = $Server.Edition
            AdminDisplayVersion  = $Server.AdminDisplayVersion
            CertSubject          = $DefaultCert.Subject
            CertFriendlyName     = $DefaultCert.FriendlyName
            CertThumbprint       = $DefaultCert.Thumbprint
            CertExpiryDate       = $DefaultCert.NotAfter
            # DaysUntilExpiry      = $DaysUntilExpiry        # integer; shows without decimals
            DaysUntilExpiryText = $DaysUntilExpiryStr    # switch GridView to show this if you want zero separators
        }
    } else {
        Write-Warning "No valid SMTP certificate found for $($Server.Name)."
    }
}

# Show results in GridView; instruct user to click OK to export
$Selected = $Results | Out-GridView -Title "Select rows and click OK to EXPORT > Exchange SMTP Certificate Details" -PassThru

if ($Selected) {
    $ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
    $ExportPath = Join-Path $ScriptPath "ExchangeSMTPCerts.csv"

    $Selected | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
    Write-Host "Selected rows exported to $ExportPath" -ForegroundColor Green
} else {
    Write-Host "No rows selected for export." -ForegroundColor Yellow
}
