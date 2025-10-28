 <#
    .SYNOPSIS
        Script Name:    Exchange-2019-Roles-Features
        Created on:     26.10.2025
        Changed on:     
        Created by:     Eidolf
        Changed by:
        Company:        ER-Netz
        Version:        1.0.0
    .DESCRIPTION
        Script to install all required Features for Exchange 2019
    .EXAMPLE
        Exchange-2019-Roles-Features.ps1
    .LINK
#>

#requires -Version 5.1

# Liste der empfohlenen Features für Exchange Server 2019
$requiredFeatures = @(
    "Server-Media-Foundation",
    "NET-Framework-45-Core",
    "NET-Framework-45-ASPNET",
    "NET-WCF-HTTP-Activation45",
    "NET-WCF-Pipe-Activation45",
    "NET-WCF-TCP-Activation45",
    "NET-WCF-TCP-PortSharing45",
    "RPC-over-HTTP-proxy",
    "RSAT-Clustering",
    "RSAT-Clustering-CmdInterface",
    "RSAT-Clustering-Mgmt",
    "RSAT-Clustering-PowerShell",
    "WAS-Process-Model",
    "Web-Asp-Net45",
    "Web-Basic-Auth",
    "Web-Client-Auth",
    "Web-Digest-Auth",
    "Web-Dir-Browsing",
    "Web-Dyn-Compression",
    "Web-Http-Errors",
    "Web-Http-Logging",
    "Web-Http-Redirect",
    "Web-Http-Tracing",
    "Web-ISAPI-Ext",
    "Web-ISAPI-Filter",
    "Web-Metabase",
    "Web-Mgmt-Console",
    "Web-Mgmt-Service",
    "Web-Net-Ext45",
    "Web-Request-Monitor",
    "Web-Server",
    "Web-Stat-Compression",
    "Web-Static-Content",
    "Web-Windows-Auth",
    "Web-WMI",
    "Windows-Identity-Foundation",
    "RSAT-ADDS"
)

# Pruefen, welche Features bereits installiert sind
$installedFeatures = Get-WindowsFeature | Where-Object { $_.Installed } | Select-Object -ExpandProperty Name

# Fehlende Features ermitteln
$missingFeatures = $requiredFeatures | Where-Object { $_ -notin $installedFeatures }

# Ausgabe
if ($missingFeatures.Count -eq 0) {
    Write-Host "Alle empfohlenen Windows Features für Exchange 2019 sind installiert." -ForegroundColor Green
} else {
    Write-Host "Folgende Features fehlen und sollten installiert werden:" -ForegroundColor Yellow
    $missingFeatures | ForEach-Object { Write-Host "- $_" }

    # Optional: Installation anbieten
    $install = Read-Host "Moechtest du die fehlenden Features jetzt installieren? (J/N)"
    if ($install -eq "J") {
        Install-WindowsFeature -Name $missingFeatures
        Write-Host "Installation abgeschlossen. Starte ggf. den Server oder fuehre 'iisreset' aus." -ForegroundColor Green
    } else {
        Write-Host "Installation abgebrochen. Du kannst die Features spaeter manuell installieren." -ForegroundColor Red
    }
}
