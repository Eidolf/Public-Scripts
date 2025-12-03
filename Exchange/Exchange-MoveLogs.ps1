
<#
.SYNOPSIS
    Script Name:    Exchange MoveLogs
    Script Info:    Moves Exchange Logs to other folder
    Created on:     01.01.2023
    Changed on:     03.12.2025
    Created by:     KSB
    Changed by:     Eidolf (with help of Copilot AI)
    Company:        ER-Netz
    Version:        1.2.0
.DESCRIPTION
    This script will move all of the configurable logs for Exchange 2013 and upward from its default place to a specified log path.
    Version 1.2.0 adds Dry-Run mode, logging, and improved symlink handling.
.EXAMPLE
    .\Exchange-MoveLogs.ps1 
    Script is used without attributes
.NOTES
    
#>

# -------------------------------
# Global Variables
# -------------------------------
$ExchangeServerName = $env:COMPUTERNAME
$LogPath = Read-Host "Set destination log path (e.g., L:\Logs)"
$DryRunInput = Read-Host "Dry-Run only? (Y/N)"
$DryRun = ($DryRunInput -eq "Y")
$LogFile = "$LogPath\ExchangeMoveLogs_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

Start-Transcript -Path $LogFile

# -------------------------------
# Helper: Execute or Simulate
# -------------------------------
function Execute-Action {
    param($ActionDescription, [ScriptBlock]$Action)

    if ($DryRun) {
        Write-Host "[DRY-RUN] $ActionDescription"
    } else {
        Write-Host $ActionDescription
        & $Action
    }
}

# -------------------------------
# Function: Move TransportService Logs
# -------------------------------
function Set-TransportServicePaths {
    param($ServerName, $BasePath)

    Execute-Action "Updating TransportService log paths for $ServerName..." {
        Set-TransportService -Identity $ServerName `
            -ConnectivityLogPath "$BasePath\Hub\Connectivity" `
            -MessageTrackingLogPath "$BasePath\MessageTracking" `
            -IrmLogPath "$BasePath\IRMLogs" `
            -ActiveUserStatisticsLogPath "$BasePath\Hub\ActiveUsersStats" `
            -ServerStatisticsLogPath "$BasePath\Hub\ServerStats" `
            -ReceiveProtocolLogPath "$BasePath\Hub\ProtocolLog\SmtpReceive" `
            -RoutingTableLogPath "$BasePath\Hub\Routing" `
            -SendProtocolLogPath "$BasePath\Hub\ProtocolLog\SmtpSend" `
            -QueueLogPath "$BasePath\Hub\QueueViewer" `
            -WlmLogPath "$BasePath\Hub\WLM" `
            -PipelineTracingPath "$BasePath\Hub\PipelineTracing" `
            -AgentLogPath "$BasePath\Hub\AgentLog" `
            -JournalLogPath "$BasePath\JournalLog" `
            -TransportHttpLogPath "$BasePath\Hub\TransportHttp"
    }
}

# -------------------------------
# Function: Move FrontendTransportService Logs
# -------------------------------
function Set-FrontendTransportServicePaths {
    param($ServerName, $BasePath)

    Execute-Action "Updating FrontendTransportService log paths..." {
        Set-FrontendTransportService -Identity $ServerName `
            -AgentLogPath "$BasePath\FrontEnd\AgentLog" `
            -ConnectivityLogPath "$BasePath\FrontEnd\Connectivity" `
            -ReceiveProtocolLogPath "$BasePath\FrontEnd\ProtocolLog\SmtpReceive" `
            -SendProtocolLogPath "$BasePath\FrontEnd\ProtocolLog\SmtpSend"
    }
}

# -------------------------------
# Function: Move MailboxTransportService Logs
# -------------------------------
function Set-MailboxTransportServicePaths {
    param($ServerName, $BasePath)

    Execute-Action "Updating MailboxTransportService log paths..." {
        Set-MailboxTransportService -Identity $ServerName `
            -ConnectivityLogPath "$BasePath\Mailbox\Connectivity" `
            -MailboxDeliveryAgentLogPath "$BasePath\Mailbox\AgentLog\Delivery" `
            -MailboxSubmissionAgentLogPath "$BasePath\Mailbox\AgentLog\Submission" `
            -ReceiveProtocolLogPath "$BasePath\Mailbox\ProtocolLog\SmtpReceive" `
            -SendProtocolLogPath "$BasePath\Mailbox\ProtocolLog\SmtpSend" `
            -PipelineTracingPath "$BasePath\Mailbox\PipelineTracing"
    }
}

# -------------------------------
# Function: Move EdgeSyncService Logs
# -------------------------------
function Set-EdgeSyncServicePath {
    param($BasePath)

    Execute-Action "Updating EdgeSyncService log path..." {
        $EdgeSyncServiceConfigVAR = Get-EdgeSyncServiceConfig
        Set-EdgeSyncServiceConfig -Identity $EdgeSyncServiceConfigVAR.Identity -LogPath "$BasePath\EdgeSync"
    }
}

# -------------------------------
# Function: Move IMAP and POP3 Logs
# -------------------------------
function Set-ImapPopPaths {
    param($BasePath)

    Execute-Action "Updating IMAP and POP3 log paths..." {
        Set-ImapSettings -LogFileLocation "$BasePath\Imap4"
        Set-PopSettings -LogFileLocation "$BasePath\Pop3"
    }
}

# -------------------------------
# Function: Move MailboxServer Logs
# -------------------------------
function Set-MailboxServerPaths {
    param($ServerName, $BasePath)

    Execute-Action "Updating MailboxServer log paths..." {
        Set-MailboxServer -Identity $ServerName `
            -CalendarRepairLogPath "$BasePath\Calendar Repair Assistant" `
            -MigrationLogFilePath "$BasePath\Managed Folder Assistant"
    }
}

# -------------------------------
# Function: Move Perfmon Logs
# -------------------------------
function Move-PerfmonLogs {
    param($BasePath)

    Execute-Action "Updating Perfmon log paths..." {
        logman -stop ExchangeDiagnosticsDailyPerformanceLog
        logman -update ExchangeDiagnosticsDailyPerformanceLog -o "$BasePath\Diagnostics\DailyPerformanceLogs\ExchangeDiagnosticsDailyPerformanceLog"
        logman -start ExchangeDiagnosticsDailyPerformanceLog

        logman -stop ExchangeDiagnosticsPerformanceLog
        logman -update ExchangeDiagnosticsPerformanceLog -o "$BasePath\Diagnostics\PerformanceLogsToBeProcessed\ExchangeDiagnosticsPerformanceLog"
        logman -start ExchangeDiagnosticsPerformanceLog
    }
}

# -------------------------------
# Function: Create Symlinks for Non-Configurable Paths
# -------------------------------
function Create-SpecialLogSymlinks {
    param($BasePath)

    $TransportService = Get-TransportService -Identity $ExchangeServerName
    $SpecialPaths = @{
        "LatencyLog"  = $TransportService.LatencyLogPath
        "GeneralLog"  = $TransportService.GeneralLogPath
    }

    Execute-Action "Stopping MSExchangeTransport service..." {
        Stop-Service MSExchangeTransport
    }

    foreach ($Name in $SpecialPaths.Keys) {
        $OldPath = $SpecialPaths[$Name]
        $NewPath = Join-Path $BasePath "Hub\$Name"

        Write-Host "Processing ${Name}:"
        Write-Host "  Old Path: $OldPath"
        Write-Host "  New Path: $NewPath"

        Execute-Action "Creating junction for ${Name}..." {
            if (!(Test-Path $NewPath)) {
                New-Item -ItemType Directory -Path $NewPath | Out-Null
            }

            if ((Test-Path $OldPath) -and (!(Get-Item $OldPath).Attributes.ToString().Contains("ReparsePoint"))) {
                Rename-Item $OldPath "${OldPath}_old"
                New-Item -ItemType Junction -Path $OldPath -Target $NewPath | Out-Null
            }
        }
    }

    Execute-Action "Starting MSExchangeTransport service..." {
        Start-Service MSExchangeTransport
    }
}

# -------------------------------
# Main Execution
# -------------------------------
Set-TransportServicePaths $ExchangeServerName $LogPath
Set-FrontendTransportServicePaths $ExchangeServerName $LogPath
Set-MailboxTransportServicePaths $ExchangeServerName $LogPath
Set-EdgeSyncServicePath $LogPath
Set-ImapPopPaths $LogPath
Set-MailboxServerPaths $ExchangeServerName $LogPath
Move-PerfmonLogs $LogPath
Create-SpecialLogSymlinks $LogPath

Write-Host "All log paths have been successfully updated!"
Stop-Transcript
