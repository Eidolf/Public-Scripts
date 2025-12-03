<#
.SYNOPSIS
    Script Name:    Exchange MoveLogs
    Script Info:    Moves Exchange Logs to other folder
    Created on:     01.01.2023
    Changed on:     03.12.2025
    Created by:     KSB
    Changed by:     Eidolf
    Company:        ER-Netz
    Version:        1.1.0
.DESCRIPTION
    This script will move all of the configurable logs for Exchange 2013 and upward from it's default place to a specified log path.
.EXAMPLE
    .\Exchange-MoveLogs.ps1 
    Script is used without attributes
.NOTES
    
#>

# Get the name of the local computer and  set it to a variable for use later on. 

$exchangeservername = $env:computername
$Logpath = read-host "Set destination log path / Ziel-Logpfad eingeben z.B. (L:\Logs)"

# Move the standard log files for  the TransportService to the same path on the L: drive that they were on C:  

Set-TransportService -Identity $exchangeservername -ConnectivityLogPath ("$Logpath"+"\Hub\Connectivity") -MessageTrackingLogPath ("$Logpath"+"\MessageTracking") -IrmLogPath ("$Logpath"+"\IRMLogs") -ActiveUserStatisticsLogPath ("$Logpath"+"\Hub\ActiveUsersStats") -ServerStatisticsLogPath ("$Logpath"+"\Hub\ServerStats") -ReceiveProtocolLogPath ("$Logpath"+"\Hub\ProtocolLog\SmtpReceive") -RoutingTableLogPath ("$Logpath"+"\Hub\Routing") -SendProtocolLogPath ("$Logpath"+"\Hub\ProtocolLog\SmtpSend") -QueueLogPath ("$Logpath"+"\Hub\QueueViewer") -WlmLogPath ("$Logpath"+"\Hub\WLM") -PipelineTracingPath ("$Logpath"+"\Hub\PipelineTracing") -AgentLogPath ("$Logpath"+"\Hub\AgentLog")

# move the path for  the PERFMON logs from the C: drive to the L: drive 

logman -stop ExchangeDiagnosticsDailyPerformanceLog

logman -update ExchangeDiagnosticsDailyPerformanceLog -o ("$Logpath"+"\Diagnostics\DailyPerformanceLogs\ExchangeDiagnosticsDailyPerformanceLog")

logman -start ExchangeDiagnosticsDailyPerformanceLog

logman -stop ExchangeDiagnosticsPerformanceLog

logman -update ExchangeDiagnosticsPerformanceLog -o ("$Logpath"+"\Diagnostics\PerformanceLogsToBeProcessed\ExchangeDiagnosticsPerformanceLog")

logman -start ExchangeDiagnosticsPerformanceLog

# Get the details on the EdgeSyncServiceConfig and  store them in a variable for use in setting the path 

$EdgeSyncServiceConfigVAR = Get-EdgeSyncServiceConfig 

# Move the Log Path using the variable we got 

Set-EdgeSyncServiceConfig -Identity $EdgeSyncServiceConfigVAR.Identity -LogPath ("$Logpath"+"\EdgeSync")

# Move the standard log files for  the FrontEndTransportService to the same path on the L: drive that they were on C: 

Set-FrontendTransportService  -Identity $exchangeservername -AgentLogPath ("$Logpath"+"\FrontEnd\AgentLog") -ConnectivityLogPath ("$Logpath"+"\FrontEnd\Connectivity") -ReceiveProtocolLogPath ("$Logpath"+"\FrontEnd\ProtocolLog\SmtpReceive") -SendProtocolLogPath ("$Logpath"+"\FrontEnd\ProtocolLog\SmtpSend")

# MOve the log path for  the IMAP server 

Set-ImapSettings -LogFileLocation ("$Logpath"+"\Imap4")

# Move the logs for  the MailBoxServer 

Set-MailboxServer -Identity $exchangeservername -CalendarRepairLogPath ("$Logpath"+"\Calendar Repair Assistant") -MigrationLogFilePath  ("$Logpath"+"\Managed Folder Assistant")

# Move the standard log files for  the MailboxTransportService to the same path on the L: drive that they were on C: 

Set-MailboxTransportService -Identity $exchangeservername -ConnectivityLogPath ("$Logpath"+"\Mailbox\Connectivity") -MailboxDeliveryAgentLogPath ("$Logpath"+"\Mailbox\AgentLog\Delivery") -MailboxSubmissionAgentLogPath ("$Logpath"+"\Mailbox\AgentLog\Submission") -ReceiveProtocolLogPath ("$Logpath"+"\Mailbox\ProtocolLog\SmtpReceive") -SendProtocolLogPath ("$Logpath"+"\Mailbox\ProtocolLog\SmtpSend") -PipelineTracingPath ("$Logpath"+"\Mailbox\PipelineTracing")

# Move the log path for  the POP3 server 

Set-PopSettings -LogFileLocation ("$Logpath"+"\Pop3")