
<#
    .SYNOPSIS
        Script Name:    Exchange Migration Script  
        Created on:     17.11.2025
        Changed on:     04.12.2025
        Created by:     Eidolf (with help of Copilot AI)
        Changed by:     Eidolf (with help of Copilot AI)
        Company:        ER-Netz
        Version:        1.2.1
    .DESCRIPTION
        A Script for migrating Exchange Server 2016 to Server 2019/SE. Primary for Database creation and move and comparing settings between Servers.
        Version 1.0.0 is a working Version for Database Migration and to check settings.
        Version 1.1.0 Implemented CAS URL comparison
        Version 1.2.0 extends the mailbox move functionality:
            - Support for MigrationBatch instead of MoveRequest.
            - Grouping by mailbox type and target database (UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox).
    .EXAMPLE
        ============================================================================
        Exchange Migration Script
        Command Reference / Präambel – 2025-11-17
        ----------------------------------------------------------------------------
        Schnellstart – typische Aufrufe
          1) Plan anzeigen (Dry-Run):
             .\Exchange_Migration_Script.ps1 -SourceVersion 15.1 -TargetVersion 15.2
        
          2) Zielordner + Datenbanken erstellen (ohne Moves/Settings):
             .\Exchange_Migration_Script.ps1 -PrepareFolders -CreateDatabases -Approve
        
          3) DB-Einstellungen vergleichen (und anwenden):
             nur vergleichen
             .\Exchange_Migration_Script.ps1 -CompareSettings
             vergleichen + anwenden
             .\Exchange_Migration_Script.ps1 -CompareSettings -ApplySettings -Approve
        
          4) Move Requests einreihen (mit Suspend):
            [Wichtig: Dieser Befehl ist das absolute minimum, sinnvoller sind Befehle unter der Überschrift Beispiel-Workflows Punkt 4]
             .\Exchange_Migration_Script.ps1 -QueueMoves -BatchNamePrefix "mailboxmove" -Approve
        
          5) Interaktive Pfadanpassung vor Ausführung:
             .\Exchange_Migration_Script.ps1 -Interactive
    
        ----------------------------------------------------------------------------
        Parameter – Übersicht
        
         Core
           -SourceVersion <string>   Quell-Exchange-Version (z.B. 15.1). Default 15.1.
           -TargetVersion <string>   Ziel-Exchange-Version (z.B. 15.2).  Default 15.2.
           -IncludeDbs   <string[]>  Optional: Liste von Quell-DBs, auf die gefiltert wird.
        
         Pfadsteuerung
           -PathMode <MirrorSource|BasePerDb>  Pfadermittlung. Default MirrorSource.
           -DbBase <string>    Basisordner für DBs, falls benötigt (Default E:\Datenbanken).
           -LogBase <string>   Basisordner für LOGs (Default E:\Logs).
           -DriveMap <hashtable> Remapping von Laufwerken, z.B. @{ 'D:'='E:' }.
        
         Aktionen (werden nur mit -Approve ausgeführt)
           -PrepareFolders     Legt Zielordner (DB/LOG) auf dem Zielserver an.
           -CreateDatabases    Erstellt Mailbox-Datenbanken inkl. Pfade (optional Mount/Exclude).
           -CompareSettings    Vergleicht DB-Settings Quelle ↔ Ziel (Scope siehe -SettingsScope).
           -ApplySettings      Wendet gefundene Diffs auf Ziel-DBs an (setzt Set-MailboxDatabase).
           -QueueMoves         Erzeugt Move Requests von Quelle → Ziel (Batchname auto-generiert).
           -Interactive        Interaktives Anpassen der Zielpfade vor Ausführung.
           -Approve            Sicherheitshebel – ohne diesen keine schreibenden Aktionen.
        
         Optionen – DB-Erstellung
           -MountAfterCreate           (Switch, Default: $true)  Mount nach Erstellung.
           -ExcludeFromProvisioning    (Switch, Default: $true)  DB vom Provisioning ausschließen.
        
         Optionen – Settings-Abgleich
           -SettingsScope <string[]>   Gruppen: Quotas, Retention, Client, Maintenance, OAB, All (Default).
           -ExportDiffs   <string>     CSV-Pfad für Settings-Diffs.
        
         Optionen – Moves
           -SuspendWhenReadyToComplete (Switch, Default: $true) MoveRequests mit Suspend.
           -BatchNamePrefix <string>   Präfix für Batchnamen (Default: FrontierMove).
        
         Review & Export
           -ExportPlan    <string>     CSV-Pfad für Export des Plans.
        
         Debug
           -DebugOAB                   Detailausgabe zur OAB-Auflösung/Normalisierung.
        
        ----------------------------------------------------------------------------
        Wichtiges Verhalten
         - Ziel-DB-Name wird aus Quell-DB inkrementiert (DB01 → DB02; sonst Name+1).
         - DB-Erstellung ist idempotent: existierende Ziele werden übersprungen (Warnung).
         - Quotas/Timespans/Booleans/OAB werden für Set-MailboxDatabase konvertiert.
         - Ohne -Approve laufen ausschließlich Read-/Plan-Schritte (Dry-Run/Safety-Guard).
        ----------------------------------------------------------------------------
        Beispiel-Workflows
         (1) Plan → Ordner/DBs → Settings → Moves
             .\Exchange_Migration_Script.ps1
             .\Exchange_Migration_Script.ps1 -PrepareFolders -CreateDatabases -Approve
             .\Exchange_Migration_Script.ps1 -CompareSettings -ApplySettings -Approve -ExportDiffs C:\Temp\DbDiffs.csv
             .\Exchange_Migration_Script.ps1 -QueueMoves -Approve
        
         (2) Nur bestimmte DBs & interaktive Pfade
             .\Exchange_Migration_Script.ps1 -IncludeDbs DB01,DB02 -Interactive -PrepareFolders -CreateDatabases -Approve
        
         (3) Pfadverlagerung via BasePerDb + DriveMap
             .\Exchange_Migration_Script.ps1 -PathMode BasePerDb -DbBase "F:\DBs" -LogBase "G:\Logs" -DriveMap @{ 'E:'='F:'; 'L:'='G:' } -PrepareFolders -CreateDatabases -Approve
         (4) Postfach-Migration per MigrationBatch:
             Arbitration mailboxes:
             Standard (Batch anlegen, kein Start):
             .\Exchange_Migration_Script.ps1 -QueueMoves -Arbitration -BatchNamePrefix "arbitration" -BadItemLimit 10 -NotifyEmail admin@domain.com -Approve

             Standard (Batch anlegen, kein Start):
             .\Exchange_Migration_Script.ps1 -QueueMoves -BatchNamePrefix "mailboxmove" -BadItemLimit 10 -NotifyEmail admin@domain.com -Approve

             Batch sofort starten:
             .\Exchange_Migration_Script.ps1 -QueueMoves -BatchNamePrefix "mailboxmove" -BadItemLimit 10 -NotifyEmail admin@domain.com -Approve -AutoStart

             CSV-Dateien behalten (Debugging):
             .\Exchange_Migration_Script.ps1 -QueueMoves -BatchNamePrefix "mailboxmove" -BadItemLimit 10 -NotifyEmail admin@domain.com -Approve -KeepCsv

             Fallback-Ziel-DB für nicht gemappte Quell-DBs:
             .\Exchange_Migration_Script.ps1 -QueueMoves -BatchNamePrefix "mailboxmove" -BadItemLimit 10 -NotifyEmail admin@domain.com -Approve -FallbackTargetDb "MDB02"
        ============================================================================
    .NOTES
        Exchange Module is needed

        Hinweis – Servergebundene URLs
        - URLs, deren Host den Quell-Servernamen enthält, werden beim Apply NICHT automatisch übernommen.
        Stattdessen erscheint ein Hinweis im Vergleich. Anpassung manuell (oder per zentralem FQDN) empfohlen.

        Zusatzoptionen (CAS)
        - -CasShowAll   Zeigt im Vergleich alle Eigenschaften inkl. Gleichständen (Status: Equal / Equal (BothEmpty)).
    .LINK

#>


[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
param(
    [string]$SourceVersion = '15.1',
    [string]$TargetVersion = '15.2',
    [ValidateSet('MirrorSource','BasePerDb')]
    [string]$PathMode = 'MirrorSource',
    [string]$DbBase='E:\\Datenbanken',
    [string]$LogBase='E:\\Logs',
    [hashtable]$DriveMap,
    [string[]]$IncludeDbs,

    # Actions
    [switch]$PrepareFolders,
    [switch]$CreateDatabases,
    [switch]$CompareSettings,
    [switch]$ApplySettings,
    [switch]$QueueMoves,

    # DB create opts
    [switch]$MountAfterCreate=$true,
    [switch]$ExcludeFromProvisioning=$true,

    # Settings scope
    [ValidateSet('Quotas','Retention','Client','Maintenance','OAB','All')]
    [string[]]$SettingsScope=@('All'),

    # Move opts
    [switch]$SuspendWhenReadyToComplete=$true,
    [string]$BatchNamePrefix='FrontierMove',

    # Review/exports
    [switch]$Interactive,
    [switch]$Approve,
[string]$ForceTargetDb,
    [string]$ExportPlan,
    [string]$ExportDiffs,

    # Debug
    [switch]$DebugOAB,

    # CAS URLs
    [switch]$CompareCasUrls,
    [switch]$ApplyCasUrls,
    [string]$SourceCasServer,
    [string]$TargetCasServer,
    [switch]$CasShowAll,
[int]$BadItemLimit = 0,
[string]$NotifyEmail,
[switch]$Arbitration,
[switch]$AcceptLargeDataLoss
)

function Test-ExchangeEnvironment { 
    Write-Host "Exchange Migration Script started..." -ForegroundColor Cyan; if (-not (Get-Command Get-ExchangeServer -ErrorAction SilentlyContinue)) 
    { Write-Warning "Run inside Exchange Management Shell!"; return $false } return $true 
}
function Get-ExchangeServerVersions { 
    $s=Get-ExchangeServer; if (-not $s){Write-Warning "No Exchange servers detected.";return @()} 
    Write-Host "`nDetected Exchange Servers:" -ForegroundColor Yellow; $s|Select Name,Edition,AdminDisplayVersion|Format-Table -AutoSize|Out-Host; return $s 
}
function Normalize-VersionString([string]$text){ if ([string]::IsNullOrWhiteSpace($text)) { 
    return $null } if ($text -match 'Version\s+(\d+)\.(\d+)') { return ("{0}.{1}" -f [int]$matches[1],[int]$matches[2]) } 
    try{ $v=[version]$text; return ("{0}.{1}" -f $v.Major,$v.Minor)}
    catch{ $null }
}
function Resolve-Servers([array]$servers,[string]$src,[string]$tgt){ 
    $src=Normalize-VersionString $src; 
    $tgt=Normalize-VersionString $tgt; 
    $map=$servers|ForEach-Object{[pscustomobject]@{Name=$_.Name;Raw=$_.AdminDisplayVersion.ToString();Ver=Normalize-VersionString ($_.AdminDisplayVersion.ToString())}}; 
    $source=$map|Where-Object{$_.Ver -eq $src}|Select-Object -First 1; 
    $target=$map|Where-Object{$_.Ver -eq $tgt}|Select-Object -First 1; 
    if(-not $source -or -not $target)
        { Write-Host ("Found versions: "+(($map|ForEach-Object{"{0}={1}" -f $_.Name,$_.Ver}) -join ', ')) -ForegroundColor DarkYellow; return $null }; 
        [pscustomobject]@{Source=$source.Name;Target=$target.Name} 
}

function Get-DatabaseInfo([string]$serverName){ 
    Get-MailboxDatabase -Server $serverName | Select-Object Name,@{N='EdbFilePath';E={[string]$_.EdbFilePath}},@{N='LogFolderPath';E={[string]$_.LogFolderPath}} 
}
function New-DbNameFromOld([string]$oldName){ 
    if($oldName -match '^(.*?)(\d+)$'){ $p=$matches[1];$d=$matches[2];$w=$d.Length;$n=[int]$d+1; return $p + $n.ToString('D'+$w)} 
    else { return $oldName+'1' } 
}
function Apply-DriveMap([string]$path){ 
    if([string]::IsNullOrWhiteSpace($path)){return $null}; 
    if(-not $DriveMap){return $path}; $drive=($path -split '(?<=:)',2)[0]; 
    if($DriveMap.ContainsKey($drive)){ return $DriveMap[$drive] + $path.Substring(2)}; return $path 
}
function Build-TargetPaths([string]$dbName,[string]$oldEdb,[string]$oldLog){ 
    if($PathMode -eq 'MirrorSource'){ 
        $oldDbFolder=$null; 
        if($oldEdb){ try{$oldDbFolder=Split-Path -Path $oldEdb -Parent}catch{} 
        } 
        $dbBase=if($oldDbFolder){ 
            try{ Split-Path -Path $oldDbFolder -Parent}
            catch{ $null } 
        } 
        else { $null }; 
        $dbBase=Apply-DriveMap $dbBase; 
        if([string]::IsNullOrWhiteSpace($dbBase)){ $dbBase=$DbBase }; 
        $newDbFolder=Join-Path -Path $dbBase -ChildPath $dbName; 
        $newEdbFile=Join-Path -Path $newDbFolder -ChildPath ("{0}.edb" -f $dbName); 
        $newLogFolder=$null; 
        if($oldLog -and $oldDbFolder -and $oldLog.StartsWith($oldDbFolder,[System.StringComparison]::OrdinalIgnoreCase)){ 
            $rel=$oldLog.Substring($oldDbFolder.Length); if([string]::IsNullOrWhiteSpace($rel)){ $rel='\\Logs' }; 
            $newLogFolder=$newDbFolder+$rel 
        } 
        else { $logBase=$null; 
            if($oldLog){ 
                try{ $logBase=Split-Path -Path (Split-Path -Path $oldLog -Parent) -Parent }
                catch{} 
            }; 
            $logBase=Apply-DriveMap $logBase; 
            if([string]::IsNullOrWhiteSpace($logBase)){ $logBase=$LogBase 
            }; 
            $newLogFolder=Join-Path -Path $logBase -ChildPath $dbName 
        }; 
    return [pscustomobject]@{DbFolder=$newDbFolder;LogFolder=$newLogFolder;EdbFile=$newEdbFile} 
    } 
    else { 
        $dbFolder=Join-Path -Path $DbBase -ChildPath $dbName; 
        $logFolder=Join-Path -Path $LogBase -ChildPath $dbName; 
        $edbFile=Join-Path -Path $dbFolder -ChildPath ("{0}.edb" -f $dbName); 
        return [pscustomobject]@{DbFolder=$dbFolder;LogFolder=$logFolder;EdbFile=$edbFile} 
    } 
}

function Invoke-Remote([string]$server,[scriptblock]$block){ 
    Invoke-Command -ComputerName $server -ErrorAction Stop -ScriptBlock $block 
}
function Test-RemoteFolder([string]$server,[string]$path){ 
    if([string]::IsNullOrWhiteSpace($path)){return $false}; 
    $p=[string]$path; 
    try{ Invoke-Remote -server $server -block { 
        param() $pp=$using:p; 
        if([string]::IsNullOrWhiteSpace($pp)){return $false}; 
        Test-Path -LiteralPath $pp -ErrorAction SilentlyContinue 
        } 
    }
    catch{ return $false } 
}
function Ensure-RemoteFolder([string]$server,[string]$path){ 
    if([string]::IsNullOrWhiteSpace($path)){return $false}; 
    $p=[string]$path; 
    try{ Invoke-Remote -server $server -block { 
        param() $pp=$using:p; 
        if([string]::IsNullOrWhiteSpace($pp)){return $false}; 
        if(-not (Test-Path -LiteralPath $pp -ErrorAction SilentlyContinue)){ 
            New-Item -Path $pp -ItemType Directory | Out-Null 
        }; 
        $true 
        } 
    }
    catch{ return $false } 
}
function Prepare-TargetFolders([string]$targetServer,[pscustomobject]$paths){ 
    if($PSCmdlet.ShouldProcess($paths.DbFolder,'Ensure remote DB folder')){[void](Ensure-RemoteFolder -server $targetServer -path $paths.DbFolder)}; 
    if($PSCmdlet.ShouldProcess($paths.LogFolder,'Ensure remote LOG folder')){[void](Ensure-RemoteFolder -server $targetServer -path $paths.LogFolder)} 
}

function Ensure-UniqueDbName([string]$name){ 
    -not [bool](Get-MailboxDatabase -Identity $name -ErrorAction SilentlyContinue) 
}
function Create-NewMailboxDatabase([string]$targetServer,[string]$dbName,[pscustomobject]$paths){ 
    if(-not (Ensure-UniqueDbName -name $dbName)){ 
        Write-Warning "Database '$dbName' already exists. Skipping creation."; return 
    }; 
    Prepare-TargetFolders -targetServer $targetServer -paths $paths; 
    $params=@{Name=$dbName;Server=$targetServer;EdbFilePath=$paths.EdbFile;LogFolderPath=$paths.LogFolder}; 
    if($ExcludeFromProvisioning){$params['IsExcludedFromProvisioning']=$true}; 
    if($PSCmdlet.ShouldProcess($dbName,'New-MailboxDatabase')){ 
        New-MailboxDatabase @params | Out-Null; Write-Host "Created DB '$dbName' on '$targetServer' (EDB=$($paths.EdbFile); LOG=$($paths.LogFolder))." -ForegroundColor Green; 
        if($MountAfterCreate){ 
            if($PSCmdlet.ShouldProcess($dbName,'Mount-Database')){ 
                try{ Mount-Database -Identity $dbName | Out-Null 
                }
                catch{ Write-Warning "Mount-Database failed for '$dbName': $($_.Exception.Message)" } 
            } 
        } 
    } 
}

$SettingsMap=@{ 
    Quotas=@('IssueWarningQuota','ProhibitSendQuota','ProhibitSendReceiveQuota','RecoverableItemsWarningQuota','RecoverableItemsQuota'); 
    Retention=@('DeletedItemRetention','RetainDeletedItemsUntilBackup','MailboxRetention'); 
    Client=@('MountAtStartup','IndexEnabled','IsExcludedFromProvisioning'); 
    Maintenance=@('CircularLoggingEnabled','BackgroundDatabaseMaintenance','QuotaNotificationSchedule','MaintenanceSchedule'); 
    OAB=@('OfflineAddressBook') 
}
function Expand-SettingsList([string[]]$scope){ 
    if(-not $scope -or $scope -contains 'All'){ 
        return ($SettingsMap.Values|ForEach-Object{$_}) 
    }; 
    $list=@(); foreach($s in $scope){ 
        if($SettingsMap.ContainsKey($s)){ 
            $list+=$SettingsMap[$s] 
        } 
    }; 
    $list|Select-Object -Unique 
}

function Convert-Quota([object]$q){ if($null -eq $q){ return $null }
    try{ if($q.PSObject.TypeNames -contains 'Microsoft.Exchange.Data.Unlimited`1[[Microsoft.Exchange.Data.ByteQuantifiedSize]]' -or $q.GetType().FullName -like 'Microsoft.Exchange.Data.Unlimited`1*ByteQuantifiedSize*'){ if($q.IsUnlimited){ return 'Unlimited' }; $s=[string]$q.Value } else { $s=[string]$q } }catch{ $s=[string]$q }
    if($s -match '^[Uu]nlimited$'){ return 'Unlimited' }
    if($s -match '\((?<bytes>[\d\.,\s]+)\s*bytes\)'){ $raw=$matches['bytes'] -replace '[^\d]',''; if($raw){ $bytes=[double]$raw; if($bytes -gt 0){ $mb=[math]::Round($bytes/1MB); if($mb -lt 1){ $mb=1 }; return ("{0} MB" -f $mb) } } }
    if($s -match '^(?<size>\d+[\.,]?\d*)\s*(?<unit>B|KB|MB|GB|TB)'){ $num=$matches['size'].Replace(',','.'); $val=[double]::Parse($num,[Globalization.CultureInfo]::InvariantCulture); switch($matches['unit'].ToUpper()){ 'TB'{$mb=[math]::Round($val*1024*1024)} 'GB'{$mb=[math]::Round($val*1024)} 'MB'{$mb=[math]::Round($val)} 'KB'{$mb=[math]::Round($val/1024)} default{$mb=[math]::Round($val/1MB)} }; if($mb -lt 1){ $mb=1 }; return ("{0} MB" -f $mb) }
    return $s 
}
function Convert-TimeSpanLike([object]$t){ 
    if($null -eq $t){ return $null }; return [string]$t 
}

function Trim-OABDisplay([string]$s){ 
    if([string]::IsNullOrWhiteSpace($s)){ 
        return $s 
    }; 
    return ($s.Trim()).TrimStart('\') 
}
function Get-OABInfo([object]$val){ $raw = if ($null -eq $val) { $null } else { [string]$val }; $trim = Trim-OABDisplay $raw; $info = [pscustomobject]@{ Raw=$raw; RawTrim=$trim; Name=$null; Dn=$null; Guid=$null; IsNull=([string]::IsNullOrWhiteSpace($raw)); Resolved=$false }; if ($info.IsNull) { return $info }
    try{ $oab=$null; if($trim -like 'CN=*,CN=Offline Address Lists*'){ $oab=Get-OfflineAddressBook -Identity $trim -ErrorAction SilentlyContinue }
         if(-not $oab -and $raw -like 'CN=*,CN=Offline Address Lists*'){ $oab=Get-OfflineAddressBook -Identity $raw -ErrorAction SilentlyContinue }
         if(-not $oab){ $oab=Get-OfflineAddressBook -Identity $trim -ErrorAction SilentlyContinue }
         if(-not $oab){ $oab=Get-OfflineAddressBook -Identity $raw -ErrorAction SilentlyContinue }
         if($oab){ $info.Name=$oab.Name; $info.Dn=$oab.DistinguishedName; $info.Guid=$oab.Guid; $info.Resolved=$true } }catch{}
    return $info 
}
function Normalize-DbOABValue([object]$val){ 
    try{ 
        if($val -and $val.DistinguishedName){ 
            return Get-OABInfo $val.DistinguishedName 
        } 
    }catch{}; 
    return Get-OABInfo $val 
}

function Get-DbSettings([string]$dbName,[string[]]$properties){ 
    $db=Get-MailboxDatabase -Identity $dbName -ErrorAction Stop; 
    $ht=@{}; 
    foreach($p in $properties){ 
        try{$ht[$p]=$db.$p 
        }
        catch{$ht[$p]=$null} 
    }; [pscustomobject]$ht | Select-Object $properties 
}

function Compare-DbSettings([string]$sourceDb,[string]$targetDb,[string[]]$properties){ $src=Get-DbSettings -dbName $sourceDb -properties $properties; $tgt=Get-DbSettings -dbName $targetDb -properties $properties -ErrorAction SilentlyContinue; if(-not $tgt){ Write-Warning "Target DB '$targetDb' does not exist. Create it first."; return @() }
    $diff=@(); foreach($p in $properties){ if($p -eq 'OfflineAddressBook'){
        $srcI=Normalize-DbOABValue $src.$p; $tgtI=Normalize-DbOABValue $tgt.$p
        if($DebugOAB){ Write-Host ("OAB DEBUG [{0}]: SRC Raw='{1}' Trim='{2}' IsNull={3} Resolved={4} Name='{5}' DN='{6}' GUID='{7}'" -f $targetDb,$srcI.Raw,$srcI.RawTrim,$srcI.IsNull,$srcI.Resolved,$srcI.Name,$srcI.Dn,$srcI.Guid) -ForegroundColor DarkCyan; Write-Host ("OAB DEBUG [{0}]: TGT Raw='{1}' Trim='{2}' IsNull={3} Resolved={4} Name='{5}' DN='{6}' GUID='{7}'" -f $targetDb,$tgtI.Raw,$tgtI.RawTrim,$tgtI.IsNull,$tgtI.Resolved,$tgtI.Name,$tgtI.Dn,$tgtI.Guid) -ForegroundColor DarkCyan }
        $same=$false; $reason=''
        if($srcI.IsNull -and $tgtI.IsNull){ $same=$true; $reason='both null' }
        elseif( ($srcI.IsNull -and -not $tgtI.IsNull) -or (-not $srcI.IsNull -and $tgtI.IsNull) ){
            $same=$false; $reason='XOR null/non-null'
        } else {
            if($srcI.Guid -and $tgtI.Guid){ $same=($srcI.Guid -eq $tgtI.Guid); $reason='GUID compare' }
            elseif($srcI.Dn -and $tgtI.Dn){ $same=($srcI.Dn -ieq $tgtI.Dn); $reason='DN compare' }
            else { $same=( ($srcI.RawTrim -ieq $tgtI.RawTrim) -or ( ($srcI.Name -ne $null) -and ($tgtI.Name -ne $null) -and ($srcI.Name -ieq $tgtI.Name) ) ); $reason='Name/Raw compare' }
        }
        if($DebugOAB){ Write-Host ("OAB DEBUG [{0}]: SAME={1} REASON={2}" -f $targetDb,$same,$reason) -ForegroundColor DarkYellow }
        if(-not $same){ $srcDisp = if ($srcI.Name) { $srcI.Name } else { $srcI.RawTrim }; $tgtDisp = if ($tgtI.Name) { $tgtI.Name } else { $tgtI.RawTrim }; $diff += [pscustomobject]@{ Database=$targetDb; Property=$p; SourceValue=$srcDisp; TargetValue=$tgtDisp; SourceIdentity=$srcI.Dn; TargetIdentity=$tgtI.Dn; SourceRaw=$srcI.Raw; TargetRaw=$tgtI.Raw } }
    } else { $sv=$src.$p; $tv=$tgt.$p; if(([string]$sv) -ne ([string]$tv)){ $diff += [pscustomobject]@{ Database=$targetDb; Property=$p; SourceValue=$sv; TargetValue=$tv } } } }
    return ,$diff 
}

function Convert-OABForSet([object]$oab){ 
    $i=Normalize-DbOABValue $oab; 
    if($i.Dn){ return $i.Dn 
    } 
    elseif($i.Name){ 
        return $i.Name 
    } 
    elseif($i.RawTrim){ 
        return $i.RawTrim 
    } 
    else { return $null } 
}
function Convert-ToExchangeValue([string]$prop,$val){ switch($prop){
    'IssueWarningQuota' { return Convert-Quota $val }
    'ProhibitSendQuota' { return Convert-Quota $val }
    'ProhibitSendReceiveQuota' { return Convert-Quota $val }
    'RecoverableItemsWarningQuota' { return Convert-Quota $val }
    'RecoverableItemsQuota' { return Convert-Quota $val }
    'DeletedItemRetention' { return Convert-TimeSpanLike $val }
    'MailboxRetention' { return Convert-TimeSpanLike $val }
    'RetainDeletedItemsUntilBackup' { return [bool]$val }
    'MountAtStartup' { return [bool]$val }
    'IndexEnabled' { return [bool]$val }
    'IsExcludedFromProvisioning' { return [bool]$val }
    'CircularLoggingEnabled' { return [bool]$val }
    'BackgroundDatabaseMaintenance' { return [bool]$val }
    'QuotaNotificationSchedule' { return [string]$val }
    'MaintenanceSchedule' { return [string]$val }
    'OfflineAddressBook' { return Convert-OABForSet $val }
    default { return $val } } }

function Apply-DbSettings([string]$targetDb,[object[]]$diffRows){ 
    if(-not $diffRows -or $diffRows.Count -eq 0){ 
        return 
    }; 
    $params=@{ Identity=$targetDb }; 
    foreach($d in $diffRows){ 
        if($d.Property -eq 'OfflineAddressBook' -and ([string]::IsNullOrWhiteSpace($d.SourceValue))){ 
            $params['OfflineAddressBook']=$null 
        } 
        else { 
            $san=Convert-ToExchangeValue -prop $d.Property -val $d.SourceValue; if($null -ne $san -and $san -ne ''){ 
                $params[$d.Property]=$san 
            } 
        } 
    }; 
    if($params.Keys.Count -gt 1){ 
        if($PSCmdlet.ShouldProcess($targetDb,'Set-MailboxDatabase (apply settings)')){ 
            Set-MailboxDatabase @params 
        } 
    } 
}

function Queue-ArbitrationMigrationBatch([string]$batchPrefix,[string]$NotificationEmails,[int]$BadItemLimit,[switch]$SplitArbitrationBySourceDb,[string]$TargetDb,[switch]$AutoStart,[switch]$KeepCsv,[string]$ForceTargetDb){
    $arbMailboxes=Get-Mailbox -Arbitration | Where-Object { $_.ServerName -ne $pair.Target };
    if($arbMailboxes.Count -eq 0){ Write-Host "No arbitration mailboxes found." -ForegroundColor DarkYellow; return }
    $targetDb=if($ForceTargetDb){$ForceTargetDb}else{$TargetDb}
    $batchName=("{0}_Arbitration_{1:yyyyMMdd-HHmm}" -f $batchPrefix,(Get-Date));
    $csvPath=Join-Path $env:TEMP ("{0}.csv" -f $batchName);
    $arbMailboxes | Select-Object @{Name='EmailAddress';Expression={$_.PrimarySmtpAddress}} | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8;
    $csvBytes=[System.IO.File]::ReadAllBytes($csvPath);
    $params=@{Name=$batchName;Local=$true;CSVData=$csvBytes;TargetDatabases=$targetDb;ArchiveTargetDatabases=$targetDb;BadItemLimit=$BadItemLimit};
    if($NotificationEmails){$params['NotificationEmails']=$NotificationEmails};
    try{
        New-MigrationBatch @params | Out-Null;
        if($AutoStart){ Start-MigrationBatch -Identity $batchName | Out-Null }
        Write-Host ("Created MigrationBatch '{0}' for {1} arbitration mailboxes -> '{2}'. AutoStart={3}" -f $batchName,$arbMailboxes.Count,$targetDb,$AutoStart) -ForegroundColor Cyan
    } catch{ Write-Warning "Failed to create MigrationBatch '$batchName': $($_.Exception.Message)" }
    if(-not $KeepCsv){ Remove-Item $csvPath -Force } else { Write-Host "CSV retained at $csvPath" -ForegroundColor Yellow }
}
function Queue-RegularMigrationBatches([string]$batchPrefix,[string]$NotificationEmails,[int]$BadItemLimit,[switch]$AutoStart,[switch]$KeepCsv,[string]$FallbackTargetDb,[string]$ForceTargetDb){
    $mailboxes=Get-Mailbox -ResultSize Unlimited | Where-Object { $_.RecipientTypeDetails -in @('UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox') -and $_.ServerName -ne $pair.Target };
    if($mailboxes.Count -eq 0){ Write-Host "No regular mailboxes found." -ForegroundColor DarkYellow; return }
    $map=@();
    foreach($mbx in $mailboxes){
        $targetDb=if($ForceTargetDb){$ForceTargetDb}else{($plan | Where-Object { $_.SourceDb -eq $mbx.Database }).TargetDb};
        if(-not $targetDb){ $targetDb=$FallbackTargetDb }
        $map += [pscustomobject]@{ Mailbox=$mbx; Type=$mbx.RecipientTypeDetails; TargetDb=$targetDb }
    }
    $groups=$map | Group-Object @{Expression={ "$($_.Type)-$($_.TargetDb)" }};
    foreach($grp in $groups){
        if($grp.Count -eq 0){ continue }
        $parts=$grp.Name -split '-';
        $type=$parts[0];
        $targetDb=$parts[1];
        $safeType=$type -replace '[^a-zA-Z0-9]','';
        $safeTarget=$targetDb -replace '[^a-zA-Z0-9]','';
        $batchName=("{0}_{1}_{2}_{3:yyyyMMdd-HHmm}" -f $batchPrefix,$safeType,$safeTarget,(Get-Date));
        $csvPath=Join-Path $env:TEMP ("{0}.csv" -f $batchName);
        $grp.Group | ForEach-Object { $_.Mailbox } | Select-Object @{Name='EmailAddress';Expression={$_.PrimarySmtpAddress}} | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8;
        $csvBytes=[System.IO.File]::ReadAllBytes($csvPath);
        $params=@{Name=$batchName;Local=$true;CSVData=$csvBytes;TargetDatabases=$targetDb;ArchiveTargetDatabases=$targetDb;BadItemLimit=$BadItemLimit};
        if($NotificationEmails){$params['NotificationEmails']=$NotificationEmails};
        try{
            New-MigrationBatch @params | Out-Null;
            if($AutoStart){ Start-MigrationBatch -Identity $batchName | Out-Null }
            Write-Host ("Created MigrationBatch '{0}' for {1} mailboxes ({2}) -> '{3}'. AutoStart={4}" -f $batchName,$grp.Count,$type,$targetDb,$AutoStart) -ForegroundColor Cyan
        } catch{ Write-Warning "Failed to create MigrationBatch '$batchName': $($_.Exception.Message)" }
        if(-not $KeepCsv){ Remove-Item $csvPath -Force } else { Write-Host "CSV retained at $csvPath" -ForegroundColor Yellow }
    }
}

if(-not (Test-ExchangeEnvironment)){ return }


function Test-IsAbsoluteHttpUrl { 
    param([string]$Url) 
    if ([string]::IsNullOrWhiteSpace($Url)) { 
        return $false 
    } 
    try { $u=[Uri]$Url; 
        return ($u.Scheme -in @('http','https') -and $u.IsAbsoluteUri) 
    } 
    catch { 
        return $false 
    } 
}
function Get-HostFromUrl { 
    param([string]$Url) 
    try { 
        return ([Uri]$Url).Host 
    } 
    catch { 
        return $null 
    } 
}
function Test-ServerBoundUrl { 
    param([string]$Url,[string]$SourceServerName) 
    if ([string]::IsNullOrWhiteSpace($Url) -or [string]::IsNullOrWhiteSpace($SourceServerName)) { 
        return $false 
    } 
    $urlHost=Get-HostFromUrl -Url $Url; 
    if (-not $urlHost){
        return $false
    } 
    $short=($SourceServerName -split '\.')[0]; 
    return ($urlHost -ieq $SourceServerName -or $urlHost -ieq $short -or $urlHost -like ($short+'.*')) 
}

function Get-CasUrlSnapshot { 
    param([string]$Server) 
    $snap=[ordered]@{
        Server=$Server;
        AutoDiscoverServiceInternalUri=$null;
        OWA_InternalUrl=$null;
        OWA_ExternalUrl=$null;
        ECP_InternalUrl=$null;
        ECP_ExternalUrl=$null;
        EWS_InternalUrl=$null;
        EWS_ExternalUrl=$null;
        EAS_InternalUrl=$null;
        EAS_ExternalUrl=$null;
        OAB_InternalUrl=$null;
        OAB_ExternalUrl=$null;
        MAPI_InternalUrl=$null;
        MAPI_ExternalUrl=$null;
        OutlookAnywhere_InternalHostname=$null;
        OutlookAnywhere_ExternalHostname=$null
    } 
    try{
        $cas=Get-ClientAccessService -Identity $Server;
        $snap.AutoDiscoverServiceInternalUri=[string]$cas.AutoDiscoverServiceInternalUri
    }
    catch{}
    try{$x=Get-OwaVirtualDirectory -Server $Server;
        $snap.OWA_InternalUrl=[string]$x.InternalUrl;
        $snap.OWA_ExternalUrl=[string]$x.ExternalUrl
    }
    catch{} 
    try{$x=Get-EcpVirtualDirectory -Server $Server;
        $snap.ECP_InternalUrl=[string]$x.InternalUrl;
        $snap.ECP_ExternalUrl=[string]$x.ExternalUrl
    }
    catch{} 
    try{$x=Get-WebServicesVirtualDirectory -Server $Server;
        $snap.EWS_InternalUrl=[string]$x.InternalUrl;
        $snap.EWS_ExternalUrl=[string]$x.ExternalUrl
    }
    catch{} 
    try{$x=Get-ActiveSyncVirtualDirectory -Server $Server;
        $snap.EAS_InternalUrl=[string]$x.InternalUrl;
        $snap.EAS_ExternalUrl=[string]$x.ExternalUrl
    }
    catch{} 
    try{$x=Get-OabVirtualDirectory -Server $Server;
        $snap.OAB_InternalUrl=[string]$x.InternalUrl;
        $snap.OAB_ExternalUrl=[string]$x.ExternalUrl
    }
    catch{} 
    try{$x=Get-MapiVirtualDirectory -Server $Server;
        $snap.MAPI_InternalUrl=[string]$x.InternalUrl;
        $snap.MAPI_ExternalUrl=[string]$x.ExternalUrl
    }
    catch{} 
    try{$x=Get-OutlookAnywhere -Server $Server;
        $snap.OutlookAnywhere_InternalHostname=[string]$x.InternalHostname;
        $snap.OutlookAnywhere_ExternalHostname=[string]$x.ExternalHostname
    }
    catch{} 
    [pscustomobject]$snap 
}

function Compare-CasUrls { 
<##
.SYNOPSIS
    Vergleicht CAS-/VD-URLs zwischen Quelle und Ziel.
#>
    param(
        [psobject]$Source,
        [psobject]$Target,
        [switch]$CasShowAll,
[int]$BadItemLimit = 0,
[string]$NotifyEmail,
[switch]$Arbitration,
[switch]$AcceptLargeDataLoss,
        [string]$SourceServerName,
        [string]$TargetServerName
        ) 
        $props='AutoDiscoverServiceInternalUri','OWA_InternalUrl','OWA_ExternalUrl','ECP_InternalUrl','ECP_ExternalUrl','EWS_InternalUrl','EWS_ExternalUrl','EAS_InternalUrl','EAS_ExternalUrl','OAB_InternalUrl','OAB_ExternalUrl','MAPI_InternalUrl','MAPI_ExternalUrl','OutlookAnywhere_InternalHostname','OutlookAnywhere_ExternalHostname';
        $rows=@();foreach($p in $props){
            $sv=[string]$Source.$p;
            $tv=[string]$Target.$p;
            $equal=($sv -eq $tv);
            $bothEmpty=([string]::IsNullOrWhiteSpace($sv)-and[string]::IsNullOrWhiteSpace($tv));
            $note=$null;
            if($p -like '*Url'-and(Test-ServerBoundUrl -Url $sv -SourceServerName $SourceServerName)){
                $note='ServerBound(Source) – manual adjust recommended'
            }
            if(-not $equal -or $CasShowAll){
                $status=if(-not $equal){
                    'Different'
                }
                elseif($bothEmpty){
                    'Equal (BothEmpty)'
                }
                else{'Equal'};
                $rows+=[pscustomobject]@{
                    Property=$p;
                    SourceValue=$sv;
                    TargetValue=$tv;
                    Status=$status;
                    Note=$note
                }
            }
        };
    ,$rows 
}

function Apply-CasUrls { 
<##
.SYNOPSIS
    Übernimmt CAS-/VD-URLs vom Snapshot auf den Zielserver (nur valide, nicht servergebunden).
#>
    param(
        [string]$TargetServer,
        [psobject]$SourceSnapshot
        ) 
    # OWA
    if(Test-IsAbsoluteHttpUrl $SourceSnapshot.OWA_InternalUrl -and -not(Test-ServerBoundUrl -Url $SourceSnapshot.OWA_InternalUrl -SourceServerName $TargetServer)){
        Get-OwaVirtualDirectory -Server $TargetServer|Set-OwaVirtualDirectory -InternalUrl $SourceSnapshot.OWA_InternalUrl -ExternalUrl $SourceSnapshot.OWA_ExternalUrl -Confirm:$false
    }
    else{
        Write-Host "Skip OWA InternalUrl (invalid or serverbound): '$($SourceSnapshot.OWA_InternalUrl)'" -ForegroundColor DarkYellow
    } 
    # ECP
    if(Test-IsAbsoluteHttpUrl $SourceSnapshot.ECP_InternalUrl -and -not(Test-ServerBoundUrl -Url $SourceSnapshot.ECP_InternalUrl -SourceServerName $TargetServer)){
        Get-EcpVirtualDirectory -Server $TargetServer|Set-EcpVirtualDirectory -InternalUrl $SourceSnapshot.ECP_InternalUrl -ExternalUrl $SourceSnapshot.ECP_ExternalUrl -Confirm:$false
    }
    else{
        Write-Host "Skip ECP InternalUrl (invalid or serverbound): '$($SourceSnapshot.ECP_InternalUrl)'" -ForegroundColor DarkYellow
    }
    # EWS
    if(Test-IsAbsoluteHttpUrl $SourceSnapshot.EWS_InternalUrl -and -not(Test-ServerBoundUrl -Url $SourceSnapshot.EWS_InternalUrl -SourceServerName $TargetServer)){
        Get-WebServicesVirtualDirectory -Server $TargetServer|Set-WebServicesVirtualDirectory -InternalUrl $SourceSnapshot.EWS_InternalUrl -ExternalUrl $SourceSnapshot.EWS_ExternalUrl -Confirm:$false
    }
    else{
        Write-Host "Skip EWS InternalUrl (invalid or serverbound): '$($SourceSnapshot.EWS_InternalUrl)'" -ForegroundColor DarkYellow
    } 
    # EAS
    if(Test-IsAbsoluteHttpUrl $SourceSnapshot.EAS_InternalUrl -and -not(Test-ServerBoundUrl -Url $SourceSnapshot.EAS_InternalUrl -SourceServerName $TargetServer)){
        Get-ActiveSyncVirtualDirectory -Server $TargetServer|Set-ActiveSyncVirtualDirectory -InternalUrl $SourceSnapshot.EAS_InternalUrl -ExternalUrl $SourceSnapshot.EAS_ExternalUrl -Confirm:$false
    }
    else{
        Write-Host "Skip EAS InternalUrl (invalid or serverbound): '$($SourceSnapshot.EAS_InternalUrl)'" -ForegroundColor DarkYellow
    }
    # OAB
    if(Test-IsAbsoluteHttpUrl $SourceSnapshot.OAB_InternalUrl -and -not(Test-ServerBoundUrl -Url $SourceSnapshot.OAB_InternalUrl -SourceServerName $TargetServer)){
        Get-OabVirtualDirectory -Server $TargetServer|Set-OabVirtualDirectory -InternalUrl $SourceSnapshot.OAB_InternalUrl -ExternalUrl $SourceSnapshot.OAB_ExternalUrl -Confirm:$false
    }
    else{
        Write-Host "Skip OAB InternalUrl (invalid or serverbound): '$($SourceSnapshot.OAB_InternalUrl)'" -ForegroundColor DarkYellow
    } 
    # MAPI
    if(Test-IsAbsoluteHttpUrl $SourceSnapshot.MAPI_InternalUrl -and -not(Test-ServerBoundUrl -Url $SourceSnapshot.MAPI_InternalUrl -SourceServerName $TargetServer)){
        Get-MapiVirtualDirectory -Server $TargetServer|Set-MapiVirtualDirectory -InternalUrl $SourceSnapshot.MAPI_InternalUrl -ExternalUrl $SourceSnapshot.MAPI_ExternalUrl -Confirm:$false
    }
    else{
        Write-Host "Skip MAPI InternalUrl (invalid or serverbound): '$($SourceSnapshot.MAPI_InternalUrl)'" -ForegroundColor DarkYellow
    } 
    # Outlook Anywhere
    $ext=$SourceSnapshot.OutlookAnywhere_ExternalHostname;$int=$SourceSnapshot.OutlookAnywhere_InternalHostname;if(-not[string]::IsNullOrWhiteSpace($ext)-or -not[string]::IsNullOrWhiteSpace($int)){
        Get-OutlookAnywhere -Server $TargetServer|Set-OutlookAnywhere -ExternalHostname $ext -InternalHostname $int -ExternalClientsRequireSsl:$true -InternalClientsRequireSsl:$true -ExternalClientAuthenticationMethod 'Negotiate' -Confirm:$false
    } 
    # AutoDiscover
    $ad=$SourceSnapshot.AutoDiscoverServiceInternalUri;
    if(Test-IsAbsoluteHttpUrl $ad){
        Get-ClientAccessService -Identity $TargetServer|Set-ClientAccessService -AutoDiscoverServiceInternalUri $ad -Confirm:$false
    }
    else{
        Write-Host "Skip AutoDiscoverServiceInternalUri (invalid/empty): '$ad'" -ForegroundColor DarkYellow
    }
}

$servers=Get-ExchangeServerVersions
$pair=Resolve-Servers -servers $servers -src $SourceVersion -tgt $TargetVersion

# CAS URL Compare/Apply
if($CompareCasUrls -or $ApplyCasUrls){
    $srcCas=if($PSBoundParameters.ContainsKey('SourceCasServer')-and$SourceCasServer){
        $SourceCasServer
    }else{
        $pair.Source
    };
    $tgtCas=if($PSBoundParameters.ContainsKey('TargetCasServer')-and$TargetCasServer){
        $TargetCasServer
    }
    else{
        $pair.Target
    };Write-Host "`nCAS URL SNAPSHOTS:" -ForegroundColor Yellow;$srcSnap=Get-CasUrlSnapshot -Server $srcCas;$tgtSnap=Get-CasUrlSnapshot -Server $tgtCas;$casDiffs=Compare-CasUrls -Source $srcSnap -Target $tgtSnap -SourceServerName $srcCas -TargetServerName $tgtCas -CasShowAll:$CasShowAll;
    if(@($casDiffs).Count -gt 0){$casDiffs|Format-Table Property,SourceValue,TargetValue,Status,Note -AutoSize|Out-Host;
        if($Approve -and $ApplyCasUrls){
            Apply-CasUrls -TargetServer $tgtCas -SourceSnapshot $srcSnap;Write-Host "Applied CAS/VD URLs from '$srcCas' to '$tgtCas'." -ForegroundColor Green
        }
    }
    else{
        Write-Host "No CAS/VD URL differences between '$srcCas' and '$tgtCas'." -ForegroundColor DarkGreen
    };
    return
}
if(-not $pair){ Write-Error "Could not find both required versions: source=$SourceVersion, target=$TargetVersion."; return }
Write-Host ("Environment OK. Source='{0}' Target='{1}'" -f $pair.Source,$pair.Target) -ForegroundColor Green

$sourceDbs=Get-DatabaseInfo -serverName $pair.Source
if($IncludeDbs -and $IncludeDbs.Count -gt 0){ $sourceDbs=$sourceDbs | Where-Object { $IncludeDbs -contains $_.Name } }

$plan = foreach($db in $sourceDbs){ $newName=New-DbNameFromOld -oldName $db.Name; $paths=Build-TargetPaths -dbName $newName -oldEdb $db.EdbFilePath -oldLog $db.LogFolderPath; $dbOk=Test-RemoteFolder -server $pair.Target -path $paths.DbFolder; $logOk=Test-RemoteFolder -server $pair.Target -path $paths.LogFolder; [pscustomobject]@{ SourceDb=$db.Name; OldEdb=$db.EdbFilePath; OldLog=$db.LogFolderPath; TargetDb=$newName; TargetEdb=$paths.EdbFile; TargetDbDir=$paths.DbFolder; TargetLog=$paths.LogFolder; DbDirExists=[bool]$dbOk; LogDirExists=[bool]$logOk } }

Write-Host "`nPLAN (review before execution):" -ForegroundColor Yellow
$plan | Select SourceDb,OldEdb,OldLog,TargetDb,TargetEdb,TargetDbDir,TargetLog,DbDirExists,LogDirExists | Format-Table -AutoSize | Out-Host
if($ExportPlan){ try{ $plan | Export-Csv -Path $ExportPlan -NoTypeInformation -Encoding UTF8; Write-Host "Plan exported to $ExportPlan" -ForegroundColor DarkGreen }catch{ Write-Warning "Export failed: $($_.Exception.Message)" } }

if($Interactive){ Write-Host "`nInteractive mode: edit target paths (ENTER to keep)." -ForegroundColor Cyan; $adjusted=@(); foreach($item in $plan){ Write-Host "`nDB: $($item.SourceDb) -> $($item.TargetDb)" -ForegroundColor Cyan; $nd=Read-Host ("Target DB folder   [{0}]" -f $item.TargetDbDir); $nl=Read-Host ("Target LOG folder  [{0}]" -f $item.TargetLog); $ne=Read-Host ("Target EDB file    [{0}]" -f $item.TargetEdb); if($nd){$item.TargetDbDir=$nd}; if($nl){$item.TargetLog=$nl}; if($ne){$item.TargetEdb=$ne}; $adjusted+=$item }; $plan=$adjusted; Write-Host "`nAdjusted PLAN:" -ForegroundColor Yellow; $plan | Select SourceDb,TargetDb,TargetEdb,TargetDbDir,TargetLog | Format-Table -AutoSize | Out-Host }

if(-not $Approve -and -not $CompareSettings){
    if($QueueMoves){
        if($Arbitration){
            Write-Host "`nDRY-RUN: Arbitration mailboxes that would be queued for move:" -ForegroundColor Cyan
            $arbMailboxes=Get-Mailbox -Arbitration
            foreach($arb in $arbMailboxes){
                $targetDb = ($plan | Where-Object { $_.SourceDb -eq $arb.Database }).TargetDb
                if(-not $targetDb){ $targetDb=$plan[0].TargetDb }
                Write-Host ("{0} (ArbitrationMailbox) -> {1}" -f $arb.DisplayName,$targetDb) -ForegroundColor Yellow
            }
        } else {
            Write-Host "`nDRY-RUN: Mailboxes that would be queued for move:" -ForegroundColor Cyan
            foreach($item in $plan){
                $mailboxes=Get-Mailbox -ResultSize Unlimited | Where-Object { $_.Database -eq $item.SourceDb }
                foreach($mbx in $mailboxes){
                    Write-Host ("{0} ({1}) -> {2}" -f $mbx.DisplayName,$mbx.RecipientTypeDetails,$item.TargetDb) -ForegroundColor Yellow
                }
            }
        }
    }
    Write-Host "`nNo changes executed. Re-run with -Approve and/or -CompareSettings when ready." -ForegroundColor DarkYellow; return
}

foreach($item in $plan){ $paths=[pscustomobject]@{ DbFolder=$item.TargetDbDir; LogFolder=$item.TargetLog; EdbFile=$item.TargetEdb }; if($Approve -and ($PrepareFolders -or $CreateDatabases)){ Prepare-TargetFolders -targetServer $pair.Target -paths $paths; if($CreateDatabases){ Create-NewMailboxDatabase -targetServer $pair.Target -dbName $item.TargetDb -paths $paths } } }

if($CompareSettings -or $ApplySettings){ $props=Expand-SettingsList -scope $SettingsScope; $allDiffs=@(); Write-Host "`nDB SETTINGS DIFFS:" -ForegroundColor Yellow; foreach($item in $plan){ $exists = -not (Ensure-UniqueDbName -name $item.TargetDb); if(-not $exists){ Write-Host "Target DB '$($item.TargetDb)' not found (create first)." -ForegroundColor DarkYellow; continue }; $diff=Compare-DbSettings -sourceDb $item.SourceDb -targetDb $item.TargetDb -properties $props; $diffList=@($diff); if($diffList.Count -gt 0){ $diffList | Format-Table * -AutoSize | Out-Host; $allDiffs += $diffList; if($Approve -and $ApplySettings){ Apply-DbSettings -targetDb $item.TargetDb -diffRows $diffList } } else { Write-Host "No differences for '$($item.TargetDb)'." -ForegroundColor DarkGreen } }; if($ExportDiffs -and @($allDiffs).Count -gt 0){ try{ $allDiffs | Export-Csv -Path $ExportDiffs -NoTypeInformation -Encoding UTF8; Write-Host "Diffs exported to $ExportDiffs" -ForegroundColor DarkGreen }catch{ Write-Warning "Export diffs failed: $($_.Exception.Message)" } } }

if($Approve -and $QueueMoves){
    if($Arbitration){
        $targetDb = ($plan | Where-Object { $_.SourceDb -like '*Verwaltung*' }).TargetDb;
        if(-not $targetDb){ $targetDb=$FallbackTargetDb }
        Queue-ArbitrationMigrationBatch -batchPrefix $BatchNamePrefix -NotificationEmails $NotifyEmail -BadItemLimit $BadItemLimit -TargetDb $targetDb -SplitArbitrationBySourceDb:$SplitArbitrationBySourceDb -AutoStart:$AutoStart -KeepCsv:$KeepCsv -ForceTargetDb $ForceTargetDb;
    } else {
        Queue-RegularMigrationBatches -batchPrefix $BatchNamePrefix -NotificationEmails $NotifyEmail -BadItemLimit $BadItemLimit -AutoStart:$AutoStart -KeepCsv:$KeepCsv -FallbackTargetDb $FallbackTargetDb -ForceTargetDb $ForceTargetDb;
    }
}
Write-Host "
Done." -ForegroundColor Green
