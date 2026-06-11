<#
.SYNOPSIS
Windows Backup Service Monitoring Script
Überwacht den Windows Backup Service und versendet E-Mails bei erfolgreichen und fehlgeschlagenen Backups

.DESCRIPTION
Dieses Script lauscht auf Events im Windows Backup Event Log und versendet E-Mails mit Details zum Backup-Status.
Es kann als Scheduled Task eingerichtet werden, um auf Backup-Events zu reagieren.

.PARAMETER SMTPServer
Der SMTP-Server für den E-Mail-Versand (erforderlich)

.PARAMETER SMTPPort
Der Port des SMTP-Servers (Standard: 25)

.PARAMETER From
Die Absender-E-Mail-Adresse (erforderlich)

.PARAMETER To
Die Empfänger-E-Mail-Adresse(n) (erforderlich)

.PARAMETER SMTPCredential
Die Anmeldedaten für den SMTP-Server (optional)

.PARAMETER UseSSL
SSL/TLS für SMTP verwenden (Standard: $false)

.PARAMETER CreateTask
Wenn gesetzt, wird der Scheduled Task erstellt und konfiguriert (Standard: $false)

.PARAMETER TaskName
Name des Scheduled Tasks (Standard: "Windows Backup Report Monitor")

.EXAMPLE
.\WindowsBackupReport.ps1 -SMTPServer "smtp.contoso.com" -From "backup@contoso.com" -To "admin@contoso.com"

.EXAMPLE
.\WindowsBackupReport.ps1 -CreateTask -SMTPServer "smtp.contoso.com" -From "backup@contoso.com" -To "admin@contoso.com"

.NOTES
Author: Admin Script
Version: 1.0
Created: March 2026
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$SMTPServer,
    
    [Parameter(Mandatory=$false)]
    [int]$SMTPPort = 25,
    
    [Parameter(Mandatory=$false)]
    [string]$From,
    
    [Parameter(Mandatory=$false)]
    [string]$To,
    
    [Parameter(Mandatory=$false)]
    [PSCredential]$SMTPCredential,
    
    [Parameter(Mandatory=$false)]
    [bool]$UseSSL = $false,
    
    [Parameter(Mandatory=$false)]
    [switch]$CreateTask,
    
    [Parameter(Mandatory=$false)]
    [string]$TaskName = "Windows Backup Report Monitor"
)

# ===========================
# Konfiguration
# ===========================
$ConfigFile = "$PSScriptRoot\WindowsBackupReport.config.xml"
$ScriptVersion = "1.0"
$LogPath = "C:\Logs\WindowsBackup"
$LogFile = Join-Path -Path $LogPath -ChildPath "WindowsBackupReport_$(Get-Date -Format 'yyyyMMdd').log"
$TranscriptFile = Join-Path -Path $LogPath -ChildPath "Transcript_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Event Log Constants
$EventLogName = "Microsoft-Windows-Backup"
$EventSourceName = "Microsoft-Windows-Backup"
$SuccessEventId = 4      # Backup erfolgreich abgeschlossen
$FailureEventId = 12     # Backup fehlgeschlagen
$WarningEventId = 8      # Backup mit Fehlern abgeschlossen

# ===========================
# Funktionen
# ===========================

function Initialize-LogPath {
    <#
    .SYNOPSIS
    Initialisiert das Log-Verzeichnis
    #>
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Initialisiere Log-Verzeichnis..." -ForegroundColor Cyan
    
    if (-not (Test-Path -Path $LogPath)) {
        try {
            New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Log-Verzeichnis erstellt: $LogPath" -ForegroundColor Green
        }
        catch {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] FEHLER: Konnte Log-Verzeichnis nicht erstellen: $_" -ForegroundColor Red
            exit 1
        }
    }
}

function Write-LogEntry {
    <#
    .SYNOPSIS
    Schreibt einen Log-Eintrag sowohl in die Datei als auch in die Konsole
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("Info", "Warning", "Error", "Success")]
        [string]$Level = "Info"
    )
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Level] $Message"
    
    # In Datei schreiben
    Add-Content -Path $LogFile -Value $LogMessage -ErrorAction SilentlyContinue
    
    # In Konsole schreiben
    $ForegroundColor = switch ($Level) {
        "Info"    { "White" }
        "Warning" { "Yellow" }
        "Error"   { "Red" }
        "Success" { "Green" }
        default   { "White" }
    }
    
    Write-Host $LogMessage -ForegroundColor $ForegroundColor
}

function Import-Configuration {
    <#
    .SYNOPSIS
    Lädt die Konfiguration aus der Config-Datei
    #>
    Write-LogEntry "Lade Konfiguration..." -Level Info
    
    if (Test-Path -Path $ConfigFile) {
        try {
            [xml]$Config = Get-Content -Path $ConfigFile
            Write-LogEntry "Konfigurationsdatei gefunden und geladen" -Level Success
            return $Config
        }
        catch {
            Write-LogEntry "FEHLER beim Laden der Konfiguration: $_" -Level Error
            return $null
        }
    }
    else {
        Write-LogEntry "Konfigurationsdatei nicht gefunden: $ConfigFile" -Level Warning
        return $null
    }
}

function Save-Configuration {
    <#
    .SYNOPSIS
    Speichert die Konfiguration in eine XML-Datei
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$SMTPServer,
        [int]$SMTPPort = 25,
        [string]$From,
        [string]$To,
        [bool]$UseSSL = $false
    )
    
    Write-LogEntry "Speichere Konfiguration..." -Level Info
    
    try {
                $xmlDoc = New-Object System.Xml.XmlDocument
                $declaration = $xmlDoc.CreateXmlDeclaration("1.0", "utf-8", $null)
                $null = $xmlDoc.AppendChild($declaration)

                $root = $xmlDoc.CreateElement("Configuration")
                $null = $xmlDoc.AppendChild($root)

                $smtpNode = $xmlDoc.CreateElement("SMTP")
                $null = $root.AppendChild($smtpNode)

                foreach ($entry in @{
                        Server = $SMTPServer
                        Port = $SMTPPort
                        From = $From
                        To = $To
                        UseSSL = $UseSSL
                }.GetEnumerator()) {
                        $node = $xmlDoc.CreateElement($entry.Key)
                        $node.InnerText = [string]$entry.Value
                        $null = $smtpNode.AppendChild($node)
                }

                $backupNode = $xmlDoc.CreateElement("Backup")
                $null = $root.AppendChild($backupNode)

                foreach ($entry in @{
                        EventLogName = $EventLogName
                        EventSourceName = $EventSourceName
                        SuccessEventId = $SuccessEventId
                        FailureEventId = $FailureEventId
                        WarningEventId = $WarningEventId
                }.GetEnumerator()) {
                        $node = $xmlDoc.CreateElement($entry.Key)
                        $node.InnerText = [string]$entry.Value
                        $null = $backupNode.AppendChild($node)
                }

                $xmlDoc.Save($ConfigFile)
        Write-LogEntry "Konfiguration gespeichert: $ConfigFile" -Level Success
        return $true
    }
    catch {
        Write-LogEntry "FEHLER beim Speichern der Konfiguration: $_" -Level Error
        return $false
    }
}

function Get-ServerLanguageContext {
    <#
    .SYNOPSIS
    Ermittelt die Spracheinstellungen des Servers
    #>
    $Culture = Get-Culture
    $UICulture = Get-UICulture
    $IsGerman = ($Culture.TwoLetterISOLanguageName -eq "de") -or ($UICulture.TwoLetterISOLanguageName -eq "de")

    return [pscustomobject]@{
        Culture = $Culture.Name
        UICulture = $UICulture.Name
        IsGerman = $IsGerman
    }
}

function ConvertTo-HtmlSafe {
    param(
        [AllowNull()]
        [string]$Value
    )

    return [System.Net.WebUtility]::HtmlEncode($Value)
}

function Get-RelevantBackupEvents {
    <#
    .SYNOPSIS
    Sucht relevante Backup-Ereignisse in Backup-, Application- und System-Logs
    #>
    param(
        [Parameter(Mandatory=$true)]
        [datetime]$Since,
        [Parameter(Mandatory=$true)]
        [datetime]$Until,
        [Parameter(Mandatory=$true)]
        $LanguageContext
    )

    $logsToScan = @($EventLogName, "Application", "System")
    $providerRegex = "(?i)wbengine|backup|volsnap|vss|shadow copy|sicherung|snap"
    $keywordRegex = if ($LanguageContext.IsGerman) {
        "(?i)backup|sicherung|sicherungsvorgang|volumenschattenkopie|vss|wbadmin|wbengine|volsnap"
    }
    else {
        "(?i)backup|volume shadow copy|shadow copy|vss|wbadmin|wbengine|volsnap"
    }

    $result = New-Object System.Collections.Generic.List[object]

    foreach ($logName in $logsToScan) {
        Write-LogEntry "Suche relevante Ereignisse in Log '$logName' zwischen $Since und $Until" -Level Info

        try {
            $logEvents = Get-WinEvent -FilterHashtable @{
                LogName = $logName
                StartTime = $Since
                EndTime = $Until
            } -ErrorAction Stop
        }
        catch {
            Write-LogEntry "Log '$logName' kann nicht gelesen werden: $_" -Level Warning
            continue
        }

        foreach ($evt in $logEvents) {
            $provider = [string]$evt.ProviderName
            $message = [string]$evt.Message
            $include = $false
            $status = "Info"

            if ($logName -eq $EventLogName -and $evt.Id -in @($SuccessEventId, $FailureEventId, $WarningEventId)) {
                $include = $true
                $status = switch ($evt.Id) {
                    $SuccessEventId { "Success" }
                    $FailureEventId { "Failure" }
                    $WarningEventId { "Warning" }
                    default { "Info" }
                }
            }
            else {
                if ($provider -match $providerRegex -or $message -match $keywordRegex) {
                    $include = $true
                }

                if ($include) {
                    if ($evt.LevelDisplayName -match "(?i)error|critical|fehler|kritisch") {
                        $status = "Failure"
                    }
                    elseif ($evt.LevelDisplayName -match "(?i)warning|warnung") {
                        $status = "Warning"
                    }
                    else {
                        $status = "Info"
                    }
                }
            }

            if ($include) {
                $result.Add([pscustomobject]@{
                    TimeCreated = $evt.TimeCreated
                    LogName = $evt.LogName
                    Id = $evt.Id
                    ProviderName = $provider
                    Level = [string]$evt.LevelDisplayName
                    Status = $status
                    Message = $message
                    RecordId = $evt.RecordId
                })
            }
        }
    }

    return $result | Sort-Object TimeCreated
}

function Get-OverallBackupStatus {
    param(
        [Parameter(Mandatory=$true)]
        [array]$Events
    )

    if (-not $Events -or $Events.Count -eq 0) {
        return "Success"
    }

    if ($Events.Status -contains "Failure") {
        return "Failure"
    }
    if ($Events.Status -contains "Warning") {
        return "Warning"
    }

    return "Success"
}

function New-BackupSummaryEmailBody {
    <#
    .SYNOPSIS
    Erstellt eine HTML-Zusammenfassung der relevanten Ereignisse der letzten Stunde
    #>
    param(
        [Parameter(Mandatory=$true)]
        [array]$Events,
        [Parameter(Mandatory=$true)]
        [datetime]$WindowStart,
        [Parameter(Mandatory=$true)]
        [datetime]$WindowEnd,
        [Parameter(Mandatory=$true)]
        $LanguageContext
    )

    $computerName = $env:COMPUTERNAME
    $eventCount = $Events.Count
    $failureCount = @($Events | Where-Object { $_.Status -eq "Failure" }).Count
    $warningCount = @($Events | Where-Object { $_.Status -eq "Warning" }).Count
    $successCount = @($Events | Where-Object { $_.Status -eq "Success" -or $_.Status -eq "Info" }).Count

    $rows = foreach ($evt in $Events) {
        $cssClass = switch ($evt.Status) {
            "Failure" { "failure" }
            "Warning" { "warning" }
            default { "success" }
        }

        $msg = ($evt.Message -replace "\r?\n", " ").Trim()
        if ($msg.Length -gt 220) {
            $msg = $msg.Substring(0, 220) + "..."
        }

        "<tr class='$cssClass'><td>$($evt.TimeCreated)</td><td>$(ConvertTo-HtmlSafe -Value $evt.LogName)</td><td>$($evt.Id)</td><td>$(ConvertTo-HtmlSafe -Value $evt.ProviderName)</td><td>$(ConvertTo-HtmlSafe -Value $evt.Level)</td><td>$(ConvertTo-HtmlSafe -Value $msg)</td></tr>"
    }

    if (-not $rows) {
        $rows = "<tr><td colspan='6'>Keine relevanten Ereignisse in der letzten Stunde gefunden.</td></tr>"
    }

    $languageNotice = if ($LanguageContext.IsGerman) {
        "Systemsprache erkannt: Deutsch ($($LanguageContext.Culture) / $($LanguageContext.UICulture))"
    }
    else {
        "System language detected: Non-German ($($LanguageContext.Culture) / $($LanguageContext.UICulture))"
    }

    return @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body { font-family: Arial, sans-serif; color: #333; }
        .header { background-color: #0a5a9c; color: white; padding: 18px; border-radius: 6px 6px 0 0; }
        .content { border: 1px solid #ddd; border-top: none; padding: 18px; border-radius: 0 0 6px 6px; }
        .summary { display: flex; gap: 14px; margin: 12px 0 16px 0; }
        .badge { padding: 8px 10px; border-radius: 4px; font-size: 13px; font-weight: bold; }
        .ok { background-color: #d4edda; color: #155724; }
        .warn { background-color: #fff3cd; color: #856404; }
        .fail { background-color: #f8d7da; color: #721c24; }
        table { width: 100%; border-collapse: collapse; margin-top: 12px; }
        th, td { border-bottom: 1px solid #ddd; padding: 8px; text-align: left; vertical-align: top; }
        th { background-color: #f3f3f3; }
        tr.success td { background-color: #f7fff9; }
        tr.warning td { background-color: #fffbef; }
        tr.failure td { background-color: #fff5f5; }
        .meta { margin-top: 14px; font-size: 12px; color: #666; }
    </style>
</head>
<body>
    <div class="header">
        <h2>Windows Backup Report (Hourly Summary)</h2>
        <p>Server: $computerName</p>
    </div>
    <div class="content">
        <p>Auswertungszeitraum: $($WindowStart.ToString('yyyy-MM-dd HH:mm:ss')) bis $($WindowEnd.ToString('yyyy-MM-dd HH:mm:ss'))</p>
        <div class="summary">
            <span class="badge ok">OK/Info: $successCount</span>
            <span class="badge warn">Warnungen: $warningCount</span>
            <span class="badge fail">Fehler: $failureCount</span>
            <span class="badge">Gesamt: $eventCount</span>
        </div>
        <table>
            <tr>
                <th>Zeit</th>
                <th>Log</th>
                <th>ID</th>
                <th>Quelle</th>
                <th>Level</th>
                <th>Nachricht (gekürzt)</th>
            </tr>
            $($rows -join "`n")
        </table>
        <div class="meta">
            <p>$languageNotice</p>
            <p>Windows Server Backup Monitoring v$ScriptVersion</p>
        </div>
    </div>
</body>
</html>
"@
}

function Send-BackupSummaryEmailReport {
    <#
    .SYNOPSIS
    Versendet die stündliche Zusammenfassung relevanter Backup-Ereignisse
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$SMTPServer,
        [int]$SMTPPort = 25,
        [Parameter(Mandatory=$true)]
        [string]$From,
        [Parameter(Mandatory=$true)]
        [string]$To,
        [Parameter(Mandatory=$true)]
        [array]$Events,
        [Parameter(Mandatory=$true)]
        [datetime]$WindowStart,
        [Parameter(Mandatory=$true)]
        [datetime]$WindowEnd,
        [Parameter(Mandatory=$true)]
        $LanguageContext,
        [PSCredential]$Credential,
        [bool]$UseSSL = $false
    )

    $overallStatus = Get-OverallBackupStatus -Events $Events
    $computerName = $env:COMPUTERNAME
    $subjectPrefix = switch ($overallStatus) {
        "Failure" { "[FEHLER]" }
        "Warning" { "[WARNUNG]" }
        default { "[OK]" }
    }

    $subject = "$subjectPrefix Windows Backup Stundenauswertung - $computerName ($($Events.Count) Ereignis(se))"
    $body = New-BackupSummaryEmailBody -Events $Events -WindowStart $WindowStart -WindowEnd $WindowEnd -LanguageContext $LanguageContext

    $mailParams = @{
        SmtpServer  = $SMTPServer
        Port        = $SMTPPort
        From        = $From
        To          = $To
        Subject     = $subject
        Body        = $body
        BodyAsHtml  = $true
        ErrorAction = "Stop"
        UseSsl      = $UseSSL
    }

    if ($Credential) {
        $mailParams["Credential"] = $Credential
    }

    try {
        Send-MailMessage @mailParams
        Write-LogEntry "Zusammenfassungs-E-Mail erfolgreich versendet an: $To" -Level Success
        return $true
    }
    catch {
        Write-LogEntry "FEHLER beim E-Mail-Versand der Zusammenfassung: $_" -Level Error
        return $false
    }
}

function Invoke-BackupMonitoring {
    <#
    .SYNOPSIS
    Führt die Backup-Überwachung durch
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$SMTPServer,
        [int]$SMTPPort = 25,
        [Parameter(Mandatory=$true)]
        [string]$From,
        [Parameter(Mandatory=$true)]
        [string]$To,
        [PSCredential]$Credential,
        [bool]$UseSSL = $false
    )
    
    Write-LogEntry "Starte stündliche Backup-Überwachung..." -Level Info

    $windowEnd = Get-Date
    $windowStart = $windowEnd.AddHours(-1)
    $languageContext = Get-ServerLanguageContext

    Write-LogEntry "Server-Sprache: $($languageContext.Culture) / UI: $($languageContext.UICulture)" -Level Info
    if ($languageContext.IsGerman) {
        Write-LogEntry "Deutsches System erkannt - deutsche und englische Backup-Muster werden berücksichtigt" -Level Info
    }
    else {
        Write-LogEntry "Nicht-deutsches System erkannt - englische Backup-Muster werden berücksichtigt" -Level Info
    }

    $events = Get-RelevantBackupEvents -Since $windowStart -Until $windowEnd -LanguageContext $languageContext
    Write-LogEntry "Relevante Ereignisse im letzten Zeitraum gefunden: $($events.Count)" -Level Info

    if (-not $events -or $events.Count -eq 0) {
        Write-LogEntry "Keine relevanten Ereignisse in der letzten Stunde gefunden. Es wird keine E-Mail versendet." -Level Info
        return
    }

    if ($events.Count -gt 1) {
        Write-LogEntry "Hinweis: Es wurden mehrere relevante Ereignisse in der letzten Stunde gefunden (normalerweise wird nur ein Abschlussereignis erwartet)." -Level Warning
    }

    $null = Send-BackupSummaryEmailReport -SMTPServer $SMTPServer -SMTPPort $SMTPPort `
        -From $From -To $To -Events $events -WindowStart $windowStart -WindowEnd $windowEnd `
        -LanguageContext $languageContext -Credential $Credential -UseSSL $UseSSL
}

function New-BackupMonitorTask {
    <#
    .SYNOPSIS
    Erstellt einen Scheduled Task zur Überwachung von Backup-Events
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$TaskName,
        [Parameter(Mandatory=$true)]
        [string]$ScriptPath
    )
    
    Write-LogEntry "Beginne Erstellung des Scheduled Tasks: $TaskName" -Level Info
    
    # Überprüfe, ob der Task bereits existiert
    $ExistingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    
    if ($ExistingTask) {
        Write-LogEntry "Scheduled Task existiert bereits. Entferne diesen..." -Level Warning
        try {
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
            Write-LogEntry "Alter Scheduled Task gelöscht" -Level Success
        }
        catch {
            Write-LogEntry "FEHLER beim Löschen des alten Tasks: $_" -Level Error
        }
    }
    
    # Erstelle die Task Action
    $TaskScript = @"
# Parameter für das Backup-Monitoring-Skript
& '$ScriptPath'
"@
    
    $TaskScript | Out-File -FilePath "$PSScriptRoot\BackupMonitorTask.ps1" -Encoding UTF8 -Force
    
    try {
        # Erstelle die Task-Aktion
        $TaskAction = New-ScheduledTaskAction `
            -Execute "powershell.exe" `
            -Argument "-NoProfile -NoLogo -NonInteractive -ExecutionPolicy Bypass -File `"$PSScriptRoot\BackupMonitorTask.ps1`""
        
        # Erstelle die Task-Einstellungen
        $TaskSettings = New-ScheduledTaskSettingsSet `
            -AllowStartIfOnBatteries `
            -RunOnlyIfNetworkAvailable `
            -StartWhenAvailable `
            -DontStopIfGoingOnBatteries `
            -MultipleInstances IgnoreNew
        
        Write-LogEntry "Erstelle Scheduled Task mit stündlichem Trigger..." -Level Info

        $TaskTrigger = New-ScheduledTaskTrigger `
            -Daily `
            -At "00:00" `
            -RepetitionInterval (New-TimeSpan -Hours 1) `
            -RepetitionDuration (New-TimeSpan -Days 1)

        # Registriere den Task mit stündlichem Trigger
        $null = Register-ScheduledTask `
            -TaskName $TaskName `
            -Action $TaskAction `
            -Trigger $TaskTrigger `
            -Settings $TaskSettings `
            -Description "Prüft stündlich Windows Backup Logs und versendet eine E-Mail-Zusammenfassung" `
            -RunLevel Highest `
            -ErrorAction Stop

        Write-LogEntry "Scheduled Task '$TaskName' erfolgreich erstellt" -Level Success
        Write-LogEntry "Trigger: stündlich (alle 60 Minuten)" -Level Success
        return $true
    }
    catch {
        Write-LogEntry "FEHLER beim Erstellen des Scheduled Tasks: $_" -Level Error
        Write-LogEntry "Stack Trace: $($_.ScriptStackTrace)" -Level Error
        return $false
    }
}


# ===========================
# HAUPTPROGRAMM
# ===========================

try {
    # Initialisiere Logging
    Initialize-LogPath
    Start-Transcript -Path $TranscriptFile -Append | Out-Null
    
    Write-LogEntry "========================" -Level Info
    Write-LogEntry "Windows Backup Reporter v$ScriptVersion" -Level Info
    Write-LogEntry "========================" -Level Info
    Write-LogEntry "PowerShell Version: $($PSVersionTable.PSVersion)" -Level Info
    Write-LogEntry "System: $env:COMPUTERNAME" -Level Info
    
    # Überprüfe Ausführungsrichtlinie
    Write-LogEntry "Überprüfe Ausführungsrichtlinie..." -Level Info
    $ExecutionPolicy = Get-ExecutionPolicy
    Write-LogEntry "Aktuelle ExecutionPolicy: $ExecutionPolicy" -Level Info
    
    # Verarbeite CreateTask Parameter
    if ($CreateTask) {
        Write-LogEntry "CreateTask-Parameter erkannt. Erstelle Scheduled Task..." -Level Info
        
        if (-not $SMTPServer -or -not $From -or -not $To) {
            Write-LogEntry "FEHLER: SMTPServer, From und To sind erforderlich für die Task-Erstellung" -Level Error
            Write-LogEntry "Beispielaufruf:" -Level Info
            Write-LogEntry ".\WindowsBackupReport.ps1 -CreateTask -SMTPServer 'smtp.contoso.com' -From 'backup@contoso.com' -To 'admin@contoso.com'" -Level Info
            exit 1
        }
        
        # Speichere die Konfiguration
        Save-Configuration -SMTPServer $SMTPServer -SMTPPort $SMTPPort `
            -From $From -To $To -UseSSL $UseSSL
        
        # Erstelle den Task
        $TaskCreated = New-BackupMonitorTask `
            -TaskName $TaskName `
            -ScriptPath $PSCommandPath
        
        if ($TaskCreated) {
            Write-LogEntry "Scheduled Task wurde erfolgreich erstellt und konfiguriert!" -Level Success
            Write-LogEntry "Der Task läuft stündlich und versendet eine Zusammenfassung der letzten Stunde." -Level Info
        }
        else {
            Write-LogEntry "FEHLER: Scheduled Task konnte nicht erstellt werden" -Level Error
        }
        
        exit 0
    }
    
    # Normale Ausführung - Überwachung starten
    Write-LogEntry "Starte normale Überwachung..." -Level Info
    
    # Lade oder verwende übergebene Konfiguration
    $Config = Import-Configuration
    
    if ($Config) {
        $SMTPServer = if ($SMTPServer) { $SMTPServer } else { $Config.Configuration.SMTP.Server }
        $SMTPPort = if ($SMTPPort -ne 25) { $SMTPPort } else { [int]$Config.Configuration.SMTP.Port }
        $From = if ($From) { $From } else { $Config.Configuration.SMTP.From }
        $To = if ($To) { $To } else { $Config.Configuration.SMTP.To }
        $UseSSL = if ($PSBoundParameters.ContainsKey('UseSSL')) { $UseSSL } else { [bool]$Config.Configuration.SMTP.UseSSL }
    }
    
    if (-not $SMTPServer -or -not $From -or -not $To) {
        Write-LogEntry "FEHLER: Erforderliche Parameter fehlen (SMTPServer, From, To)" -Level Error
        Write-LogEntry "Beispielaufruf:" -Level Info
        Write-LogEntry ".\WindowsBackupReport.ps1 -SMTPServer 'smtp.contoso.com' -From 'backup@contoso.com' -To 'admin@contoso.com'" -Level Info
        exit 1
    }
    
    Write-LogEntry "Konfiguration:" -Level Info
    Write-LogEntry "  SMTP-Server: $SMTPServer" -Level Info
    Write-LogEntry "  SMTP-Port: $SMTPPort" -Level Info
    Write-LogEntry "  Von: $From" -Level Info
    Write-LogEntry "  An: $To" -Level Info
    
    # Starte Backup-Überwachung
    Invoke-BackupMonitoring -SMTPServer $SMTPServer -SMTPPort $SMTPPort `
        -From $From -To $To -Credential $SMTPCredential -UseSSL $UseSSL
    
    Write-LogEntry "Überwachung abgeschlossen" -Level Success
    Write-LogEntry "========================" -Level Info
}
catch {
    Write-LogEntry "KRITISCHER FEHLER: $_" -Level Error
    Write-LogEntry "Stacktrace: $($_.ScriptStackTrace)" -Level Error
    exit 1
}
finally {
    Stop-Transcript | Out-Null
}
