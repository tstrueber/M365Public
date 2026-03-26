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

function Load-Configuration {
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
        [string]$SMTPPort = 25,
        [string]$From,
        [string]$To,
        [bool]$UseSSL = $false
    )
    
    Write-LogEntry "Speichere Konfiguration..." -Level Info
    
    $config = @"
<?xml version="1.0" encoding="utf-8"?>
<Configuration>
  <SMTP>
    <Server>$SMTPServer</Server>
    <Port>$SMTPPort</Port>
    <From>$From</From>
    <To>$To</To>
    <UseSSL>$UseSSL</UseSSL>
  </SMTP>
  <Backup>
    <EventLogName>$EventLogName</EventLogName>
    <EventSourceName>$EventSourceName</EventSourceName>
    <SuccessEventId>$SuccessEventId</SuccessEventId>
    <FailureEventId>$FailureEventId</FailureEventId>
    <WarningEventId>$WarningEventId</WarningEventId>
  </Backup>
</Configuration>
"@
    
    try {
        Set-Content -Path $ConfigFile -Value $config -Encoding UTF8
        Write-LogEntry "Konfiguration gespeichert: $ConfigFile" -Level Success
        return $true
    }
    catch {
        Write-LogEntry "FEHLER beim Speichern der Konfiguration: $_" -Level Error
        return $false
    }
}

function Get-BackupEventDetails {
    <#
    .SYNOPSIS
    Ruft die Details eines Backup-Events vom Event Log ab
    #>
    param(
        [Parameter(Mandatory=$true)]
        [int]$EventId,
        [int]$MaxEvents = 1
    )
    
    Write-LogEntry "Suche Backup-Events mit ID $EventId..." -Level Info
    
    try {
        $Events = Get-WinEvent -FilterHashtable @{
            LogName = $EventLogName
            Id = $EventId
        } -MaxEvents $MaxEvents -ErrorAction SilentlyContinue
        
        if ($Events) {
            Write-LogEntry "Gefundene Events: $($Events.Count)" -Level Info
            return $Events
        }
        else {
            Write-LogEntry "Keine Events mit ID $EventId gefunden" -Level Warning
            return $null
        }
    }
    catch {
        Write-LogEntry "FEHLER beim Abrufen der Events: $_" -Level Error
        return $null
    }
}

function Create-BackupEmailBody {
    <#
    .SYNOPSIS
    Erstellt den E-Mail-Body für Backup-Berichte
    #>
    param(
        [Parameter(Mandatory=$true)]
        $Event,
        [Parameter(Mandatory=$true)]
        [ValidateSet("Success", "Failure", "Warning")]
        [string]$Status
    )
    
    $ComputerName = $env:COMPUTERNAME
    $EventTime = $Event.TimeCreated
    $EventMessage = $Event.Message
    
    $StatusColor = switch ($Status) {
        "Success" { "&#10003; Erfolgreich" }
        "Failure" { "&#10005; Fehlgeschlagen" }
        "Warning" { "⚠ Mit Fehlern abgeschlossen" }
    }
    
    $Body = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body { font-family: Arial, sans-serif; color: #333; }
        .header { background-color: #0078d4; color: white; padding: 20px; border-radius: 5px 5px 0 0; }
        .content { padding: 20px; border: 1px solid #ddd; border-radius: 0 0 5px 5px; }
        .status { padding: 10px; margin: 10px 0; border-radius: 3px; }
        .success { background-color: #d4edda; border-left: 4px solid #28a745; color: #155724; }
        .failure { background-color: #f8d7da; border-left: 4px solid #dc3545; color: #721c24; }
        .warning { background-color: #fff3cd; border-left: 4px solid #ffc107; color: #856404; }
        .details { background-color: #f5f5f5; padding: 10px; border-radius: 3px; font-family: monospace; white-space: pre-wrap; word-wrap: break-word; }
        table { width: 100%; border-collapse: collapse; margin: 10px 0; }
        th { background-color: #f0f0f0; padding: 8px; text-align: left; border-bottom: 2px solid #ddd; }
        td { padding: 8px; border-bottom: 1px solid #ddd; }
        .footer { margin-top: 20px; padding-top: 10px; border-top: 1px solid #ddd; font-size: 12px; color: #666; }
    </style>
</head>
<body>
    <div class="header">
        <h2>Windows Backup Report</h2>
        <p>Server: $ComputerName</p>
    </div>
    <div class="content">
        <div class="status $(if ($Status -eq 'Success') { 'success' } elseif ($Status -eq 'Failure') { 'failure' } else { 'warning' })">
            <strong>Status: $StatusColor</strong>
        </div>
        
        <table>
            <tr>
                <th>Eigenschaft</th>
                <th>Wert</th>
            </tr>
            <tr>
                <td>Zeitstempel</td>
                <td>$EventTime</td>
            </tr>
            <tr>
                <td>Event-ID</td>
                <td>$($Event.Id)</td>
            </tr>
            <tr>
                <td>Computer</td>
                <td>$ComputerName</td>
            </tr>
            <tr>
                <td>Quelle</td>
                <td>$($Event.ProviderName)</td>
            </tr>
        </table>
        
        <h3>Event Details:</h3>
        <div class="details">
$EventMessage
        </div>
        
        <div class="footer">
            <p>Dieses Skript wurde um $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ausgeführt.</p>
            <p>Windows Server Backup Monitoring v$ScriptVersion</p>
        </div>
    </div>
</body>
</html>
"@
    
    return $Body
}

function Send-BackupEmailReport {
    <#
    .SYNOPSIS
    Versendet einen Backup-Bericht per E-Mail
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$SMTPServer,
        [string]$SMTPPort = 25,
        [Parameter(Mandatory=$true)]
        [string]$From,
        [Parameter(Mandatory=$true)]
        [string]$To,
        [Parameter(Mandatory=$true)]
        $Event,
        [Parameter(Mandatory=$true)]
        [ValidateSet("Success", "Failure", "Warning")]
        [string]$Status,
        [PSCredential]$Credential,
        [bool]$UseSSL = $false
    )
    
    Write-LogEntry "Vorbereitung E-Mail-Versand für Status: $Status..." -Level Info
    
    try {
        $ComputerName = $env:COMPUTERNAME
        
        $Subject = switch ($Status) {
            "Success" { "✓ Windows Backup erfolgreich - $ComputerName" }
            "Failure" { "✗ Windows Backup fehlgeschlagen - $ComputerName" }
            "Warning" { "⚠ Windows Backup mit Fehlern - $ComputerName" }
        }
        
        $Body = Create-BackupEmailBody -Event $Event -Status $Status
        
        $MailParams = @{
            SmtpServer    = $SMTPServer
            Port          = $SMTPPort
            From          = $From
            To            = $To
            Subject       = $Subject
            Body          = $Body
            BodyAsHtml    = $true
            ErrorAction   = "Stop"
        }
        
        if ($Credential) {
            $MailParams["Credential"] = $Credential
        }
        
        if ($UseSSL) {
            $MailParams["UseSsl"] = $true
        }
        
        Send-MailMessage @MailParams
        Write-LogEntry "E-Mail erfolgreich versendet an: $To" -Level Success
        return $true
    }
    catch {
        Write-LogEntry "FEHLER beim E-Mail-Versand: $_" -Level Error
        return $false
    }
}

function Get-LatestBackupEvent {
    <#
    .SYNOPSIS
    Ruft das neueste Backup-Event ab
    #>
    param(
        [int]$EventId
    )
    
    Write-LogEntry "Suche das neueste Event mit ID $EventId..." -Level Info
    
    try {
        $Event = Get-WinEvent -FilterHashtable @{
            LogName = $EventLogName
            Id = $EventId
        } -MaxEvents 1 -ErrorAction SilentlyContinue
        
        if ($Event) {
            Write-LogEntry "Event gefunden vom $(($Event).TimeCreated)" -Level Success
            return $Event
        }
        else {
            Write-LogEntry "Kein Event mit ID $EventId gefunden" -Level Warning
            return $null
        }
    }
    catch {
        Write-LogEntry "FEHLER beim Abrufen des Events: $_" -Level Error
        return $null
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
        [string]$SMTPPort = 25,
        [Parameter(Mandatory=$true)]
        [string]$From,
        [Parameter(Mandatory=$true)]
        [string]$To,
        [PSCredential]$Credential,
        [bool]$UseSSL = $false
    )
    
    Write-LogEntry "Starte Backup-Überwachung..." -Level Info
    
    # Prüfe auf erfolgreiche Backups
    Write-LogEntry "Prüfe auf erfolgreiche Backups..." -Level Info
    $SuccessEvent = Get-LatestBackupEvent -EventId $SuccessEventId
    
    if ($SuccessEvent) {
        $EventTime = $SuccessEvent.TimeCreated
        $CheckTime = (Get-Date).AddMinutes(-5)
        
        if ($EventTime -gt $CheckTime) {
            Write-LogEntry "Aktuelles erfolgreiches Backup-Event gefunden" -Level Info
            Send-BackupEmailReport -SMTPServer $SMTPServer -SMTPPort $SMTPPort `
                -From $From -To $To -Event $SuccessEvent `
                -Status "Success" -Credential $Credential -UseSSL $UseSSL
        }
    }
    
    # Prüfe auf fehlgeschlagene Backups
    Write-LogEntry "Prüfe auf fehlgeschlagene Backups..." -Level Info
    $FailureEvent = Get-LatestBackupEvent -EventId $FailureEventId
    
    if ($FailureEvent) {
        $EventTime = $FailureEvent.TimeCreated
        $CheckTime = (Get-Date).AddMinutes(-5)
        
        if ($EventTime -gt $CheckTime) {
            Write-LogEntry "Aktuelles fehlgeschlagenes Backup-Event gefunden" -Level Info
            Send-BackupEmailReport -SMTPServer $SMTPServer -SMTPPort $SMTPPort `
                -From $From -To $To -Event $FailureEvent `
                -Status "Failure" -Credential $Credential -UseSSL $UseSSL
        }
    }
    
    # Prüfe auf Backups mit Fehlern
    Write-LogEntry "Prüfe auf Backups mit Fehlern..." -Level Info
    $WarningEvent = Get-LatestBackupEvent -EventId $WarningEventId
    
    if ($WarningEvent) {
        $EventTime = $WarningEvent.TimeCreated
        $CheckTime = (Get-Date).AddMinutes(-5)
        
        if ($EventTime -gt $CheckTime) {
            Write-LogEntry "Aktuelles Backup-Event mit Fehlern gefunden" -Level Info
            Send-BackupEmailReport -SMTPServer $SMTPServer -SMTPPort $SMTPPort `
                -From $From -To $To -Event $WarningEvent `
                -Status "Warning" -Credential $Credential -UseSSL $UseSSL
        }
    }
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
        [string]$ScriptPath,
        [string]$SMTPServer,
        [Parameter(Mandatory=$true)]
        [string]$From,
        [Parameter(Mandatory=$true)]
        [string]$To
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
& '$ScriptPath' -SMTPServer '$SMTPServer' -From '$From' -To '$To'
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
        
        Write-LogEntry "Erstelle Scheduled Task mit Event-Triggern..." -Level Info
        
        # Erstelle zuerst einen dummy-Trigger (wird später durch XML ersetzt)
        $DummyTrigger = New-ScheduledTaskTrigger -AtLogon
        
        # Registriere den Task zuerst mit Dummy-Trigger
        $Task = Register-ScheduledTask `
            -TaskName $TaskName `
            -Action $TaskAction `
            -Trigger $DummyTrigger `
            -Settings $TaskSettings `
            -Description "Überwacht Windows Backup Events und versendet E-Mail-Benachrichtigungen" `
            -RunLevel Highest `
            -ErrorAction Stop
        
        Write-LogEntry "Basis-Task erstellt, füge Event-Trigger hinzu..." -Level Info
        
        # Nutze TaskScheduler COM-API für Event-basierte Trigger (PowerShell 7 kompatibel)
        try {
            $TaskScheduler = New-Object -ComObject "Schedule.Service"
            $TaskScheduler.Connect()
            
            $RootFolder = $TaskScheduler.GetFolder("\")
            $ScheduledTask = $RootFolder.GetTask($TaskName)
            $Definition = $ScheduledTask.Definition
            
            # Entferne den Dummy-Trigger
            while ($Definition.Triggers.Count -gt 0) {
                $Definition.Triggers.Remove(1)
            }
            
            # Erstelle Event Trigger 1: Success (Event ID 4)
            $EventTrigger1 = $Definition.Triggers.Create(9)  # 9 = EventTrigger
            $EventTrigger1.Subscription = "<QueryList><Query Id='0'><Select Path='Microsoft-Windows-Backup'>*[System[(EventID=4)]]</Select></Query></QueryList>"
            $EventTrigger1.Enabled = $true
            $EventTrigger1.Id = "EventTrigger_Success"
            
            # Erstelle Event Trigger 2: Failure (Event ID 12)
            $EventTrigger2 = $Definition.Triggers.Create(9)  # 9 = EventTrigger
            $EventTrigger2.Subscription = "<QueryList><Query Id='0'><Select Path='Microsoft-Windows-Backup'>*[System[(EventID=12)]]</Select></Query></QueryList>"
            $EventTrigger2.Enabled = $true
            $EventTrigger2.Id = "EventTrigger_Failure"
            
            # Erstelle Event Trigger 3: Warning (Event ID 8)
            $EventTrigger3 = $Definition.Triggers.Create(9)  # 9 = EventTrigger
            $EventTrigger3.Subscription = "<QueryList><Query Id='0'><Select Path='Microsoft-Windows-Backup'>*[System[(EventID=8)]]</Select></Query></QueryList>"
            $EventTrigger3.Enabled = $true
            $EventTrigger3.Id = "EventTrigger_Warning"
            
            # Speichere die aktualisierte Task-Definition
            $RootFolder.RegisterTaskDefinition($TaskName, $Definition, 6, $null, $null, 3) | Out-Null
            
            Write-LogEntry "Scheduled Task '$TaskName' erfolgreich erstellt" -Level Success
            Write-LogEntry "Event-Trigger automatisch hinzugefügt:" -Level Success
            Write-LogEntry "  ✓ Event ID 4 (erfolgreicher Backup)" -Level Success
            Write-LogEntry "  ✓ Event ID 12 (fehlgeschlagener Backup)" -Level Success
            Write-LogEntry "  ✓ Event ID 8 (Backup mit Fehlern)" -Level Success
            
            return $true
        }
        catch {
            Write-LogEntry "FEHLER beim Konfigurieren der Event-Trigger via COM: $_" -Level Error
            Write-LogEntry "Der Task wurde erstellt, aber die Event-Trigger müssen manuell konfiguriert werden." -Level Warning
            return $true
        }
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
            -ScriptPath $PSCommandPath `
            -SMTPServer $SMTPServer `
            -From $From `
            -To $To
        
        if ($TaskCreated) {
            Write-LogEntry "Scheduled Task wurde erfolgreich erstellt und konfiguriert!" -Level Success
            Write-LogEntry "Der Task wird sofort bei den folgenden Backups ausgelöst:" -Level Info
            Write-LogEntry "  • Wenn ein Backup erfolgreich abgeschlossen wird" -Level Info
            Write-LogEntry "  • Wenn ein Backup fehlschlägt" -Level Info
            Write-LogEntry "  • Wenn ein Backup mit Fehlern abgeschlossen wird" -Level Info
        }
        else {
            Write-LogEntry "FEHLER: Scheduled Task konnte nicht erstellt werden" -Level Error
        }
        
        exit 0
    }
    
    # Normale Ausführung - Überwachung starten
    Write-LogEntry "Starte normale Überwachung..." -Level Info
    
    # Lade oder verwende übergebene Konfiguration
    $Config = Load-Configuration
    
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
