<#
.SYNOPSIS
Setup-Script für Windows Backup Report Monitor
Vereinfachte Konfiguration und Task-Erstellung

.DESCRIPTION
Dieses Script unterstützt Sie beim Setup des Windows Backup Report Monitor Scripts.
Es fragt Sie nach den erforderlichen Parametern und erstellt den Scheduled Task.

.EXAMPLE
.\Setup-WindowsBackupReport.ps1

#>

param(
    [switch]$SkipTaskCreation
)

# ===========================
# Farben definieren
# ===========================
$InfoColor = "Cyan"
$SuccessColor = "Green"
$WarningColor = "Yellow"
$ErrorColor = "Red"

# ===========================
# Funktionen
# ===========================

function Write-Header {
    param([string]$Text)
    Write-Host "`n" -NoNewline
    Write-Host "=" * 60 -ForegroundColor $InfoColor
    Write-Host $Text -ForegroundColor $InfoColor -NoNewline
    Write-Host "`n" + ("=" * 60) -ForegroundColor $InfoColor
}

function Write-Info {
    param([string]$Text)
    Write-Host "[i] $Text" -ForegroundColor $InfoColor
}

function Write-Success {
    param([string]$Text)
    Write-Host "[✓] $Text" -ForegroundColor $SuccessColor
}

function Write-Warning {
    param([string]$Text)
    Write-Host "[!] $Text" -ForegroundColor $WarningColor
}

function Write-Error {
    param([string]$Text)
    Write-Host "[✗] $Text" -ForegroundColor $ErrorColor
}

function Test-YesInput {
    param([string]$InputText)
    return $InputText -match '^(?i:j|ja|y|yes|true|1)$'
}

function Test-NoInput {
    param([string]$InputText)
    return $InputText -match '^(?i:n|nein|no|false|0)$'
}

function Test-AdminPrivileges {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Error "Dieses Script erfordert Administrator-Rechte!"
        Write-Info "Bitte starten Sie PowerShell als Administrator neu."
        exit 1
    }
}

function Test-PowerShellVersion {
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        Write-Error "PowerShell 5.0 oder höher erforderlich!"
        Write-Info "Aktuelle Version: $($PSVersionTable.PSVersion)"
        exit 1
    }
}

function Get-SMTPConfiguration {
    Write-Header "SMTP-Konfiguration"
    
    Write-Info "Geben Sie Ihre SMTP-Einstellungen ein:"
    
    $Config = @{}
    
    # SMTP-Server
    do {
        $Server = Read-Host "SMTP-Server (z.B. smtp.contoso.com)"
        if ([string]::IsNullOrWhiteSpace($Server)) {
            Write-Warning "SMTP-Server ist erforderlich!"
        }
    } while ([string]::IsNullOrWhiteSpace($Server))
    $Config.Server = $Server
    
    # SMTP-Port
    do {
        $PortInput = Read-Host "SMTP-Port [25]"

        if ([string]::IsNullOrWhiteSpace($PortInput)) {
            $Config.Port = 25
            $ValidPort = $true
        }
        else {
            $ParsedPort = 0
            $ValidPort = [int]::TryParse($PortInput, [ref]$ParsedPort) -and $ParsedPort -ge 1 -and $ParsedPort -le 65535

            if ($ValidPort) {
                $Config.Port = $ParsedPort
            }
            else {
                Write-Warning "Ungültiger SMTP-Port. Bitte geben Sie eine Zahl zwischen 1 und 65535 ein."
            }
        }
    } while (-not $ValidPort)

    Write-Info "SMTP-Port: $($Config.Port)"
    
    # Von-Adresse
    do {
        $From = Read-Host "Von-E-Mail-Adresse (z.B. backup@contoso.com)"
        if ([string]::IsNullOrWhiteSpace($From)) {
            Write-Warning "Von-E-Mail-Adresse ist erforderlich!"
        }
    } while ([string]::IsNullOrWhiteSpace($From))
    $Config.From = $From
    
    # An-Adresse(n)
    do {
        $To = Read-Host "An-E-Mail-Adresse(n) (mehrere durch ; getrennt)"
        if ([string]::IsNullOrWhiteSpace($To)) {
            Write-Warning "An-E-Mail-Adresse ist erforderlich!"
        }
    } while ([string]::IsNullOrWhiteSpace($To))
    $Config.To = $To
    
    # SSL/TLS
    $SSLInput = Read-Host "SSL/TLS verwenden? (j/n) [n]"
    $Config.UseSSL = Test-YesInput -InputText $SSLInput
    Write-Info "SSL/TLS: $($Config.UseSSL)"

    Write-Info "Hinweis: SMTP-Anmeldedaten werden im Setup nicht erfasst, da sie nicht sicher im Scheduled Task gespeichert werden."
    
    return $Config
}

function Test-SMTPConnection {
    param($Config)
    
    Write-Header "SMTP-Verbindungstest"
    
    Write-Info "Teste Verbindung zu SMTP-Server..."
    
    try {
        $recipientList = @(
            $Config.To.Split(";") |
                ForEach-Object { $_.Trim() } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )

        if ($recipientList.Count -eq 0) {
            throw "Keine gültigen Empfängeradressen gefunden."
        }

        foreach ($recipient in $recipientList) {
            $null = New-Object System.Net.Mail.MailAddress($recipient)
        }

        $SMTPClient = New-Object Net.Mail.SmtpClient($Config.Server, $Config.Port)
        
        if ($Config.UseSSL) {
            $SMTPClient.EnableSsl = $true
        }
        
        $SMTPClient.Send($Config.From, $recipientList[0], "Test", "Test")
        $SMTPClient.Dispose()
        
        Write-Success "SMTP-Verbindung erfolgreich!"
        return $true
    }
    catch {
        Write-Warning "SMTP-Verbindungstest fehlgeschlagen: $_"
        Write-Info "Das Script wird trotzdem erstellt. Überprüfen Sie die Einstellungen später."
        return $false
    }
}

function Resolve-MainScriptPath {
    $candidates = @(
        (Join-Path -Path $PSScriptRoot -ChildPath "WindowsBackupReport.ps1"),
        "c:\Scripts\GitHub\M365\Windows Server\WindowsBackupReport.ps1"
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -Path $candidate) {
            return $candidate
        }
    }

    return $null
}

function New-BackupMonitorTask {
    param($Config)
    
    Write-Header "Scheduled Task erstellen"
    
    $ScriptPath = Resolve-MainScriptPath
    
    # Überprüfe, ob das Script existiert
    if ([string]::IsNullOrWhiteSpace($ScriptPath) -or -not (Test-Path -Path $ScriptPath)) {
        Write-Error "WindowsBackupReport.ps1 konnte nicht gefunden werden."
        Write-Info "Erwartete Pfade: $PSScriptRoot\WindowsBackupReport.ps1 oder c:\Scripts\GitHub\M365\Windows Server\WindowsBackupReport.ps1"
        return $false
    }
    
    $TaskName = "Windows Backup Report Monitor"
    
    Write-Info "Erstelle Scheduled Task: $TaskName"
    
    # Prüfe, ob Task bereits existiert
    $ExistingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    
    if ($ExistingTask) {
        $Confirm = Read-Host "Task existiert bereits. Überschreiben? (j/n) [j]"
        if (Test-NoInput -InputText $Confirm) {
            Write-Warning "Task-Erstellung abgebrochen"
            return $false
        }
        
        try {
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
            Write-Info "Alter Task gelöscht"
        }
        catch {
            Write-Error "Konnte alten Task nicht löschen: $_"
            return $false
        }
    }
    
    Write-Info "Verwende Hauptscript: $ScriptPath"
    Write-Info "Delegiere Task-Erstellung an das Hauptscript (stündlicher Trigger)."

    try {
        $PowerShellExe = if (Get-Command pwsh -ErrorAction SilentlyContinue) { "pwsh" } else { "powershell.exe" }
        $Arguments = @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", $ScriptPath,
            "-CreateTask",
            "-SMTPServer", $Config.Server,
            "-SMTPPort", [string]$Config.Port,
            "-From", $Config.From,
            "-To", $Config.To,
            "-TaskName", $TaskName
        )

        if ($Config.UseSSL) {
            $Arguments += "-UseSSL"
        }

        $TaskProcess = Start-Process -FilePath $PowerShellExe -ArgumentList $Arguments -NoNewWindow -Wait -PassThru

        if ($TaskProcess.ExitCode -ne 0) {
            Write-Error "Task-Erstellung durch das Hauptscript fehlgeschlagen (ExitCode: $($TaskProcess.ExitCode))"
            return $false
        }

        Write-Success "Scheduled Task erfolgreich erstellt: $TaskName"
        Write-Success "Trigger: stündlich (alle 60 Minuten)"
        return $true
    }
    catch {
        Write-Error "Fehler beim Erstellen des Scheduled Tasks: $_"
        return $false
    }
}

function Show-Summary {
    param($Config)
    
    Write-Header "Zusammenfassung"
    
    Write-Host ""
    Write-Host "SMTP-Konfiguration:" -ForegroundColor $InfoColor
    Write-Host "  Server: $($Config.Server):$($Config.Port)" -ForegroundColor White
    Write-Host "  Von: $($Config.From)" -ForegroundColor White
    Write-Host "  An: $($Config.To)" -ForegroundColor White
    Write-Host "  SSL/TLS: $($Config.UseSSL)" -ForegroundColor White
    Write-Host "  Anmeldedaten: nicht im Setup enthalten" -ForegroundColor White
    Write-Host ""
    
    Write-Host "Logs werden gespeichert unter:" -ForegroundColor $InfoColor
    Write-Host "  C:\Logs\WindowsBackup\" -ForegroundColor White
    Write-Host ""
    
    $MainScriptPath = Resolve-MainScriptPath

    Write-Host "Script-Dateien:" -ForegroundColor $InfoColor
    if ($MainScriptPath) {
        Write-Host "  Hauptscript: $MainScriptPath" -ForegroundColor White
    }
    else {
        Write-Host "  Hauptscript: nicht gefunden" -ForegroundColor $WarningColor
    }
    Write-Host "  Setup-Script: $PSCommandPath" -ForegroundColor White
    Write-Host ""
}

# ===========================
# Hauptprogramm
# ===========================

try {
    # Überprüfungen
    Test-AdminPrivileges
    Test-PowerShellVersion
    
    Write-Header "Windows Backup Report Monitor - Setup"
    
    # Konfiguration abrufen
    $SMTPConfig = Get-SMTPConfiguration
    
    # SMTP testen
    Show-Summary -Config $SMTPConfig
    
    $TestSMTP = Read-Host "SMTP-Verbindung testen? (j/n) [n]"
    if (Test-YesInput -InputText $TestSMTP) {
        Test-SMTPConnection -Config $SMTPConfig
    }
    
    # Task erstellen
    if (-not $SkipTaskCreation) {
        $CreateTask = Read-Host "Scheduled Task erstellen? (j/n) [j]"
        if (-not (Test-NoInput -InputText $CreateTask)) {
            $TaskCreated = New-BackupMonitorTask -Config $SMTPConfig
            
            if ($TaskCreated) {
                Write-Success "Setup erfolgreich abgeschlossen!"
            }
            else {
                Write-Warning "Setup teilweise abgeschlossen. Bitte überprüfen Sie die Fehler."
            }
        }
    }
    
    Write-Header "Setup abgeschlossen"
    Write-Success "Vielen Dank für die Nutzung von Windows Backup Report Monitor!"
}
catch {
    Write-Error "Kritischer Fehler während des Setups: $_"
    exit 1
}
