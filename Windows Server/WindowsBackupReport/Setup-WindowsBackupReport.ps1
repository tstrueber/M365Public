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
    $PortInput = Read-Host "SMTP-Port [25]"
    $Config.Port = if ([string]::IsNullOrWhiteSpace($PortInput)) { 25 } else { [int]$PortInput }
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
    $Config.UseSSL = $SSLInput -eq "j" -or $SSLInput -eq "yes"
    Write-Info "SSL/TLS: $($Config.UseSSL)"
    
    # Anmeldedaten
    $AuthInput = Read-Host "SMTP-Authentifizierung erforderlich? (j/n) [n]"
    if ($AuthInput -eq "j" -or $AuthInput -eq "yes") {
        Write-Info "Geben Sie Ihre SMTP-Anmeldedaten ein:"
        $Config.Credential = Get-Credential -Message "SMTP-Anmeldedaten" -UserName $Config.From
    }
    
    return $Config
}

function Test-SMTPConnection {
    param($Config)
    
    Write-Header "SMTP-Verbindungstest"
    
    Write-Info "Teste Verbindung zu SMTP-Server..."
    
    try {
        $SMTPClient = New-Object Net.Mail.SmtpClient($Config.Server, $Config.Port)
        
        if ($Config.UseSSL) {
            $SMTPClient.EnableSsl = $true
        }
        
        if ($Config.Credential) {
            $SMTPClient.Credentials = $Config.Credential
        }
        
        $SMTPClient.Send($Config.From, $Config.To.Split(";")[0].Trim(), "Test", "Test")
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

function Create-ScheduledTask {
    param($Config)
    
    Write-Header "Scheduled Task erstellen"
    
    $ScriptPath = "c:\Scripts\GitHub\M365\Windows Server\WindowsBackupReport.ps1"
    
    # Überprüfe, ob das Script existiert
    if (-not (Test-Path -Path $ScriptPath)) {
        Write-Error "WindowsBackupReport.ps1 nicht gefunden unter: $ScriptPath"
        return $false
    }
    
    $TaskName = "Windows Backup Report Monitor"
    
    Write-Info "Erstelle Scheduled Task: $TaskName"
    
    # Prüfe, ob Task bereits existiert
    $ExistingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    
    if ($ExistingTask) {
        $Confirm = Read-Host "Task existiert bereits. Überschreiben? (j/n) [j]"
        if ($Confirm -eq "n") {
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
    
    # Konstruiere die Argumente
    $Arguments = @(
        "-NoProfile"
        "-NoLogo"
        "-NonInteractive"
        "-ExecutionPolicy Bypass"
        "-File `"$ScriptPath`""
        "-SMTPServer `"$($Config.Server)`""
        "-SMTPPort $($Config.Port)"
        "-From `"$($Config.From)`""
        "-To `"$($Config.To)`""
    )
    
    if ($Config.UseSSL) {
        $Arguments += "-UseSSL `$true"
    }
    
    if ($Config.Credential) {
        $Arguments += "-SMTPCredential (Get-Credential)"
    }
    
    $ArgumentString = $Arguments -join " "
    
    try {
        # Erstelle Task-Aktion
        $TaskAction = New-ScheduledTaskAction `
            -Execute "powershell.exe" `
            -Argument $ArgumentString
        
        # Erstelle Task-Einstellungen
        $TaskSettings = New-ScheduledTaskSettingsSet `
            -AllowStartIfOnBatteries `
            -RunOnlyIfNetworkAvailable `
            -StartWhenAvailable `
            -DontStopIfGoingOnBatteries `
            -MultipleInstances IgnoreNew
        
        # Erstelle Event Trigger
        Write-Info "Erstelle Event-Trigger..."
        
        # Trigger 1: Erfolgreicher Backup (Event ID 4)
        $Trigger1 = New-ScheduledTaskTrigger -CimTriggerType EventTrigger -TriggerAtLogon $false
        $Trigger1.StateChecks = @()
        $Trigger1.Subscription = "<QueryList><Query Id='0'><Select Path='Windows Server Backup'>*[System[(EventID=4)]]</Select></Query></QueryList>"
        
        # Trigger 2: Fehlgeschlagener Backup (Event ID 12)
        $Trigger2 = New-ScheduledTaskTrigger -CimTriggerType EventTrigger -TriggerAtLogon $false
        $Trigger2.StateChecks = @()
        $Trigger2.Subscription = "<QueryList><Query Id='0'><Select Path='Windows Server Backup'>*[System[(EventID=12)]]</Select></Query></QueryList>"
        
        # Trigger 3: Backup mit Fehlern (Event ID 8)
        $Trigger3 = New-ScheduledTaskTrigger -CimTriggerType EventTrigger -TriggerAtLogon $false
        $Trigger3.StateChecks = @()
        $Trigger3.Subscription = "<QueryList><Query Id='0'><Select Path='Windows Server Backup'>*[System[(EventID=8)]]</Select></Query></QueryList>"
        
        Write-Success "Event Trigger erstellt"
        
        # Registriere den Task
        Register-ScheduledTask `
            -TaskName $TaskName `
            -Action $TaskAction `
            -Trigger @($Trigger1, $Trigger2, $Trigger3) `
            -Settings $TaskSettings `
            -Description "Überwacht Windows Backup Events und versendet E-Mail-Benachrichtigungen" `
            -RunLevel Highest `
            -ErrorAction Stop | Out-Null
        
        Write-Success "Scheduled Task erfolgreich erstellt: $TaskName"
        Write-Success "Event-Trigger automatisch hinzugefügt:"
        Write-Host "  ✓ Event ID 4 (erfolgreicher Backup)" -ForegroundColor $SuccessColor
        Write-Host "  ✓ Event ID 12 (fehlgeschlagener Backup)" -ForegroundColor $SuccessColor
        Write-Host "  ✓ Event ID 8 (Backup mit Fehlern)" -ForegroundColor $SuccessColor
        
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
    Write-Host "  Anmeldedaten: $(if ($Config.Credential) { 'ja' } else { 'nein' })" -ForegroundColor White
    Write-Host ""
    
    Write-Host "Logs werden gespeichert unter:" -ForegroundColor $InfoColor
    Write-Host "  C:\Logs\WindowsBackup\" -ForegroundColor White
    Write-Host ""
    
    Write-Host "Script-Dateien:" -ForegroundColor $InfoColor
    Write-Host "  Hauptscript: c:\Scripts\GitHub\M365\Windows Server\WindowsBackupReport.ps1" -ForegroundColor White
    Write-Host "  Dokumentation: c:\Scripts\GitHub\M365\Windows Server\README_WindowsBackupReport.md" -ForegroundColor White
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
    if ($TestSMTP -eq "j") {
        Test-SMTPConnection -Config $SMTPConfig
    }
    
    # Task erstellen
    if (-not $SkipTaskCreation) {
        $CreateTask = Read-Host "Scheduled Task erstellen? (j/n) [j]"
        if ($CreateTask -ne "n") {
            $TaskCreated = Create-ScheduledTask -Config $SMTPConfig
            
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
