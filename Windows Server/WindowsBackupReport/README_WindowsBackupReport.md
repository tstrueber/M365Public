# Windows Backup Report Monitor

Ein PowerShell-Script zur automatisierten Überwachung von Windows Server Backups mit E-Mail-Benachrichtigungen.

## 📋 Anforderungen

- **PowerShell 5.0** oder höher
- **Windows Server 2016/2019** oder höher
- **Administrator-Rechte** erforderlich
- Zugang zu E-Mail/SMTP-Server

## 🚀 Schnellstart

### 1. Script-Parameter

| Parameter | Erforderlich | Beschreibung |
|-----------|-------------|-------------|
| `-SMTPServer` | Ja* | SMTP-Server für E-Mail-Versand |
| `-SMTPPort` | Nein | SMTP-Port (Standard: 25) |
| `-From` | Ja* | Absender-E-Mail-Adresse |
| `-To` | Ja* | Empfänger-E-Mail-Adresse(n) |
| `-SMTPCredential` | Nein | Anmeldedaten für SMTP (PSCredential) |
| `-UseSSL` | Nein | SSL/TLS verwenden (Standard: $false) |
| `-CreateTask` | Nein | Erstellt Scheduled Task |
| `-TaskName` | Nein | Name des Tasks (Standard: "Windows Backup Report Monitor") |

*erforderlich bei direkter Ausführung oder Task-Erstellung

### 2. Task-Setup (Methode 1: Automatisiert)

```powershell
.\WindowsBackupReport.ps1 -CreateTask `
  -SMTPServer "smtp.contoso.com" `
  -From "backup@contoso.com" `
  -To "admin@contoso.com"
```

Das Script zeigt dann Anweisungen für die Konfiguration im Task Scheduler.

### 3. Task-Setup (Methode 2: Manuell)

1. **Task Scheduler öffnen:** `taskschd.msc`
2. **Create Task** klicken und folgende Details eingeben:

**General Tab:**
- Name: "Windows Backup Report Monitor"
- Description: "Überwacht Windows Backup Events und versendet E-Mail-Benachrichtigungen"
- "Run with highest privileges" ✓

**Triggers Tab:**
Erstellen Sie **3 Trigger** (nacheinander):

**Trigger 1 - Erfolgreicher Backup:**
- "New..."
- "Begin the task:" → "On an event"
- Log: "Windows Server Backup"
- Source: "Backup"
- Event ID: **4**
- OK

**Trigger 2 - Fehlgeschlagener Backup:**
- "New..."
- "Begin the task:" → "On an event"
- Log: "Windows Server Backup"
- Source: "Backup"
- Event ID: **12**
- OK

**Trigger 3 - Backup mit Fehlern:**
- "New..."
- "Begin the task:" → "On an event"
- Log: "Windows Server Backup"
- Source: "Backup"
- Event ID: **8**
- OK

**Actions Tab:**
- "New..."
- Program/script: `powershell.exe`
- Add arguments:
```
-NoProfile -NoLogo -NonInteractive -ExecutionPolicy Bypass -File "c:\Scripts\GitHub\M365\Windows Server\WindowsBackupReport.ps1" -SMTPServer "your.smtp.server" -From "backup@domain.com" -To "admin@domain.com"
```
- Start in: `c:\Scripts\GitHub\M365\Windows Server`
- OK

**Conditions Tab:**
- Optional: "Run only if network is available" ✓

**Settings Tab:**
- "If the task fails, restart every:" optional
- "Stop the task if it runs longer than:" optional

## 📧 E-Mail-Benachrichtigungen

Das Script sendet HTML-formatierte E-Mails mit:

- **Erfolgreicher Backup:** ✓ Grüner Status mit Details
- **Fehlgeschlagener Backup:** ✗ Roter Status mit Fehlerdetails
- **Backup mit Fehlern:** ⚠ Gelber Status mit Warnung

Jede E-Mail enthält:
- Backup-Status
- Zeitstempel
- Event-ID
- Computer-Name
- Quelle
- Vollständige Event-Nachricht

## 📝 Logging

**Transcript-Logfiles:**
- Pfad: `C:\Logs\WindowsBackup\`
- Dateiname: `Transcript_YYYYMMdd_HHmmss.log`
- Format: Start-Transcript Output mit allen Details

**Tageslogfiles:**
- Pfad: `C:\Logs\WindowsBackup\`
- Dateiname: `WindowsBackupReport_YYYYMMdd.log`
- Format: `[Timestamp] [Level] Message`

## ⚙️ Konfigurationsdatei

Das Script erstellt automatisch eine Konfigurationsdatei:

**Pfad:** `c:\Scripts\GitHub\M365\Windows Server\WindowsBackupReport.config.xml`

**Format:**
```xml
<?xml version="1.0" encoding="utf-8"?>
<Configuration>
  <SMTP>
    <Server>smtp.contoso.com</Server>
    <Port>25</Port>
    <From>backup@contoso.com</From>
    <To>admin@contoso.com</To>
    <UseSSL>False</UseSSL>
  </SMTP>
  <Backup>
    <EventLogName>Windows Server Backup</EventLogName>
    <EventSourceName>Backup</EventSourceName>
    <SuccessEventId>4</SuccessEventId>
    <FailureEventId>12</FailureEventId>
    <WarningEventId>8</WarningEventId>
  </Backup>
</Configuration>
```

## 🔍 Fehlerbehandlung

Das Script implementiert:
- ✓ Try-Catch-Blöcke für alle kritischen Operationen
- ✓ Detaillierte Fehlerlogging in Dateien und Konsole
- ✓ Validierung aller erforderlichen Parameter
- ✓ Überprüfung von Verzeichnissen und Dateien
- ✓ Event-Log-Verfügbarkeitsprüfung

## 🔐 Sicherheit

**Empfohlene Maßnahmen:**

1. **Ausführungsrichtlinie:** Script muss mit `-ExecutionPolicy Bypass` ausgeführt werden
2. **Task-Berechtigungen:** Mit "Run with highest privileges" ausführen
3. **SMTP-Anmeldedaten:** Verwenden Sie `Get-Credential` für sichere Eingabe
4. **LogPath:** Ensure only authorized users can access logs

## 🧪 Testen des Scripts

### Test 1: Manuelle Ausführung
```powershell
.\WindowsBackupReport.ps1 `
  -SMTPServer "smtp.contoso.com" `
  -From "backup@contoso.com" `
  -To "admin@contoso.com"
```

### Test 2: Logs prüfen
```powershell
Get-ChildItem "C:\Logs\WindowsBackup\" | Sort-Object LastWriteTime -Descending | Select-Object -First 5
Get-Content "C:\Logs\WindowsBackup\WindowsBackupReport_$(Get-Date -Format 'yyyyMMdd').log" -Tail 20
```

### Test 3: Event Log prüfen
```powershell
Get-WinEvent -FilterHashtable @{LogName="Windows Server Backup"; Id=4} -MaxEvents 5
Get-WinEvent -FilterHashtable @{LogName="Windows Server Backup"; Id=12} -MaxEvents 5
Get-WinEvent -FilterHashtable @{LogName="Windows Server Backup"; Id=8} -MaxEvents 5
```

## 📊 Event IDs - Referenz

| Event ID | Beschreibung | Action |
|----------|-------------|--------|
| 4 | Backup erfolgreich abgeschlossen | E-Mail versenden |
| 8 | Backup mit Fehlern abgeschlossen | E-Mail versenden |
| 12 | Backup fehlgeschlagen | E-Mail versenden |

## 🐛 Troubleshooting

### Problem: "Access Denied" beim Event Log
**Lösung:** Script mit Administrator-Rechten ausführen
```powershell
Start-Process powershell -ArgumentList "-NoExit -Command `"& 'path\WindowsBackupReport.ps1'`"" -Verb RunAs
```

### Problem: E-Mails werden nicht versendet
**Überprüfen:**
1. SMTP-Server erreichbar: `Test-NetConnection smtp.contoso.com -Port 25`
2. Anmeldedaten korrekt: `Send-MailMessage` mit manuellen Parametern testen
3. Logs prüfen: `Get-Content C:\Logs\WindowsBackup\*` -Tail 50

### Problem: Task wird nicht ausgelöst
**Überprüfen:**
1. Event IDs im Task Scheduler korrekt konfiguriert
2. Task-Berechtigungen: "Run with highest privileges"
3. Event Log verfügbar: `Get-WinEvent -ListLog "Windows Server Backup"`
4. Task-Logs: Event Viewer → "Windows Logs" → "System"

### Problem: Transcript-Datei nicht gefunden
**Überprüfen:**
1. Verzeichnis existiert: `Test-Path C:\Logs\WindowsBackup\`
2. Schreibberechtigung: `icacls C:\Logs\WindowsBackup\`

## 📞 Support

**Log-Sammlung für Support:**
```powershell
$LogPath = "C:\Logs\WindowsBackup"
$ZipPath = "$env:USERPROFILE\Desktop\WindowsBackupLogs.zip"
Compress-Archive -Path $LogPath -DestinationPath $ZipPath -Force
Write-Host "Logs gepackt: $ZipPath"
```

## 📄 Version

- **Script-Version:** 1.0
- **Kompatibilität:** PowerShell 5.0+
- **Betriebssysteme:** Windows Server 2016, 2019 und neuer
- **Erstellt:** März 2026

---

**Hinweis:** Dieses Script erfordert regelmäßige Überprüfung der Log-Dateien und Event-Trigger-Konfiguration auf dem jeweiligen Server.
