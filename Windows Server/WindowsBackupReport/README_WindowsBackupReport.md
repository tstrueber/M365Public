# Windows Backup Report Monitor

PowerShell-Loesung zur stündlichen Ueberwachung von Windows-Backup-Ereignissen mit einer HTML-Zusammenfassungs-E-Mail.

## Ueberblick

Das Hauptscript laeuft stündlich und wertet die letzten 60 Minuten aus.

- Sucht relevante Backup-Ereignisse in:
  - Microsoft-Windows-Backup
  - Application
  - System
- Berücksichtigt deutsche und englische Ereignismuster.
- Sendet genau eine Zusammenfassungs-E-Mail pro Lauf.
- Bewertet den Gesamtstatus fuer den Betreff:
  - [FEHLER] wenn mindestens ein Fehler gefunden wurde
  - [WARNUNG] wenn mindestens eine Warnung gefunden wurde
  - [OK] sonst

## Anforderungen

- PowerShell 5.0 oder hoeher
- Windows Server 2016/2019 oder neuer
- Administrator-Rechte (fuer Task-Erstellung und Event-Log-Zugriff)
- SMTP-Server fuer Mailversand

## Parameter des Hauptscripts

| Parameter | Erforderlich | Beschreibung |
|-----------|-------------|-------------|
| -SMTPServer | Ja* | SMTP-Server fuer E-Mail-Versand |
| -SMTPPort | Nein | SMTP-Port (Standard: 25) |
| -From | Ja* | Absender-E-Mail-Adresse |
| -To | Ja* | Empfaenger-E-Mail-Adresse(n), mehrere mit ; |
| -SMTPCredential | Nein | SMTP-Anmeldedaten als PSCredential (nur bei direkter Ausfuehrung) |
| -UseSSL | Nein | SSL/TLS verwenden |
| -CreateTask | Nein | Erstellt/aktualisiert Scheduled Task |
| -TaskName | Nein | Name des Tasks (Standard: Windows Backup Report Monitor) |

* erforderlich bei direkter Ausfuehrung oder Task-Erstellung

## Schnellstart

### 1. Setup-Script verwenden

```powershell
.\Setup-WindowsBackupReport.ps1
```

Das Setup-Script:

- fragt SMTP-Werte interaktiv ab
- validiert Port und Empfaengerformat
- kann SMTP testen
- delegiert die Task-Erstellung an das Hauptscript

### 2. Task direkt mit Hauptscript erstellen

```powershell
.\WindowsBackupReport.ps1 -CreateTask `
  -SMTPServer "smtp.contoso.com" `
  -SMTPPort 25 `
  -From "backup@contoso.com" `
  -To "admin@contoso.com" `
  -UseSSL
```

## Scheduled Task Verhalten

Der Task wird als stündlicher Trigger erstellt.

- Trigger: alle 60 Minuten
- Aktion: Hauptscript aufrufen
- RunLevel: Highest
- MultipleInstances: IgnoreNew

Hinweis: Event-basierte Trigger (ID 4/8/12) werden nicht mehr verwendet.

## E-Mail-Inhalt

Die stündliche HTML-E-Mail enthaelt:

- Auswertungszeitraum (letzte Stunde)
- Anzahl OK/Info, Warnungen, Fehler, Gesamt
- Tabelle relevanter Ereignisse (Zeit, Log, ID, Quelle, Level, gekuerzte Nachricht)
- erkannte Server-Sprache (Culture/UI Culture)

## SMTP-Credentials

Wichtiger Hinweis:

- Das Setup-Script speichert aus Sicherheitsgruenden keine SMTP-Credentials in einem Scheduled Task.
- Fuer SMTP-Authentifizierung im geplanten Betrieb muss ein separates sicheres Credential-Konzept verwendet werden (z. B. verschluesselte Credential-Datei + Entschluesselung durch das Task-Konto).

## Logging

Das Hauptscript schreibt nach C:\Logs\WindowsBackup\:

- Transcript: Transcript_YYYYMMdd_HHmmss.log
- Tageslog: WindowsBackupReport_YYYYMMdd.log

Format Tageslog:

```text
[Timestamp] [Level] Message
```

## Konfigurationsdatei

Bei Task-Erstellung wird eine XML-Konfiguration am Script-Standort geschrieben:

- WindowsBackupReport.config.xml

Die XML wird sicher ueber XmlDocument erzeugt (korrektes Escaping von Sonderzeichen).

## Manuelle Tests

### 1. Einmaliger Lauf

```powershell
.\WindowsBackupReport.ps1 `
  -SMTPServer "smtp.contoso.com" `
  -SMTPPort 25 `
  -From "backup@contoso.com" `
  -To "admin@contoso.com"
```

### 2. Logs pruefen

```powershell
Get-ChildItem "C:\Logs\WindowsBackup\" |
  Sort-Object LastWriteTime -Descending |
  Select-Object -First 5

Get-Content "C:\Logs\WindowsBackup\WindowsBackupReport_$(Get-Date -Format 'yyyyMMdd').log" -Tail 50
```

### 3. Relevante Event Logs pruefen

```powershell
Get-WinEvent -ListLog "Microsoft-Windows-Backup"
Get-WinEvent -ListLog "Application"
Get-WinEvent -ListLog "System"
```

## Troubleshooting

### Task laeuft nicht

Pruefen:

1. Task vorhanden und aktiviert (`Get-ScheduledTask -TaskName "Windows Backup Report Monitor"`)
2. Letzter Lauf/Resultat (`Get-ScheduledTaskInfo -TaskName "Windows Backup Report Monitor"`)
3. Konto und Berechtigungen (Run with highest privileges)

### Keine oder zu viele Ereignisse in der Mail

Pruefen:

1. Event-Volumen in letzter Stunde in den Logs Microsoft-Windows-Backup, Application, System
2. Ob Drittanbieter-Backupsoftware viele zusaetzliche Events erzeugt
3. Tageslog fuer Hinweise zur Ereignis-Selektion

### SMTP Versand scheitert

Pruefen:

1. Erreichbarkeit (`Test-NetConnection smtp.contoso.com -Port 25`)
2. TLS/Port-Kombination (z. B. 25/587)
3. Wenn Auth erforderlich ist: Credential-Strategie fuer Scheduled Task umsetzen

## Version

- Script-Version: 1.0
- Architektur: stündliche Zusammenfassung (letzte 60 Minuten)
- Kompatibilitaet: PowerShell 5.0+
