# gepflegte Variablen
$phishingpolicyname = "Phish Policy"
$exportfilepath = "c:\temp"
# set to false when you want to overwrite the impersonation users every time you run the script
# if set to true the script adds the current list of VIP flagged users to the impersonation users list
# this way you continue to protect your users from impersonation of VIP users after they eventually left the company
$addimpersonationusers = $true
$prioritygroupfilter1 = "alias -like 'priority_group_1*'"
$prioritygroupfilter2 = "alias -like 'priority_group_2*'"
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$transcriptPath = Join-Path -Path $exportfilepath -ChildPath "$timestamp-PriorityAccounts_and_Impersonation_Users-transcript.txt"
$poslogJsonPath = Join-Path -Path $exportfilepath -ChildPath "$timestamp-poslog.json"
$neglogJsonPath = Join-Path -Path $exportfilepath -ChildPath "$timestamp-neglog.json"

# Log-Variablen
$poslog = [System.Collections.Generic.List[string]]::new()
$neglog = [System.Collections.Generic.List[string]]::new()

# function for setting VIP tag to users
function Set-VipTag {
    param($userlist)

    foreach ($user in $userlist) {
        try {
            Set-User $user.alias -VIP $true -Confirm:$false -ErrorAction Stop
            $alias = $user.alias
            $displayname = $user.displayname
            Write-Host "Set VIP tag to user $displayname with alias $alias"
            $poslog.Add("VIP tag gesetzt: $displayname ($alias)")
        }
        catch {
            $alias = $user.alias
            $displayname = $user.displayname
            $errorMessage = "Fehler beim Setzen des VIP tags fuer $displayname ($alias): $($_.Exception.Message)"
            Write-Host $errorMessage -ForegroundColor Red
            $neglog.Add($errorMessage)
        }
    }
}

# function for getting the current VIP user list and optionally exporting it
function Get-VipUserList {
    param([switch]$Export)

    try {
        $date = Get-Date -Format FileDateTime
        $vipusers = Get-User -IsVIP -ErrorAction Stop
        if ($Export) {
            $vipusers | ConvertTo-Json | Out-File "$exportfilepath\$date-vip-users.json"
            $poslog.Add("VIP-Liste exportiert nach $exportfilepath\\$date-vip-users.json")
        }
        return $vipusers
    }
    catch {
        $errorMessage = "Fehler beim Abrufen/Exportieren der VIP-Liste: $($_.Exception.Message)"
        $neglog.Add($errorMessage)
        throw
    }
}

$transcriptStarted = $false

try {
    Write-Host "[1/8] Initialisiere Lauf und Logdateien"
    if (-not (Test-Path -Path $exportfilepath)) {
        New-Item -Path $exportfilepath -ItemType Directory -Force | Out-Null
        $poslog.Add("Exportverzeichnis erstellt: $exportfilepath")
    }

    Start-Transcript -Path $transcriptPath -Force
    $transcriptStarted = $true
    $poslog.Add("Transkript gestartet: $transcriptPath")

    Write-Host "[2/8] Verbinde zu Exchange Online"
    try {
        Connect-ExchangeOnline -ErrorAction Stop
        $poslog.Add("Verbindung zu Exchange Online aufgebaut")
    }
    catch {
        $errorMessage = "Fehler bei Connect-ExchangeOnline: $($_.Exception.Message)"
        $neglog.Add($errorMessage)
        throw
    }

    Write-Host "[3/8] Lese aktuelle VIP-Liste"
    $vipusers = Get-VipUserList -Export
    $vipusers | Sort-Object DisplayName |
        Select-Object DisplayName, WindowsEmailAddress, UserPrincipalName |
        Format-Table -AutoSize
    $poslog.Add("Aktuelle VIP-Liste gelesen: $($vipusers.Count) Eintraege")

    Write-Host "[4/8] Entferne VIP-Tag von aktueller Liste"
    foreach ($vipuser in $vipusers) {
        try {
            $vipuser | Set-User -VIP $false -Confirm:$false -ErrorAction Stop
            $alias = $vipuser.alias
            $displayname = $vipuser.displayname
            Write-Host "Removed VIP tag from user $displayname with alias $alias"
            $poslog.Add("VIP tag entfernt: $displayname ($alias)")
        }
        catch {
            $alias = $vipuser.alias
            $displayname = $vipuser.displayname
            $errorMessage = "Fehler beim Entfernen des VIP tags fuer $displayname ($alias): $($_.Exception.Message)"
            Write-Host $errorMessage -ForegroundColor Red
            $neglog.Add($errorMessage)
        }
    }

    Write-Host "[5/8] Ermittle neue VIP-User aus Gruppen"
    try {
        $fe0users = Get-DistributionGroup -Filter $prioritygroupfilter1 |
            ForEach-Object { Get-DistributionGroupMember $_.name -ErrorAction Stop }
        $fe1users = Get-DistributionGroup -Filter $prioritygroupfilter2 |
            ForEach-Object { Get-DistributionGroupMember $_.name -ErrorAction Stop }
        $poslog.Add("Gruppenmitglieder gelesen: FE0=$($fe0users.Count), FE1=$($fe1users.Count)")
    }
    catch {
        $errorMessage = "Fehler beim Lesen der konfigurierten Prioritaetsgruppen: $($_.Exception.Message)"
        $neglog.Add($errorMessage)
        throw
    }

    Write-Host "[6/8] Setze neue VIP-Tags"
    Set-VipTag -userlist $fe0users
    Set-VipTag -userlist $fe1users

    Write-Host "[7/8] Aktualisiere Impersonation-Schutz"
    $updatedVipUsers = Get-VipUserList -Export
    try {
        $date = Get-Date -Format FileDateTime
        $TargetedUsersToProtect_Current = (Get-AntiPhishPolicy $phishingpolicyname -ErrorAction Stop).TargetedUsersToProtect
        $TargetedUsersToProtect_Current | Out-File "$exportfilepath\$date-impersonation-users.txt"
        $TargetedUsersToProtect_Current
        $poslog.Add("Aktuelle Impersonation-Liste exportiert nach $exportfilepath\\$date-impersonation-users.txt")

        $TargetedUsersToProtect_New = foreach ($vipuser in $updatedVipUsers) {
            $vipuser.DisplayName, $vipuser.UserPrincipalName -join ";"
        }

        if ($addimpersonationusers -eq $true) {
            $TargetedUsersToProtect_combined = $TargetedUsersToProtect_Current + $TargetedUsersToProtect_New
            $TargetedUsersToProtect = $TargetedUsersToProtect_combined | Sort-Object -Unique
            Set-AntiPhishPolicy -Identity $phishingpolicyname -TargetedUsersToProtect $TargetedUsersToProtect -ErrorAction Stop
            Write-Host "Adding the new users to protect from user impersonation to the current list:"
            $TargetedUsersToProtect
            $poslog.Add("Impersonation-Liste erweitert: $($TargetedUsersToProtect.Count) Eintraege gesamt")
        }
        else {
            Set-AntiPhishPolicy -Identity $phishingpolicyname -TargetedUsersToProtect $TargetedUsersToProtect_New -ErrorAction Stop
            Write-Host "Overwriting the current list of users to protect from user impersonation with the new one:"
            $TargetedUsersToProtect_New
            $poslog.Add("Impersonation-Liste ueberschrieben: $($TargetedUsersToProtect_New.Count) Eintraege")
        }
    }
    catch {
        $errorMessage = "Fehler beim Aktualisieren des Impersonation-Schutzes: $($_.Exception.Message)"
        $neglog.Add($errorMessage)
        throw
    }
}
catch {
    $errorMessage = "Skriptlauf abgebrochen: $($_.Exception.Message)"
    Write-Host $errorMessage -ForegroundColor Red
    $neglog.Add($errorMessage)
}
finally {
    Write-Host "[8/8] Abschluss und Log-Ausgabe"

    try {
        $poslog | ConvertTo-Json | Out-File $poslogJsonPath
        $poslog.Add("PosLog als JSON exportiert: $poslogJsonPath")
    }
    catch {
        $jsonError = "Fehler beim Export von PosLog als JSON: $($_.Exception.Message)"
        Write-Host $jsonError -ForegroundColor Red
        $neglog.Add($jsonError)
    }

    try {
        $neglog | ConvertTo-Json | Out-File $neglogJsonPath
        $poslog.Add("NegLog als JSON exportiert: $neglogJsonPath")
    }
    catch {
        $jsonError = "Fehler beim Export von NegLog als JSON: $($_.Exception.Message)"
        Write-Host $jsonError -ForegroundColor Red
        $neglog.Add($jsonError)
    }

    Write-Host "Positive Log-Eintraege: $($poslog.Count)"
    $poslog | ForEach-Object { Write-Host "[POS] $_" }

    Write-Host "Negative Log-Eintraege: $($neglog.Count)"
    $neglog | ForEach-Object { Write-Host "[NEG] $_" -ForegroundColor Red }

    if ($transcriptStarted) {
        Stop-Transcript
    }
}