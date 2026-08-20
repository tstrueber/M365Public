# Verbindung zu Active Directory und Microsoft Graph
# Benötigte Module: ActiveDirectory, Microsoft.Graph.Users
Connect-MgGraph -Scopes "User.ReadWrite.All" -NoWelcome

$OuDn = "OU=Benutzer,DC=contoso,DC=local"

$Zeitstempel    = Get-Date -Format "yyyyMMdd_HHmmss"
$TranscriptPath = Join-Path $PSScriptRoot "ImmutableId_Set_$Zeitstempel.log"

Start-Transcript -Path $TranscriptPath

function Escape-ODataValue {
    param([string]$Value)

    return $Value -replace "'", "''"
}

try {
    $OnPremUsers = Get-ADUser -SearchBase $OuDn -Filter * `
        -Properties UserPrincipalName, Mail, GivenName, Surname, ObjectGUID, mS-DS-ConsistencyGuid `
        -ErrorAction Stop

    foreach ($User in $OnPremUsers) {
        Write-Host "`nPrüfe: $($User.SamAccountName)" -ForegroundColor Cyan

        try {
            # Source Anchor bevorzugt aus mS-DS-ConsistencyGuid;
            # falls nicht belegt, wird ObjectGUID verwendet.
            if ($User.'mS-DS-ConsistencyGuid') {
                $ImmutableId = [Convert]::ToBase64String(
                    [byte[]]$User.'mS-DS-ConsistencyGuid'
                )
            }
            else {
                $ImmutableId = [Convert]::ToBase64String(
                    $User.ObjectGUID.ToByteArray()
                )
            }

            $CloudUser = $null

            # 1. Suche über den On-Premises-UPN
            if ($User.UserPrincipalName) {
                try {
                    $CloudUser = Get-MgUser `
                        -UserId $User.UserPrincipalName `
                        -Property Id,UserPrincipalName,Mail,OnPremisesImmutableId,OnPremisesSyncEnabled `
                        -ErrorAction Stop
                }
                catch {
                    # Nicht gefunden – weitere Suchwege folgen.
                }
            }

            # 2. Suche über Mail-Attribut
            if (-not $CloudUser -and $User.Mail) {
                $Mail = Escape-ODataValue $User.Mail

                $CloudUser = Get-MgUser `
                    -Filter "mail eq '$Mail'" `
                    -ConsistencyLevel eventual `
                    -Property Id,UserPrincipalName,Mail,OnPremisesImmutableId,OnPremisesSyncEnabled `
                    -All `
                    -ErrorAction Stop |
                    Select-Object -First 1
            }

            # 3. Suche über Vor- und Nachnamen
            if (-not $CloudUser -and $User.GivenName -and $User.Surname) {
                $GivenName = Escape-ODataValue $User.GivenName
                $Surname   = Escape-ODataValue $User.Surname

                $CloudUser = Get-MgUser `
                    -Filter "givenName eq '$GivenName' and surname eq '$Surname'" `
                    -ConsistencyLevel eventual `
                    -Property Id,UserPrincipalName,Mail,OnPremisesImmutableId,OnPremisesSyncEnabled `
                    -All `
                    -ErrorAction Stop |
                    Select-Object -First 1
            }

            if (-not $CloudUser) {
                Write-Host "Nicht gefunden: $($User.SamAccountName)" -ForegroundColor Yellow
                continue
            }

            Write-Host "Gefunden: $($CloudUser.UserPrincipalName)" -ForegroundColor Green

            # Für bereits synchronisierte Benutzer kann die Immutable ID nicht per Graph gesetzt werden.
            if ($CloudUser.OnPremisesSyncEnabled) {
                Write-Host "Übersprungen: Benutzer wird bereits synchronisiert." -ForegroundColor Yellow
                continue
            }

            if ($CloudUser.OnPremisesImmutableId) {
                Write-Host "Übersprungen: Immutable ID ist bereits gesetzt ($($CloudUser.OnPremisesImmutableId))." -ForegroundColor Yellow
                continue
            }

            Update-MgUser `
                -UserId $CloudUser.Id `
                -OnPremisesImmutableId $ImmutableId `
                -ErrorAction Stop

            Write-Host "Immutable ID gesetzt: $ImmutableId" -ForegroundColor Green
        }
        catch {
            Write-Host "Fehler bei $($User.SamAccountName): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}
catch {
    Write-Host "Fehler beim Auslesen der OU: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Stop-Transcript
}