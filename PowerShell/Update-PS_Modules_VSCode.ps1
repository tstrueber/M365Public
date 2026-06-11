# Aktualisieren von Powershell 7, den gängigen Powershell Modulen und VSCode
# variablen
param(
  [switch]$PostPowerShellUpdateCheck,
  [string]$ExpectedPowerShellVersion
)

$Scriptpath = "C:\SCRIPTS\PROD\Update_VSCode_PS_Modules"
$Scriptname = "Update_VSVode_PS_Modules"
$temppath = "C:\Temp" #temp path for downloading the setup files
$psmodulepath = "C:\Program Files\PowerShell\Modules"
$tempPsmodulepath = "C:\Temp\InstallPSModules"

function Get-GitHubVersion($url)
{
  $counter = 1
  $maxtries = 15
  $waittimeseconds = 60
  $GitHubversion = $null
  do {
    Write-Host "Versuche jetzt die Version von Github abzurufen - Versuch $counter / $maxtries"    
    $request = [System.Net.WebRequest]::Create($url)
    #$request.Proxy = New-Object System.Net.WebProxy($proxyserver, $proxyport)
    try{
      $response = $request.GetResponse()
      $realTagUrl = $response.ResponseUri.OriginalString
      $GitHubversion = $realTagUrl.split('/')[-1].Trim('v')
      Write-Host "Version erfolgreich von GitHub abgerufen: $GitHubversion"
      return $GitHubversion
    } 
    catch{
      Write-Host "$(get-date -Format "HH:mm:ss" ) Fehler beim Abruf der GitHub Versionsnummer! Warte jetzt $waittimeseconds Sekunden. Fehler: $psitem" -ForegroundColor Red
      Start-Sleep -Seconds $waittimeseconds
    }
    if($counter -eq 15)
    { Write-Host "Maximale Zahl an Versuche erreicht, breche ab!" }
    $counter++
  }
  while (
    $counter -le $maxtries `
    -and $null -eq $GitHubversion )

  return "ERROR"
}

# Start 
Set-location $Scriptpath
if (-not (Test-Path -Path ".\LOG")) {
  New-Item -Path ".\LOG" -ItemType Directory -Force | Out-Null
}
if (-not (Test-Path -Path $temppath)) {
  New-Item -Path $temppath -ItemType Directory -Force | Out-Null
}
Start-Transcript -Path ".\LOG\$(Get-Date -format "yyyyMMdd_HHmmss")_$($Scriptname.Split(".")[0])_TS.LOG"

# check if we are running in VSCode
if ($Host.Name -match 'Visual Studio Code') {
  Write-Host "Running in VS Code (PowerShell extension) -> Script will be terminated!" -ForegroundColor Red
  Write-Host "Please run in elevated PS7 Shell!" -ForegroundColor Red
  Read-Host "Press Enter to exit"
  # now exit the script
  Stop-Transcript
  exit 1
}

Write-host "Check that we have administrator rights to install and update modules + use PS7..."
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
If (!$IsAdmin) {
    Write-Host "You must be signed in as an administrator to update modules" -ForegroundColor Red
    Stop-Transcript
    exit 1
}
If ($Host.Version.Major -lt 7) {
    Write-Host ("Current version is {0}. This script requires PowerShell 7" -f $Host.Version) -ForegroundColor Red
    Stop-Transcript
    exit 1
}

if ($PostPowerShellUpdateCheck) {
  $currentPwshVersion = $PSVersionTable.PSVersion.ToString()

  if ([string]::IsNullOrWhiteSpace($ExpectedPowerShellVersion)) {
    Write-Host "Post-Update-Check aktiv, aber keine Zielversion ubergeben." -ForegroundColor Red
    Stop-Transcript
    exit 1
  }

  [version]$installedVersion = $currentPwshVersion
  [version]$expectedVersion = $ExpectedPowerShellVersion

  if ($installedVersion -lt $expectedVersion) {
    Write-Host "PowerShell Re-Run-Check fehlgeschlagen. Installiert: $installedVersion ; Erwartet: $expectedVersion" -ForegroundColor Red
    Stop-Transcript
    exit 1
  }

  Write-Host "PowerShell Re-Run-Check erfolgreich. Installiert: $installedVersion ; Erwartet: $expectedVersion" -ForegroundColor Green
  Stop-Transcript
  exit 0
}

#region powershell modules
########################################################################################################################################################################################

# UpdateOffice365PowerShellModules.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/UpdateOffice365PowerShellModules.PS1
# Very simple script to check for updates to a defined set of PowerShell modules used to manage Office 365 services
# If an update for a module is found, it is downloaded and applied.
# Once all modules are checked for updates, we remove older versions that might be present on the workstation. V2.1 improves the processing of Microsoft Graph SDK sub-modules

# Define the set of modules installed and updated from the PowerShell Gallery that we want to maintain - edit this set of modules to include the modules 
# you want to process.
[int]$InstalledModules = 0; [int]$UpdatedModules = 0; [int]$RemovedModules = 0
#$O365Modules = @("PackageManagement", "Az.Accounts", "Az.Automation", "AIPService", "Az.Keyvault", "MicrosoftTeams", "Microsoft.Graph", "Microsoft.Graph.Beta", "ExchangeOnlineManagement", "Microsoft.Online.Sharepoint.PowerShell", "ORCA",  "Pnp.PowerShell", "MSCommerce", "Microsoft365DSC", "MSAL.PS", "WhiteboardAdmin", "ImportExcel")
$O365Modules = @("PackageManagement", "Microsoft.Graph", "Microsoft.Graph.Beta", "ExchangeOnlineManagement", "Pnp.PowerShell", "MSAL.PS", "ImportExcel")
$O365Modules = $O365Modules 
Write-Host ("Starting up and preparing to process these modules: {0}" -f ($O365Modules -join ", ")) -foregroundcolor Yellow
[int]$UpdatedModules = 0; [int]$RemovedModules = 0; [int]$InstalledModules = 0

# We're installing from the PowerShell Gallery so make sure that it's trusted
Set-PSRepository -Name PsGallery -InstallationPolicy Trusted

Write-Host  ""
Write-Host  ""
Write-Host "######### Check and update all modules to make sure that we're at the latest version" -ForegroundColor Cyan
ForEach ($Module in $O365Modules) 
{
   Write-Host "Checking and updating module" $Module
   $CurrentModule = Find-Module -Name $Module -Repository PSGallery
   If ($CurrentModule) {
     $CurrentVersion = $CurrentModule.Version
      If ($CurrentVersion -isnot [string]) {
        $CurrentVersion = $CurrentVersion.Major.toString() + "." + $CurrentVersion.Minor.toString() + "." + $CurrentVersion.Build.toString()
      }
     Write-Host ("Current version of the {0} module in the PowerShell Gallery is {1}" -f $Module, $CurrentVersion)
   }

   # Check if the module is installed on this PC
   Clear-Variable PCModule -WarningAction SilentlyContinue -ErrorAction SilentlyContinue # make sure we start fresh
   $PCModule = Get-Module -Name $Module -ListAvailable -ErrorAction SilentlyContinue 

   If (!($PCModule)) { 
   # No version of the module found. It's in our list, so we install it.
     Write-Host ("No module found on this PC for {0}" -f $Module)
     Write-Host ("Installing module {0}..." -f $Module)  -foregroundcolor Yellow
     #Install-Module $Module -Scope AllUsers -Confirm:$False -AllowClobber -Force -Repository PSGallery
     try {
      Save-Module $Module -Path $psmodulepath -Force -ErrorAction Stop -WarningAction Continue
      Write-Host "Module $Module installed in CommonUse folder" -ForegroundColor Green
     }
     catch {
      write-host "Installation of Module $Module failed" -ForegroundColor Red
      }
     $InstalledModules++
   }
   
   If ($PCModule) 
   {
      # We have at least one version of the module installed. Check if it's the latest version
      $PCModule = $PCModule | Sort-Object version -Descending #absteigend sortieren

      foreach($PCModuleEntry in $PCModule)
      {
        $PCVersion = $PCModuleEntry.Version
        If ($PCVersion -isnot [string])
        {
          if($null -eq $PCVersion.Revision -or $PCVersion.Revision -eq -1) 
          {
            # Revision is missing or -1
            $PCVersion = $PCVersion.Major.toString() + "." + $PCVersion.Minor.toString() + "." + $PCVersion.Build.toString() 
          } 
          else
          {
            $PCVersion = $PCVersion.Major.toString() + "." + $PCVersion.Minor.toString() + "." + $PCVersion.Build.toString() + "." + $PCVersion.Revision.ToString() 
          }           
        }
        If ($PCVersion -eq $CurrentVersion)
        { 
          Write-Host ("Latest version of $Module is installed on this PC - no need to update" ) 
          Write-host "Module Path = $($PCModuleEntry.Path)"
        } 
        Else
        {
          Write-Host ("Installing the latest version of the module. Installed Version: $PCVersion ; Available Version: $CurrentVersion") -foregroundcolor Green
          try 
          {
            # speichern im Temp Ordner
            Save-Module $Module -Path $tempPsmodulepath -Force -Repository PSGallery -ErrorAction Stop -WarningAction Continue
            # Kopieren in den richtigen Ordner
            $ModuleFolders = Get-ChildItem -Path $tempPsmodulepath -Directory

            foreach ($entry in $ModuleFolders) {
                $SourceModulePath = $entry.FullName
                $DestinationModulePath = Join-Path -Path $psmodulepath -ChildPath $entry.Name

                # Kopiere rekursiv alle Inhalte in den Zielmodulordner, überschreibe vorhandene Dateien
                try {
                  Copy-Item -Path "$SourceModulePath\*" -Destination $DestinationModulePath -Recurse -Force -WarningAction Continue -ErrorAction Stop
                  Write-Host "Kopieren von $SourceModulePath nach $DestinationModulePath erfolgreich." -ForegroundColor Green
                }
                catch {
                  Write-Host "Fehler beim Kopieren von $SourceModulePath nach $DestinationModulePath - Fehler: $PSItem" -ForegroundColor Red
                }
            }

            # löschen der Temporären Ordner
            Get-ChildItem -Path $tempPsmodulepath -Directory | Remove-Item -Recurse -Force

            $UpdatedModules++
            Write-Host "Neue Modulversion von $Module im CommonUse Ordner abgelegt!"
          }
          catch 
          {
            write-host "Update of Module $Module failed - Error: $psitem" -ForegroundColor Red
          }          
        } 
      }
    }# End if 
} # End ForEach Module

Write-Host  ""
Write-Host  ""
Write-Host  "############ Check and remove older versions of the modules from the System..." -ForegroundColor Cyan
[array]$SetofInstalledModules = Get-InstalledModule
[array]$GraphModules = $SetOfInstalledModules | Where-Object {$_.Name -Like "*Microsoft.Graph*"} | Select-Object -ExpandProperty Name
$ModulesToProcess = $O365Modules + $GraphModules | Sort-Object -Unique

ForEach ($Module in $ModulesToProcess)
{
   Write-Host "Checking for older versions of" $Module
   [array]$AllVersions = Get-Module -Name $Module -ListAvailable -ErrorAction SilentlyContinue
   If ($AllVersions) 
   {
     $AllVersions = $AllVersions | Sort-Object Version -Descending 
     $MostRecentVersion = $AllVersions[0].Version
     If ($MostRecentVersion -isnot [string]) 
     { # Handle PowerShell 5 - PowerShell 7 returns a string
        $MostRecentVersion = $MostRecentVersion.Major.toString() + "." + $MostRecentVersion.Minor.toString() + "." + $MostRecentVersion.Build.toString()
     }
     Write-Host ("Most recent version of $Module is $MostRecentVersion" )
     If ($AllVersions.Count -gt 1 ) 
     { # More than a single version installed
      ForEach ($Version in $AllVersions) 
      { #Check each version and remove old versions
        If ($Version.version -lt $MostRecentVersion) 
        { # Old version - remove
           Write-Host ("Uninstalling version {0} of module {1}" -f $Version.Version, $Module) -foregroundcolor Red 
           
          # Löschen des Modul-Ordners inkl. aller enthaltenen Dateien
          if (Test-Path -Path $Version.ModuleBase) {
            try {
                Remove-Item -Path $Version.ModuleBase -Recurse -Force -WarningAction Continue -ErrorAction Stop
                Write-Host "Ordner $($Version.ModuleBase) wurde erfolgreich gelöscht." -ForegroundColor Green
            } catch {
                Write-Host "Fehler beim Löschen des Ordners $($Version.ModuleBase) - Fehler: $Psitem" -ForegroundColor Red
            }
          } else {
            Write-Host "Ordner $($Version.ModuleBase) existiert nicht." -ForegroundColor Yellow
          }
           $RemovedModules++
         } #End if version check
      } # End ForEach versions 
    } 
    Else 
    {
        Write-Host ("No earlier versions of {0} module to remove" -f $Module)
    } # End check for more than one version
  } #End If
} #End ForEach

Write-Host ("Installed modules: {0} Updated modules: {1}  Removed old versions of modules: {2}" -f $InstalledModules, $UpdatedModules, $RemovedModules)

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.

#endregion powershell modules


#region vscode
########################################################################################################################################################################################

Write-Host  ""
Write-Host  ""
Write-Host "########## check if newer Visual Studio Code Version is available" -ForegroundColor Cyan
$GitHubversion = $null
$localversion = $null

# Version abrufen, Fehler wird in der Funktion abgefangen!
$GitHubversion = Get-GitHubVersion -url "https://github.com/microsoft/vscode/releases/latest"

$localversion = (Get-ChildItem "C:\Program Files\Microsoft VS Code\Code.exe").VersionInfo.ProductVersion
if($GitHubversion -eq "ERROR" -or [string]::IsNullOrWhiteSpace($GitHubversion)){
  Write-Host "Fehler beim Abruf der aktuellen Version von GitHub. Daher kein Update."
}
elseif($GitHubversion -ne $localversion) 
{
  Write-Host "Neuere Version auf Github gefunden: $GitHubversion"
  Write-Host "Lokale Version: $localversion"
  Write-Host "Aktualisiere VSCode..."
  # Define the download URL and the destination
  $Destination = "$temppath\vscode_installer.exe"
  $VSCodeUrl = "https://code.visualstudio.com/sha/download?build=stable&os=win32-x64"

  # Close any running instances of VS Code
  Write-Host "Closing any running instances of VS Code..."
  Get-Process -Name "Code" -ErrorAction SilentlyContinue | ForEach-Object { $_.CloseMainWindow(); Stop-Process -Id $_.Id -Force }

  # Download VSCode installer
  Write-Host "Downloading VSCode..."
  Invoke-WebRequest -Uri $VSCodeUrl -OutFile $Destination -SkipCertificateCheck

  # Install VS Code silently
  Write-Host "Installing VSCode..."
  try 
  {
    Start-Process -FilePath $Destination -ArgumentList '/verysilent /mergetasks=!runcode' -Wait -Passthru
    Write-Host "Installation successfully started in background"
  }
  catch 
  {
    Write-Host "error starting vscode installation in background"
  }

  # Remove installer
  Write-Host "Removing installation file..."
  Remove-Item $Destination

  Write-Host "VSCode installation completed!"
}
else 
{
  Write-Host "Aktuellste VSCode Version ist installiert $GitHubversion - kein Update notwendig!" -ForegroundColor Green
}

#endregion vscode

#region vscode extensions
##############################################################################################################################################################################

#endregion vscode extensions
Write-Host "Aktuelle Versionen der VSCode Extensions:"
code --list-extensions --show-versions

Write-Host "Ein Update der VSCode Extensions funktioniert aktuell noch nicht automatisch..."

#Write-Host "Starte Update der VSCode Extensions"
#code --update-extensions

#region powershell
##############################################################################################################################################################################

Write-Host  ""
Write-Host  ""
Write-Host "############ check if newer version of PowerShell is available..."
# check if newer PowerShell Version is available
$GitHubversion = $null
$localversion = $null

# Version abrufen, Fehler wird in der Funktion abgefangen!
$GitHubversion = Get-GitHubVersion -url "https://github.com/PowerShell/PowerShell/releases/latest"

$localversion = $PSVersionTable.PSVersion.Major.ToString() +"."+ $PSVersionTable.PSVersion.Minor.ToString() +"."+ $PSVersionTable.PSVersion.Patch.ToString()
if($GitHubversion -eq "ERROR" -or [string]::IsNullOrWhiteSpace($GitHubversion)){
  Write-Host "Fehler beim Abruf der aktuellen Version von GitHub. Daher kein Update."
}
elseif($GitHubversion -ne $localversion) 
{
  Write-Host "Neuere Version auf Github gefunden: $GitHubversion"
  Write-Host "Lokale Version: $localversion"
  Write-Host "Aktualisiere PowerShell in einem separaten Prozess..."

  # Das Update wird in einen separaten Windows-PowerShell-Prozess ausgelagert,
  # damit die laufende pwsh-Session nicht während des MSI-Updates abstürzt.
  $PowerShellUpdateLogPath = Join-Path -Path $Scriptpath -ChildPath "LOG\$(Get-Date -format \"yyyyMMdd_HHmmss\")_PowerShellUpdate.log"
  $PowerShellUpdaterScriptPath = Join-Path -Path $env:TEMP -ChildPath "Update-PowerShell7-Detached.ps1"

  $PowerShellUpdaterScript = @'
param(
  [Parameter(Mandatory = $true)]
  [int]$ParentProcessId,
  [Parameter(Mandatory = $true)]
  [string]$TargetVersion,
  [Parameter(Mandatory = $true)]
  [string]$LogPath,
  [Parameter(Mandatory = $true)]
  [string]$MainScriptPath
)

$ErrorActionPreference = "Stop"

function Write-UpdateLog {
  param(
    [string]$Message,
    [string]$Level = "INFO"
  )

  $line = "$(Get-Date -Format \"yyyy-MM-dd HH:mm:ss\") [$Level] $Message"
  Add-Content -Path $LogPath -Value $line
}

Write-UpdateLog -Message "Detached updater started. Waiting for parent process $ParentProcessId to exit."
while (Get-Process -Id $ParentProcessId -ErrorAction SilentlyContinue) {
  Start-Sleep -Seconds 2
}

try {
  Write-UpdateLog -Message "Starting PowerShell installation via install-powershell.ps1"
  & ([ScriptBlock]::Create((Invoke-RestMethod -Uri "https://aka.ms/install-powershell.ps1"))) -UseMSI -Quiet

  # Kurze Wartezeit, damit Dateien/Registry nach dem MSI-Setup stabil verfugbar sind.
  Start-Sleep -Seconds 5

  $pwshCommand = Get-Command -Name "pwsh.exe" -ErrorAction SilentlyContinue
  if (-not $pwshCommand) {
    throw "pwsh.exe was not found after the update."
  }

  $installedVersionString = & $pwshCommand.Source -NoLogo -NoProfile -Command '$PSVersionTable.PSVersion.ToString()'
  if (-not $installedVersionString) {
    throw "Could not read installed PowerShell version."
  }

  [version]$installedVersion = $installedVersionString
  [version]$requiredVersion = $TargetVersion

  if ($installedVersion -lt $requiredVersion) {
    throw "Installed version $installedVersion is lower than required version $requiredVersion."
  }

  Write-UpdateLog -Message "PowerShell update successful. Installed version: $installedVersion"

  $reRunArgs = @(
    "-NoLogo",
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", "`"$MainScriptPath`"",
    "-PostPowerShellUpdateCheck",
    "-ExpectedPowerShellVersion", "`"$TargetVersion`""
  )

  Write-UpdateLog -Message "Starting main script rerun in pwsh for post-update validation."
  $reRunProcess = Start-Process -FilePath $pwshCommand.Source -ArgumentList $reRunArgs -PassThru -Wait -WindowStyle Normal

  if ($reRunProcess.ExitCode -ne 0) {
    throw "Main script rerun failed with exit code $($reRunProcess.ExitCode)."
  }

  Write-UpdateLog -Message "Main script rerun finished successfully."
  exit 0
}
catch {
  Write-UpdateLog -Level "ERROR" -Message "PowerShell update failed: $($_.Exception.Message)"
  exit 1
}
'@

  Set-Content -Path $PowerShellUpdaterScriptPath -Value $PowerShellUpdaterScript -Encoding UTF8 -Force

  $PowerShellUpdaterArgs = @(
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", "`"$PowerShellUpdaterScriptPath`"",
    "-ParentProcessId", $PID,
    "-TargetVersion", $GitHubversion,
    "-LogPath", "`"$PowerShellUpdateLogPath`"",
    "-MainScriptPath", "`"$PSCommandPath`""
  )

  Start-Process -FilePath "powershell.exe" -ArgumentList $PowerShellUpdaterArgs -WindowStyle Hidden | Out-Null

  Write-Host "PowerShell-Update wurde im Hintergrund gestartet."
  Write-Host "Prufprotokoll: $PowerShellUpdateLogPath"
  Write-Host "Dieses Skript wird jetzt beendet, damit die laufende pwsh-Instanz das Update nicht blockiert." -ForegroundColor Yellow
  Stop-Transcript
  exit 0
}
else
{
  Write-Host "Aktuellste PowerShell Version ist installiert $GitHubversion - kein Update notwendig!" -ForegroundColor Green
}

#endregion powershell

Stop-Transcript
