# Update PowerShell 7, common PowerShell modules, and VS Code
# Version 2.3 - 2026-11-12
# for the most recent version of the script check out the public github repository: https://github.com/tstrueber/M365Public/blob/main/PowerShell/Update-PS_Modules_VSCode.ps1
#region Variables
param(
  [switch]$PostPowerShellUpdateCheck,
  [string]$ExpectedPowerShellVersion,
  [switch]$UpdatePSOnly,
  [switch]$UpdateModulesOnly,
  [switch]$UpdateVSCodeOnly,
  [switch]$UpdateGitOnly
)

$ScriptPath = if ($PSCommandPath) { $PSCommandPath } 
              elseif ($MyInvocation.MyCommand.Path) { $MyInvocation.MyCommand.Path } 
              elseif ($psEditor) { $psEditor.GetEditorContext().CurrentFile.Path }
$ScriptDirectory = Split-Path -Parent $ScriptPath
$Logpath = join-path $ScriptDirectory "LOG"
$Scriptname = [System.IO.Path]::GetFileNameWithoutExtension($ScriptPath) # extract scriptname from $ScriptPath and remove the extension
$temppath = "C:\Temp" #temp path for downloading the setup files
$VSCodeInstallPath = "C:\Program Files\Microsoft VS Code" # system-wide VS Code installation
$psmodulepath = "C:\Program Files\PowerShell\Modules" # where the updated PowerShell modules are installed
$tempPsmodulepath = "C:\Temp\InstallPSModules" # temporary path for installing PowerShell modules

# Define the set of modules installed and updated from the PowerShell Gallery that we want to maintain - edit this set of modules to include the modules 
# you want to process.
#$O365Modules = @("PackageManagement", "Az.Accounts", "Az.Automation", "AIPService", "Az.Keyvault", "MicrosoftTeams", "Microsoft.Graph", "Microsoft.Graph.Beta", "ExchangeOnlineManagement", "Microsoft.Online.Sharepoint.PowerShell", "ORCA",  "Pnp.PowerShell", "MSCommerce", "Microsoft365DSC", "MSAL.PS", "WhiteboardAdmin", "ImportExcel")
$O365Modules = @("PackageManagement", "AIPService", "MicrosoftTeams", "Microsoft.Graph", "Microsoft.Graph.Beta", "ExchangeOnlineManagement", "Microsoft.Online.Sharepoint.PowerShell", "ORCA",  "Pnp.PowerShell", "MSCommerce", "MSAL.PS", "WhiteboardAdmin", "ImportExcel")

# Versions that must NOT be installed (per module). Optionally, "*" can be used as a global key.
# Example:
# "ExchangeOnlineManagement" = @("3.10.1")
# If the blocked version is already installed locally, it will be removed.
# The script then automatically uses the newest allowed PSGallery version as target.
$ModuleVersionBlockList = @{
   "ExchangeOnlineManagement" = @("3.10.0")
}

#endregion Variables

#region functions
####################################################################################################################################
function Get-GitHubVersion($url)
{
  $counter = 1
  $maxtries = 15
  $waittimeseconds = 60
  do {
    Write-Host "Retrieving version from GitHub - attempt $counter / $maxtries"
    try {
      $request = [System.Net.WebRequest]::Create($url)
      $response = $request.GetResponse()
      $realTagUrl = $response.ResponseUri.OriginalString
      $response.Close()
      $GitHubversion = $realTagUrl.split('/')[-1].Trim('v')
      Write-Host "Successfully retrieved version from GitHub: $GitHubversion"
      return $GitHubversion
    }
    catch {
      Write-Host "$(get-date -Format "HH:mm:ss") Failed to retrieve GitHub version! Error: $psitem" -ForegroundColor Red
      if ($counter -lt $maxtries) {
        Write-Host "Waiting $waittimeseconds seconds before the next attempt." -ForegroundColor Red
        Start-Sleep -Seconds $waittimeseconds
      }
      else {
        Write-Host "Maximum number of attempts reached, giving up!" -ForegroundColor Red
      }
    }
    $counter++
  }
  while ($counter -le $maxtries)
  return $null
}

function Save-FileDownload
{
  param(
    [Parameter(Mandatory)][string]$Url,
    [Parameter(Mandatory)][string]$Destination,
    [int]$MaxTries = 60,
    [int]$WaitTimeSeconds = 5,
    [switch]$SkipCertificateCheck
  )
  # The progress bar slows down Invoke-WebRequest considerably for large files.
  $ProgressPreference = 'SilentlyContinue'
  $counter = 1
  $success = $false
  do {
    Write-Host "$(Get-Date -Format "HH:mm:ss") Attempt $counter / $MaxTries : downloading $Url"
    try {
      $iwrParams = @{
        Uri                      = $Url
        OutFile                  = $Destination
        ErrorAction              = 'Stop'
        OperationTimeoutSeconds  = 60
        ConnectionTimeoutSeconds = 60
      }
      if ($SkipCertificateCheck) { $iwrParams.SkipCertificateCheck = $true }
      Invoke-WebRequest @iwrParams

      if (-not (Test-Path -Path $Destination) -or (Get-Item -Path $Destination).Length -eq 0) {
        throw "The downloaded file is missing or empty: $Destination"
      }
      $success = $true
      Write-Host "Download completed successfully: $Destination" -ForegroundColor Green
    }
    catch {
      Write-Host "$(Get-Date -Format "HH:mm:ss") Download failed! Error: $PSItem" -ForegroundColor Red
      if (Test-Path -Path $Destination) {
        Remove-Item -Path $Destination -Force -ErrorAction SilentlyContinue
      }
      if ($counter -lt $MaxTries) {
        Write-Host "Waiting $WaitTimeSeconds seconds before the next attempt." -ForegroundColor Red
        Start-Sleep -Seconds $WaitTimeSeconds
      }
      else {
        Write-Host "Maximum number of attempts reached, giving up!" -ForegroundColor Red
      }
    }
    $counter++
  }
  while (-not $success -and $counter -le $MaxTries)
  return $success
}

function Get-BlockedVersionsForModule {
  param(
    [string]$ModuleName,
    [hashtable]$BlockedVersionMap
  )

  $blocked = @()

  if ($BlockedVersionMap.ContainsKey("*")) {
    $blocked += @($BlockedVersionMap["*"])
  }
  if ($BlockedVersionMap.ContainsKey($ModuleName)) {
    $blocked += @($BlockedVersionMap[$ModuleName])
  }

  return @(
    $blocked |
    ForEach-Object { $_.ToString().Trim() } |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
    Sort-Object -Unique
  )
}

function Get-PreferredGalleryVersion {
  param(
    [string]$ModuleName,
    [string[]]$BlockedVersions
  )

  # Without blocked versions only the latest stable version is needed. Querying all versions is considerably
  # slower, which matters with the large number of Microsoft Graph sub-modules.
  if ($BlockedVersions.Count -gt 0) {
    [array]$galleryVersions = Find-Module -Name $ModuleName -Repository PSGallery -AllVersions -ErrorAction Stop
  }
  else {
    [array]$galleryVersions = Find-Module -Name $ModuleName -Repository PSGallery -ErrorAction Stop
  }

  $sortableVersions = foreach ($candidate in $galleryVersions) {
    $candidateVersionString = $candidate.Version.ToString()
    [version]$candidateVersion = $null

    if (-not [version]::TryParse($candidateVersionString, [ref]$candidateVersion)) {
      Write-Host ("Skipping unparsable PSGallery version '{0}' for module {1}" -f $candidateVersionString, $ModuleName) -ForegroundColor Yellow
      continue
    }

    [pscustomobject]@{
      Version = $candidateVersion
      VersionString = $candidateVersionString
    }
  }

  $sortedVersions = $sortableVersions | Sort-Object Version -Descending

  foreach ($candidate in $sortedVersions) {
    if ($BlockedVersions -contains $candidate.VersionString) {
      continue
    }
    return $candidate.Version
  }

  return $null
}

function Get-GraphSubModuleNames {
  param(
    [Parameter(Mandatory = $true)]
    [string]$MetaModuleName,
    [Parameter(Mandatory = $true)]
    [string]$ManagedModulePath
  )

  # The meta modules Microsoft.Graph and Microsoft.Graph.Beta only reference their sub-modules as dependencies.
  # Save-ModuleVersionToModulePath copies only the requested module folder, so every sub-module has to be
  # checked and updated on its own. Otherwise e.g. Microsoft.Graph can be current while
  # Microsoft.Graph.Applications is still outdated.
  $names = [System.Collections.Generic.List[string]]::new()
  $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

  try {
    $metaGalleryModule = Find-Module -Name $MetaModuleName -Repository PSGallery -ErrorAction Stop | Select-Object -First 1
    foreach ($dependency in $metaGalleryModule.Dependencies) {
      $dependencyName = [string]$dependency.Name
      if (-not [string]::IsNullOrWhiteSpace($dependencyName) -and $seen.Add($dependencyName)) {
        $names.Add($dependencyName)
      }
    }
    Write-Host ("{0} references {1} sub-modules in PSGallery." -f $MetaModuleName, $names.Count)
  }
  catch {
    Write-Host "Could not read the sub-modules of $MetaModuleName from PSGallery - Error: $PSItem" -ForegroundColor Red
  }

  # Additionally include locally installed sub-modules in the managed module path that are not (or no longer)
  # referenced by the meta module. Microsoft.Graph.Beta.* is excluded when processing Microsoft.Graph.
  $normalizedRoot = [System.IO.Path]::GetFullPath($ManagedModulePath).TrimEnd('\') + '\'
  $localModules = Get-Module -Name "$MetaModuleName.*" -ListAvailable -ErrorAction SilentlyContinue |
    Where-Object {
      ($MetaModuleName -ne 'Microsoft.Graph' -or $_.Name -notlike 'Microsoft.Graph.Beta*') -and
      ([System.IO.Path]::GetFullPath($_.ModuleBase).StartsWith($normalizedRoot, [System.StringComparison]::OrdinalIgnoreCase))
    }
  foreach ($localModule in $localModules) {
    if ($seen.Add($localModule.Name)) {
      $names.Add($localModule.Name)
    }
  }

  # Microsoft.Graph.Authentication is the base of all sub-modules, so process it first.
  $authentication = @($names | Where-Object { $_ -eq 'Microsoft.Graph.Authentication' })
  $others = @($names | Where-Object { $_ -ne 'Microsoft.Graph.Authentication' } | Sort-Object)
  return @($authentication + $others)
}

function Remove-InstalledModuleFolder {
  param(
    [Parameter(Mandatory = $true)]
    [string]$ModuleName,
    [Parameter(Mandatory = $true)]
    [version]$ModuleVersion,
    [Parameter(Mandatory = $true)]
    [string]$ModuleBasePath,
    [Parameter(Mandatory = $true)]
    [string]$AllowedRootPath
  )

  # Module versions are removed by deleting the corresponding version folder.
  # This is intentional because modules are deployed via Save-Module.
  # Only folders below the managed module path are touched. Modules shipped with PowerShell ($PSHOME\Modules),
  # Windows PowerShell or user scope modules are never deleted.
  $normalizedBase = [System.IO.Path]::GetFullPath($ModuleBasePath).TrimEnd('\') + '\'
  $normalizedRoot = [System.IO.Path]::GetFullPath($AllowedRootPath).TrimEnd('\') + '\'
  if (-not $normalizedBase.StartsWith($normalizedRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
    Write-Host ("Skipping {0} {1}: {2} is outside the managed module path {3}." -f $ModuleName, $ModuleVersion, $ModuleBasePath, $AllowedRootPath) -ForegroundColor Yellow
    return $false
  }

  if (-not (Test-Path -Path $ModuleBasePath)) {
    Write-Host "Folder $ModuleBasePath does not exist." -ForegroundColor Yellow
    return $false
  }

  try {
    Remove-Item -Path $ModuleBasePath -Recurse -Force -WarningAction Continue -ErrorAction Stop
    Write-Host ("Folder {0} (module {1} {2}) was removed successfully." -f $ModuleBasePath, $ModuleName, $ModuleVersion) -ForegroundColor Green
    return $true
  }
  catch {
    Write-Host "Failed to remove folder $ModuleBasePath - Error: $psitem" -ForegroundColor Red
    return $false
  }
}

function Save-ModuleVersionToModulePath {
  param(
    [Parameter(Mandatory = $true)]
    [string]$ModuleName,
    [Parameter(Mandatory = $true)]
    [version]$RequiredVersion,
    [Parameter(Mandatory = $true)]
    [string]$TempPath,
    [Parameter(Mandatory = $true)]
    [string]$TargetPath
  )

  if (Test-Path -Path $TempPath) {
    Get-ChildItem -Path $TempPath -Directory -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
  }
  else {
    New-Item -Path $TempPath -ItemType Directory -Force | Out-Null
  }

  # Save to a writable temp folder first to avoid permission issues with dependency folders in Program Files.
  Save-Module $ModuleName -Path $TempPath -RequiredVersion $RequiredVersion.ToString() -Force -Repository PSGallery -ErrorAction Stop -WarningAction Continue

  # Copy only the requested module folder/version into the global module path.
  $SavedModuleRoot = Join-Path -Path $TempPath -ChildPath $ModuleName
  if (-not (Test-Path -Path $SavedModuleRoot)) {
    throw "Saved module path was not found: $SavedModuleRoot"
  }

  $SavedVersions = Get-ChildItem -Path $SavedModuleRoot -Directory -ErrorAction Stop
  $TargetModuleRoot = Join-Path -Path $TargetPath -ChildPath $ModuleName
  if (-not (Test-Path -Path $TargetModuleRoot)) {
    New-Item -Path $TargetModuleRoot -ItemType Directory -Force | Out-Null
  }

  foreach ($VersionFolder in $SavedVersions) {
    $DestinationVersionPath = Join-Path -Path $TargetModuleRoot -ChildPath $VersionFolder.Name
    if (Test-Path -Path $DestinationVersionPath) {
      Remove-Item -Path $DestinationVersionPath -Recurse -Force -ErrorAction Stop
    }
    Copy-Item -Path $VersionFolder.FullName -Destination $TargetModuleRoot -Recurse -Force -ErrorAction Stop
    Write-Host "Copy from $($VersionFolder.FullName) to $DestinationVersionPath completed successfully." -ForegroundColor Green
  }

  Get-ChildItem -Path $TempPath -Directory -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
}

function Invoke-PowerShellZipInstall {
  param(
    [Parameter(Mandatory = $true)]
    [string]$ZipPath,
    [Parameter(Mandatory = $true)]
    [string]$DestinationPath,
    [Parameter(Mandatory = $true)]
    [string]$InstallLogPath
  )

  # The running pwsh.exe locks its own files in the target folder. The extraction is therefore done by an
  # elevated Windows PowerShell 5.1 helper process that waits until this session has ended.
  $q = { param([string]$s) "'" + $s.Replace("'", "''") + "'" }
  $helperScript = @"
`$ErrorActionPreference = 'Stop'
Start-Transcript -Path $(& $q $InstallLogPath) | Out-Null
try {
  Wait-Process -Id $PID -Timeout 300 -ErrorAction SilentlyContinue
  if (-not (Test-Path -Path $(& $q $DestinationPath))) { New-Item -Path $(& $q $DestinationPath) -ItemType Directory -Force | Out-Null }
  Expand-Archive -Path $(& $q $ZipPath) -DestinationPath $(& $q $DestinationPath) -Force
  Remove-Item -Path $(& $q $ZipPath) -Force -ErrorAction SilentlyContinue
  Write-Host 'PowerShell zip extraction completed successfully.'
}
catch {
  Write-Host "PowerShell zip extraction failed. Error: `$_"
  exit 1
}
finally {
  Stop-Transcript | Out-Null
}
"@

  $encodedCommand = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($helperScript))
  Start-Process -FilePath "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" `
    -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden', '-EncodedCommand', $encodedCommand) `
    -WindowStyle Hidden -ErrorAction Stop | Out-Null
}
#endregion functions

#region prerequisites
####################################################################################################################################

if (-not (Test-Path -Path $Logpath)) {
  New-Item -Path $Logpath -ItemType Directory -Force | Out-Null
}
Set-Location $Logpath
Start-Transcript -Path "$Logpath\$(Get-Date -format "yyyyMMdd_HHmmss")_$($Scriptname.Split(".")[0])_TS.LOG"
$TranscriptActive = $true

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
  # Compare Major.Minor.Patch only - [version] cannot parse preview suffixes like "7.6.0-preview.1".
  $installedVersion = [version]::new($PSVersionTable.PSVersion.Major, $PSVersionTable.PSVersion.Minor, $PSVersionTable.PSVersion.Patch)
  [version]$expectedVersion = $null

  if (-not [version]::TryParse(($ExpectedPowerShellVersion -replace '-.*$', ''), [ref]$expectedVersion)) {
    Write-Host "Post-update check is enabled, but no valid target version was provided ('$ExpectedPowerShellVersion')." -ForegroundColor Red
    Stop-Transcript
    exit 1
  }

  if ($installedVersion -lt $expectedVersion) {
    Write-Host "PowerShell rerun check failed. Installed: $installedVersion ; Expected: $expectedVersion" -ForegroundColor Red
    Stop-Transcript
    exit 1
  }

  Write-Host "PowerShell rerun check succeeded. Installed: $installedVersion ; Expected: $expectedVersion" -ForegroundColor Green
  Stop-Transcript
  exit 0
}

$HasScopedUpdateSelection = $UpdatePSOnly -or $UpdateModulesOnly -or $UpdateVSCodeOnly -or $UpdateGitOnly
if ($HasScopedUpdateSelection) {
  $RunPowerShellUpdate = [bool]$UpdatePSOnly
  $RunModuleUpdate = [bool]$UpdateModulesOnly
  $RunVSCodeUpdate = [bool]$UpdateVSCodeOnly
  $RunGitUpdate = [bool]$UpdateGitOnly
}
else {
  # Default behavior: run all update parts when no scope parameter is specified.
  $RunPowerShellUpdate = $true
  $RunModuleUpdate = $true
  $RunVSCodeUpdate = $true
  $RunGitUpdate = $true
}

Write-Host "Update selection -> Modules: $RunModuleUpdate ; VS Code: $RunVSCodeUpdate ; Git: $RunGitUpdate ; PowerShell: $RunPowerShellUpdate" -ForegroundColor Cyan
#endregion prerequisites

#region powershell modules
########################################################################################################################################################################################

# UpdateOffice365PowerShellModules.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/UpdateOffice365PowerShellModules.PS1
# Very simple script to check for updates to a defined set of PowerShell modules used to manage Office 365 services
# If an update for a module is found, it is downloaded and applied.
# Once all modules are checked for updates, we remove older versions that might be present on the workstation. V2.1 improves the processing of Microsoft Graph SDK sub-modules

[int]$InstalledModules = 0; [int]$UpdatedModules = 0; [int]$RemovedModules = 0

if ($RunModuleUpdate) {
  Write-Host ("Starting up and preparing to process these modules: {0}" -f ($O365Modules -join ", ")) -ForegroundColor Yellow

  # We're installing from the PowerShell Gallery so make sure that it's trusted
  Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

  # Expand the Microsoft Graph meta modules into their sub-modules. The sub-modules are processed before their
  # meta module, so each of them is checked and updated individually.
  $GraphMetaModuleNames = @("Microsoft.Graph", "Microsoft.Graph.Beta")
  $ModulesToCheck = [System.Collections.Generic.List[string]]::new()
  $ModulesToCheckSeen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($Module in $O365Modules) {
    if ($GraphMetaModuleNames -contains $Module) {
      Write-Host "Resolving sub-modules of $Module..."
      foreach ($SubModule in (Get-GraphSubModuleNames -MetaModuleName $Module -ManagedModulePath $psmodulepath)) {
        if ($ModulesToCheckSeen.Add($SubModule)) {
          $ModulesToCheck.Add($SubModule)
        }
      }
    }
    if ($ModulesToCheckSeen.Add($Module)) {
      $ModulesToCheck.Add($Module)
    }
  }
  Write-Host ("{0} modules (including Microsoft Graph sub-modules) will be checked." -f $ModulesToCheck.Count) -ForegroundColor Yellow

  Write-Host ""
  Write-Host ""
  Write-Host "######### Check and update all modules to make sure that we're at the latest version" -ForegroundColor Cyan
  foreach ($Module in $ModulesToCheck) {
    Write-Host ("Checking module {0}" -f $Module)
    $BlockedVersionsForModule = Get-BlockedVersionsForModule -ModuleName $Module -BlockedVersionMap $ModuleVersionBlockList
    if ($BlockedVersionsForModule.Count -gt 0) {
      Write-Host ("Blocked versions for {0}: {1}" -f $Module, ($BlockedVersionsForModule -join ", ")) -ForegroundColor Yellow
    }

    try {
      $CurrentVersion = Get-PreferredGalleryVersion -ModuleName $Module -BlockedVersions $BlockedVersionsForModule
    }
    catch {
      Write-Host "Could not read module versions from PSGallery for $Module - Error: $PSItem" -ForegroundColor Red
      continue
    }

    if ($null -eq $CurrentVersion) {
      Write-Host "No installable version found for module $Module (all versions are blocked)." -ForegroundColor Red
      continue
    }

    Write-Host ("Preferred (allowed) PSGallery version of {0} is {1}" -f $Module, $CurrentVersion)

    # Remove blocked versions immediately if they are already installed
    [array]$PCModule = Get-Module -Name $Module -ListAvailable -ErrorAction SilentlyContinue
    foreach ($InstalledEntry in $PCModule) {
      $InstalledVersionString = $InstalledEntry.Version.ToString()
      if ($BlockedVersionsForModule -contains $InstalledVersionString) {
        Write-Host ("Blocked version detected for {0}: {1}. Removing {2}" -f $Module, $InstalledVersionString, $InstalledEntry.ModuleBase) -ForegroundColor Yellow
        if (Remove-InstalledModuleFolder -ModuleName $Module -ModuleVersion $InstalledEntry.Version -ModuleBasePath $InstalledEntry.ModuleBase -AllowedRootPath $psmodulepath) {
          $RemovedModules++
        }
      }
    }

    # Re-read the installed versions after the cleanup above
    [array]$PCModule = Get-Module -Name $Module -ListAvailable -ErrorAction SilentlyContinue | Sort-Object Version -Descending

    if (-not $PCModule) {
      # No version of the module found. It's in our list, so we install it.
      Write-Host ("No module found on this PC for {0}" -f $Module)
      Write-Host ("Installing module {0} version {1}..." -f $Module, $CurrentVersion) -ForegroundColor Yellow
      try {
        Save-ModuleVersionToModulePath -ModuleName $Module -RequiredVersion $CurrentVersion -TempPath $tempPsmodulepath -TargetPath $psmodulepath
        Write-Host "Module $Module installed in the common module folder." -ForegroundColor Green
        $InstalledModules++
      }
      catch {
        Write-Host "Module installation failed for $Module - Error: $PSItem" -ForegroundColor Red
      }
      continue
    }

    $HighestInstalledVersion = $PCModule[0].Version
    if ($HighestInstalledVersion -eq $CurrentVersion) {
      Write-Host "Latest allowed version of $Module is installed on this PC - no need to update"
      Write-Host "Module Path = $($PCModule[0].Path)"
    }
    elseif ($HighestInstalledVersion -gt $CurrentVersion) {
      Write-Host ("Installed version {0} of {1} is newer than preferred PSGallery version {2}. Skipping downgrade." -f $HighestInstalledVersion, $Module, $CurrentVersion) -ForegroundColor Yellow
    }
    else {
      Write-Host ("Installing latest allowed version of module. Installed Version: {0} ; Allowed PSGallery Version: {1}" -f $HighestInstalledVersion, $CurrentVersion) -ForegroundColor Green
      try {
        Save-ModuleVersionToModulePath -ModuleName $Module -RequiredVersion $CurrentVersion -TempPath $tempPsmodulepath -TargetPath $psmodulepath
        $UpdatedModules++
        Write-Host "New allowed module version of $Module has been deployed to the Modules folder." -ForegroundColor Green
      }
      catch {
        Write-Host "Module update failed for $Module - Error: $PSItem" -ForegroundColor Red
      }
    }
  } # End foreach module

  Write-Host ""
  Write-Host ""
  Write-Host "############ Check and remove older versions of the modules from the System..." -ForegroundColor Cyan
  # Get-InstalledModule only lists modules installed via Install-Module. Modules here are deployed via
  # Save-Module, so the Graph sub-modules are discovered via Get-Module -ListAvailable instead.
  [array]$GraphModules = Get-Module -Name "Microsoft.Graph*" -ListAvailable -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name -Unique
  $ModulesToProcess = @($ModulesToCheck) + $GraphModules | Sort-Object -Unique

  foreach ($Module in $ModulesToProcess) {
    Write-Host "Checking for older versions of $Module"
    [array]$AllVersions = Get-Module -Name $Module -ListAvailable -ErrorAction SilentlyContinue | Sort-Object Version -Descending
    if (-not $AllVersions) {
      continue
    }

    $BlockedVersionsForModule = Get-BlockedVersionsForModule -ModuleName $Module -BlockedVersionMap $ModuleVersionBlockList
    $MostRecentAllowedEntry = $AllVersions | Where-Object { $BlockedVersionsForModule -notcontains $_.Version.ToString() } | Select-Object -First 1
    if ($MostRecentAllowedEntry) {
      $MostRecentVersion = $MostRecentAllowedEntry.Version
      Write-Host "Most recent allowed version of $Module is $MostRecentVersion"
    }
    else {
      Write-Host ("No allowed installed version found for {0}. Only blocked versions are installed or no usable version exists." -f $Module) -ForegroundColor Yellow
      $MostRecentVersion = $null
    }

    # Blocked versions are always removed, even when only one version is installed.
    # Non-blocked versions are reduced to the newest allowed version.
    $RemovedAnyVersion = $false
    foreach ($Version in $AllVersions) {
      $IsBlockedVersion = $BlockedVersionsForModule -contains $Version.Version.ToString()
      $IsOlderThanAllowed = $MostRecentVersion -and ($Version.Version -lt $MostRecentVersion)

      if ($IsBlockedVersion -or $IsOlderThanAllowed) {
        Write-Host ("Uninstalling version {0} of module {1}" -f $Version.Version, $Module) -ForegroundColor Red
        if (Remove-InstalledModuleFolder -ModuleName $Module -ModuleVersion $Version.Version -ModuleBasePath $Version.ModuleBase -AllowedRootPath $psmodulepath) {
          $RemovedModules++
          $RemovedAnyVersion = $true
        }
      }
    }

    if (-not $RemovedAnyVersion) {
      Write-Host ("No earlier or blocked versions of {0} module to remove" -f $Module)
    }
  } # End foreach module

  Write-Host ("Installed modules: {0} ; Updated modules: {1} ; Removed old versions of modules: {2}" -f $InstalledModules, $UpdatedModules, $RemovedModules)
}
else {
  Write-Host "Skipping PowerShell module update (selection)." -ForegroundColor Yellow
}

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the need of your organization. Never run any code downloaded from the Internet without
# first validating the code in a non-production environment.

#endregion powershell modules


#region vscode
########################################################################################################################################################################################

$VSCodeExe = Join-Path -Path $VSCodeInstallPath -ChildPath "Code.exe"
$VSCodeCli = Join-Path -Path $VSCodeInstallPath -ChildPath "bin\code.cmd"

if (-not $RunVSCodeUpdate) {
  Write-Host "Skipping VS Code update (selection)." -ForegroundColor Yellow
}
elseif (-not (Test-Path -LiteralPath $VSCodeExe)) {
  Write-Host "VS Code is not installed in $VSCodeInstallPath. Skipping VS Code update." -ForegroundColor Yellow
}
else {
  Write-Host ""
  Write-Host ""
  Write-Host "########## Check whether a newer Visual Studio Code version is available" -ForegroundColor Cyan

  $GitHubversion = Get-GitHubVersion -url "https://github.com/microsoft/vscode/releases/latest"
  $localversion = (Get-Item -LiteralPath $VSCodeExe).VersionInfo.ProductVersion

  [version]$VSCodeRemoteVersion = $null
  [version]$VSCodeLocalVersion = $null
  if (-not [version]::TryParse([string]$GitHubversion, [ref]$VSCodeRemoteVersion)) {
    Write-Host "Failed to retrieve a valid VS Code version from GitHub ('$GitHubversion'). Skipping update." -ForegroundColor Red
  }
  elseif (-not [version]::TryParse([string]$localversion, [ref]$VSCodeLocalVersion)) {
    Write-Host "Could not parse the local VS Code version ('$localversion'). Skipping update." -ForegroundColor Red
  }
  elseif ($VSCodeRemoteVersion -le $VSCodeLocalVersion) {
    Write-Host "Latest VS Code version is installed ($localversion) - no update required." -ForegroundColor Green
  }
  else {
    Write-Host "Newer version found on GitHub: $GitHubversion"
    Write-Host "Local version: $localversion"
    Write-Host "Updating VS Code..."
    $Destination = Join-Path -Path $temppath -ChildPath "vscode_installer.exe"
    $VSCodeUrl = "https://code.visualstudio.com/sha/download?build=stable&os=win32-x64"

    if (-not (Test-Path -Path $temppath)) {
      New-Item -Path $temppath -ItemType Directory -Force | Out-Null
    }

    # Download first, so VS Code is only closed when the installer is actually available.
    Write-Host "Downloading VS Code..."
    if (-not (Save-FileDownload -Url $VSCodeUrl -Destination $Destination)) {
      Write-Host "VS Code download failed after all retries. Skipping installation." -ForegroundColor Red
    }
    else {
      Write-Host "Closing any running instances of VS Code..."
      $VSCodeProcesses = Get-Process -Name "Code" -ErrorAction SilentlyContinue
      $VSCodeProcesses | ForEach-Object { [void]$_.CloseMainWindow() }
      $VSCodeProcesses | Wait-Process -Timeout 10 -ErrorAction SilentlyContinue
      Get-Process -Name "Code" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

      Write-Host "Installing VS Code..."
      try {
        $VSCodeProcess = Start-Process -FilePath $Destination -ArgumentList '/VERYSILENT', '/NORESTART', '/MERGETASKS=!runcode' -Wait -PassThru -ErrorAction Stop
        if ($VSCodeProcess.ExitCode -eq 0) {
          Write-Host "VS Code installation completed successfully." -ForegroundColor Green
        }
        else {
          Write-Host "VS Code installer exited with code $($VSCodeProcess.ExitCode)." -ForegroundColor Red
        }
      }
      catch {
        Write-Host "Failed to run the VS Code installer. Error: $PSItem" -ForegroundColor Red
      }

      Write-Host "Removing installation file..."
      Remove-Item -Path $Destination -Force -ErrorAction SilentlyContinue
    }
  }

  #endregion vscode

  #region vscode extensions
  ##############################################################################################################################################################################

  # Extensions are per user - this updates the extensions of the account running this script.
  # The CLI is called by its full path because "code" may not be on PATH in this session yet.
  if (Test-Path -LiteralPath $VSCodeCli) {
    Write-Host "Current VS Code extension versions:"
    & $VSCodeCli --list-extensions --show-versions

    Write-Host "Starting VS Code extension update..."
    Write-Host "The DeprecationWarning for VS Code extensions update may appear but can be safely ignored. If no update is available, you will see 'No extension to update'."
    & $VSCodeCli --update-extensions
  }
  else {
    Write-Host "VS Code CLI not found at $VSCodeCli. Skipping extension update." -ForegroundColor Yellow
  }

  #endregion vscode extensions
}

#region git
##############################################################################################################################################################################

if ($RunGitUpdate) {
  Write-Host  ""
  Write-Host  ""
  Write-Host "########## Check whether a newer Git for Windows version is available" -ForegroundColor Cyan

  $GitExe = (Get-Command -Name "git.exe" -ErrorAction SilentlyContinue | Select-Object -First 1).Source
  if (-not $GitExe -and (Test-Path -LiteralPath "$env:ProgramFiles\Git\cmd\git.exe")) {
    $GitExe = "$env:ProgramFiles\Git\cmd\git.exe"
  }

  if (-not $GitExe) {
    Write-Host "Git is not installed on this system. Skipping Git update." -ForegroundColor Yellow
  }
  else {
    # Release tags look like v2.47.1.windows.1 ; local output like "git version 2.47.1.windows.1"
    $GitLocalRaw = ((& $GitExe --version) -replace '^git version\s+', '').Trim()
    $GitHubversion = Get-GitHubVersion -url "https://github.com/git-for-windows/git/releases/latest"

    $GitVersionPattern = '^(?<base>\d+\.\d+\.\d+)\.windows\.(?<rev>\d+)$'
    if ([string]::IsNullOrWhiteSpace($GitHubversion) -or $GitHubversion -notmatch $GitVersionPattern) {
      Write-Host "Failed to retrieve a valid Git version from GitHub ('$GitHubversion'). Skipping update." -ForegroundColor Red
    }
    elseif ($GitLocalRaw -notmatch $GitVersionPattern) {
      Write-Host "Could not parse the local Git version ('$GitLocalRaw'). Skipping update." -ForegroundColor Red
    }
    else {
      $GitLocalMatch = [regex]::Match($GitLocalRaw, $GitVersionPattern)
      $GitRemoteMatch = [regex]::Match($GitHubversion, $GitVersionPattern)
      $GitHubversionBase = $GitRemoteMatch.Groups['base'].Value
      $GitHubversionRev = [int]$GitRemoteMatch.Groups['rev'].Value
      [version]$GitRemoteComparable = "$GitHubversionBase.$GitHubversionRev"
      [version]$GitLocalComparable = "$($GitLocalMatch.Groups['base'].Value).$([int]$GitLocalMatch.Groups['rev'].Value)"

      if ($GitRemoteComparable -gt $GitLocalComparable) {
        Write-Host "Newer version found on GitHub: $GitHubversion"
        Write-Host "Local version: $GitLocalRaw"
        Write-Host "Updating Git for Windows..."

        # Git for Windows does not publish an MSI; the official installer is an Inno Setup EXE.
        # Revision 1 -> Git-2.47.1-64-bit.exe ; revision n>1 -> Git-2.47.1.n-64-bit.exe
        $GitFileVersion = if ($GitHubversionRev -gt 1) { "$GitHubversionBase.$GitHubversionRev" } else { $GitHubversionBase }
        $GitInstallerFileName = "Git-$GitFileVersion-64-bit.exe"
        $GitInstallerPath = Join-Path -Path $temppath -ChildPath $GitInstallerFileName
        $GitInstallLogPath = Join-Path -Path $Logpath -ChildPath "$(Get-Date -format "yyyyMMdd_HHmmss")_Git-$GitFileVersion-install.log"
        $GitInstallerUrl = "https://github.com/git-for-windows/git/releases/download/v$GitHubversion/$GitInstallerFileName"

        Write-Host "Download URL: $GitInstallerUrl"
        Write-Host "Git install log: $GitInstallLogPath"

        if (-not (Test-Path -Path $temppath)) {
          New-Item -Path $temppath -ItemType Directory -Force | Out-Null
        }

        Write-Host "Downloading Git installer..."
        if (-not (Save-FileDownload -Url $GitInstallerUrl -Destination $GitInstallerPath)) {
          Write-Host "Git download failed after all retries. Skipping installation." -ForegroundColor Red
        }
        else {
          Write-Host "Installing Git silently..."
          try {
            $GitArgs = @('/VERYSILENT', '/NORESTART', '/SUPPRESSMSGBOXES', '/NOCANCEL', "/LOG=`"$GitInstallLogPath`"")
            $GitProcess = Start-Process -FilePath $GitInstallerPath -ArgumentList $GitArgs -Wait -PassThru -ErrorAction Stop
            if ($GitProcess.ExitCode -eq 0) {
              Write-Host "Git installation completed successfully." -ForegroundColor Green
            }
            else {
              Write-Host "Git installer exited with code $($GitProcess.ExitCode). See $GitInstallLogPath" -ForegroundColor Red
            }
          }
          catch {
            Write-Host "Failed to run the Git installer. Error: $PSItem" -ForegroundColor Red
          }

          Write-Host "Removing installation file..."
          Remove-Item -Path $GitInstallerPath -Force -ErrorAction SilentlyContinue
        }
      }
      else {
        Write-Host "Latest Git version is installed ($GitLocalRaw) - no update required." -ForegroundColor Green
      }
    }
  }
}
else {
  Write-Host "Skipping Git update (selection)." -ForegroundColor Yellow
}

#endregion git

#region powershell
##############################################################################################################################################################################

if (-not $RunPowerShellUpdate) {
  Write-Host "Skipping PowerShell update (selection)." -ForegroundColor Yellow
}
else {
  Write-Host ""
  Write-Host ""
  Write-Host "############ Check whether a newer PowerShell version is available..." -ForegroundColor Cyan

  # Prefer winget when available.
  $WingetCommand = Get-Command -Name "winget.exe" -ErrorAction SilentlyContinue
  if ($WingetCommand) {
    Write-Host "winget is available. Checking for a PowerShell upgrade..." -ForegroundColor Cyan
    # --exact avoids matching Microsoft.PowerShell.Preview; the agreement switches prevent interactive prompts.
    $WingetCommonArgs = @('--id', 'Microsoft.PowerShell', '--exact', '--source', 'winget', '--accept-source-agreements', '--disable-interactivity')

    $WingetListOutput = $null
    try {
      Write-Host "Running: winget list $($WingetCommonArgs -join ' ') --upgrade-available"
      $WingetListOutput = & $WingetCommand.Source list @WingetCommonArgs --upgrade-available 2>&1 | Out-String
      $WingetListExitCode = $LASTEXITCODE
      if (-not [string]::IsNullOrWhiteSpace($WingetListOutput)) {
        Write-Host $WingetListOutput.Trim()
      }
    }
    catch {
      Write-Host "winget check failed. Error: $PSItem" -ForegroundColor Red
      $WingetListExitCode = -1
    }

    if ($WingetListExitCode -eq 0 -and $WingetListOutput -match "Microsoft\.PowerShell") {
      Write-Host "PowerShell upgrade is available. Starting winget upgrade..." -ForegroundColor Yellow
      $WingetUpgradeArgs = $WingetCommonArgs + @('--silent', '--accept-package-agreements')
      Write-Host "Running: winget upgrade $($WingetUpgradeArgs -join ' ')"
      & $WingetCommand.Source upgrade @WingetUpgradeArgs
      if ($LASTEXITCODE -eq 0) {
        Write-Host "PowerShell upgrade via winget completed. Restart the PowerShell session to use the new version." -ForegroundColor Green
      }
      else {
        Write-Host "PowerShell upgrade via winget failed with exit code $LASTEXITCODE." -ForegroundColor Red
      }
    }
    else {
      Write-Host "No newer PowerShell version found via winget." -ForegroundColor Green
    }
  }
  else {
    Write-Host "winget is not available. Falling back to the ZIP based PowerShell update." -ForegroundColor Yellow

    $GitHubversion = Get-GitHubVersion -url "https://github.com/PowerShell/PowerShell/releases/latest"
    $PwshLocalVersion = [version]::new($PSVersionTable.PSVersion.Major, $PSVersionTable.PSVersion.Minor, $PSVersionTable.PSVersion.Patch)
    [version]$PwshRemoteVersion = $null

    if (-not [version]::TryParse([string]$GitHubversion, [ref]$PwshRemoteVersion)) {
      Write-Host "Failed to retrieve a valid PowerShell version from GitHub ('$GitHubversion'). Skipping update." -ForegroundColor Red
    }
    elseif ($PwshRemoteVersion -le $PwshLocalVersion) {
      Write-Host "Latest PowerShell version is installed ($PwshLocalVersion) - no update required." -ForegroundColor Green
    }
    else {
      Write-Host "Newer version found on GitHub: $GitHubversion"
      Write-Host "Local version: $PwshLocalVersion"

      $PwshInstallPath = Join-Path -Path $env:ProgramFiles -ChildPath "PowerShell\7"
      $installerFileName = "PowerShell-$GitHubversion-win-x64.zip"
      $installerPath = Join-Path -Path $temppath -ChildPath $installerFileName
      $installLogPath = Join-Path -Path $Logpath -ChildPath "$(Get-Date -Format "yyyyMMdd_HHmmss")_PowerShell-$GitHubversion-zip-install.log"
      $installerUrl = "https://github.com/PowerShell/PowerShell/releases/download/v$GitHubversion/$installerFileName"

      Write-Host "Download URL: $installerUrl"
      Write-Host "Target folder: $PwshInstallPath"
      Write-Host "ZIP install log: $installLogPath"

      if (-not (Test-Path -Path $temppath)) {
        New-Item -Path $temppath -ItemType Directory -Force | Out-Null
      }

      Write-Host "Downloading PowerShell ZIP..."
      if (-not (Save-FileDownload -Url $installerUrl -Destination $installerPath)) {
        Write-Host "PowerShell ZIP download failed after all retries. Skipping installation." -ForegroundColor Red
      }
      else {
        try {
          Write-Host "Unblocking downloaded ZIP file..."
          Unblock-File -Path $installerPath -ErrorAction Stop

          Write-Host "Starting elevated extraction to $PwshInstallPath. It runs after this session has ended (pwsh.exe files are locked while in use). Check the install log afterwards."
          Invoke-PowerShellZipInstall -ZipPath $installerPath -DestinationPath $PwshInstallPath -InstallLogPath $installLogPath
          Write-Host "PowerShell ZIP extraction was started in the background. Close this session to let it finish." -ForegroundColor Green
        }
        catch {
          Write-Host "Failed to start the PowerShell ZIP extraction. Error: $PSItem" -ForegroundColor Red
        }
      }
    }
  }
}

#endregion powershell

if ($TranscriptActive) {
  Stop-Transcript
}
