# Update PowerShell 7, common PowerShell modules, and VS Code
# region Variables
param(
  [switch]$PostPowerShellUpdateCheck,
  [string]$ExpectedPowerShellVersion
)

$Logpath = "C:\Temp"
$Scriptname = "Update_VSCode_PS_Modules"
$temppath = "C:\Temp" #temp path for downloading the setup files
$psmodulepath = "C:\Program Files\PowerShell\Modules"
$tempPsmodulepath = "C:\Temp\InstallPSModules"

# Define the set of modules installed and updated from the PowerShell Gallery that we want to maintain - edit this set of modules to include the modules 
# you want to process.
#$O365Modules = @("PackageManagement", "Az.Accounts", "Az.Automation", "AIPService", "Az.Keyvault", "MicrosoftTeams", "Microsoft.Graph", "Microsoft.Graph.Beta", "ExchangeOnlineManagement", "Microsoft.Online.Sharepoint.PowerShell", "ORCA",  "Pnp.PowerShell", "MSCommerce", "Microsoft365DSC", "MSAL.PS", "WhiteboardAdmin", "ImportExcel")
$O365Modules = @("PackageManagement", "Microsoft.Graph", "Microsoft.Graph.Beta", "ExchangeOnlineManagement", "Pnp.PowerShell", "MSAL.PS", "ImportExcel")

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
  try {
    Write-Host "Retrieving version from GitHub..."
    $request = [System.Net.WebRequest]::Create($url)
    #$request.Proxy = New-Object System.Net.WebProxy($proxyserver, $proxyport)
    $response = $request.GetResponse()
    $realTagUrl = $response.ResponseUri.OriginalString
    $GitHubversion = $realTagUrl.split('/')[-1].Trim('v')
    Write-Host "Successfully retrieved version from GitHub: $GitHubversion"
    return $GitHubversion
  }
  catch {
    Write-Host "$(get-date -Format "HH:mm:ss") Failed to retrieve GitHub version! Error: $psitem" -ForegroundColor Red
    return "ERROR"
  }
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

  [array]$galleryVersions = Find-Module -Name $ModuleName -Repository PSGallery -AllVersions -ErrorAction Stop

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

function Remove-InstalledModuleFolder {
  param(
    [Parameter(Mandatory = $true)]
    [string]$ModuleName,
    [Parameter(Mandatory = $true)]
    [version]$ModuleVersion,
    [Parameter(Mandatory = $true)]
    [string]$ModuleBasePath
  )

  # Module versions are removed by deleting the corresponding version folder.
  # This is intentional because modules are deployed via Save-Module.
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
#endregion functions

#region prerequisites
####################################################################################################################################

Set-location $Logpath
Start-Transcript -Path "$Logpath\$(Get-Date -format "yyyyMMdd_HHmmss")_$($Scriptname.Split(".")[0])_TS.LOG"

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
    Write-Host "Post-update check is enabled, but no target version was provided." -ForegroundColor Red
    Stop-Transcript
    exit 1
  }

  [version]$installedVersion = $currentPwshVersion
  [version]$expectedVersion = $ExpectedPowerShellVersion

  if ($installedVersion -lt $expectedVersion) {
    Write-Host "PowerShell rerun check failed. Installed: $installedVersion ; Expected: $expectedVersion" -ForegroundColor Red
    Stop-Transcript
    exit 1
  }

  Write-Host "PowerShell rerun check succeeded. Installed: $installedVersion ; Expected: $expectedVersion" -ForegroundColor Green
  Stop-Transcript
  exit 0
}
#endregion prerequisites

#region powershell modules
########################################################################################################################################################################################

# UpdateOffice365PowerShellModules.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/UpdateOffice365PowerShellModules.PS1
# Very simple script to check for updates to a defined set of PowerShell modules used to manage Office 365 services
# If an update for a module is found, it is downloaded and applied.
# Once all modules are checked for updates, we remove older versions that might be present on the workstation. V2.1 improves the processing of Microsoft Graph SDK sub-modules

[int]$InstalledModules = 0; [int]$UpdatedModules = 0; [int]$RemovedModules = 0

Write-Host ("Starting up and preparing to process these modules: {0}" -f ($O365Modules -join ", ")) -foregroundcolor Yellow

# We're installing from the PowerShell Gallery so make sure that it's trusted
Set-PSRepository -Name PsGallery -InstallationPolicy Trusted

Write-Host  ""
Write-Host  ""
Write-Host "######### Check and update all modules to make sure that we're at the latest version" -ForegroundColor Cyan
ForEach ($Module in $O365Modules) 
{
  Write-Host ("Checking module {0}" -f $Module)
   $BlockedVersionsForModule = Get-BlockedVersionsForModule -ModuleName $Module -BlockedVersionMap $ModuleVersionBlockList
   if ($BlockedVersionsForModule.Count -gt 0) {
     Write-Host ("Blocked versions for {0}: {1}" -f $Module, ($BlockedVersionsForModule -join ", ")) -ForegroundColor Yellow
   }

   try {
     $CurrentVersion = Get-PreferredGalleryVersion -ModuleName $Module -BlockedVersions $BlockedVersionsForModule
   }
   catch {
     Write-Host "Could not read module versions from PSGallery for $Module - Error: $psitem" -ForegroundColor Red
     continue
   }

   if ([string]::IsNullOrWhiteSpace($CurrentVersion)) {
     Write-Host "No installable version found for module $Module (all versions are blocked)." -ForegroundColor Red
     continue
   }

   Write-Host ("Preferred (allowed) PSGallery version of {0} is {1}" -f $Module, $CurrentVersion.ToString())

   # Check if the module is installed on this PC
   Clear-Variable PCModule -WarningAction SilentlyContinue -ErrorAction SilentlyContinue # make sure we start fresh
   [array]$PCModule = Get-Module -Name $Module -ListAvailable -ErrorAction SilentlyContinue 

  # Remove blocked versions immediately if they are already installed
   foreach ($InstalledEntry in $PCModule) {
     $InstalledVersionString = $InstalledEntry.Version.ToString()
     if ($BlockedVersionsForModule -contains $InstalledVersionString) {
       Write-Host ("Blocked version detected for {0}: {1}. Removing {2}" -f $Module, $InstalledVersionString, $InstalledEntry.ModuleBase) -ForegroundColor Yellow
       if (Remove-InstalledModuleFolder -ModuleName $Module -ModuleVersion $InstalledEntry.Version -ModuleBasePath $InstalledEntry.ModuleBase) {
         $RemovedModules++
       }
     }
   }

   [array]$PCModule = Get-Module -Name $Module -ListAvailable -ErrorAction SilentlyContinue
   $PCModule = $PCModule | Sort-Object Version -Descending

   [version]$HighestInstalledVersion = $null
   if ($PCModule) {
     $HighestInstalledVersion = $PCModule[0].Version
   }

   If (!($PCModule)) { 
   # No version of the module found. It's in our list, so we install it.
     Write-Host ("No module found on this PC for {0}" -f $Module)
      Write-Host ("Installing module {0} version {1}..." -f $Module, $CurrentVersion.ToString())  -foregroundcolor Yellow
     #Install-Module $Module -Scope AllUsers -Confirm:$False -AllowClobber -Force -Repository PSGallery
     try {
      Save-ModuleVersionToModulePath -ModuleName $Module -RequiredVersion $CurrentVersion -TempPath $tempPsmodulepath -TargetPath $psmodulepath
      Write-Host "Module $Module installed in the common module folder." -ForegroundColor Green
      $InstalledModules++
     }
     catch {
      write-host "Module installation failed for $Module - Error: $psitem" -ForegroundColor Red
      }
   }

   If ($PCModule -and $HighestInstalledVersion -eq $CurrentVersion)
   {
      Write-Host ("Latest allowed version of $Module is installed on this PC - no need to update" )
      Write-Host "Module Path = $($PCModule[0].Path)"
   }
    elseif ($PCModule -and $HighestInstalledVersion -gt $CurrentVersion)
    {
      Write-Host ("Installed version {0} of {1} is newer than preferred PSGallery version {2}. Skipping downgrade." -f $HighestInstalledVersion, $Module, $CurrentVersion) -ForegroundColor Yellow
    }
   elseif ($PCModule)
   {
      Write-Host ("Installing latest allowed version of module. Installed Version: {0} ; Allowed PSGallery Version: {1}" -f $HighestInstalledVersion, $CurrentVersion) -ForegroundColor Green
      try
      {
        Save-ModuleVersionToModulePath -ModuleName $Module -RequiredVersion $CurrentVersion -TempPath $tempPsmodulepath -TargetPath $psmodulepath
        $UpdatedModules++
        Write-Host "New allowed module version of $Module has been deployed to the Modules folder."
      }
      catch
      {
        write-host "Module update failed for $Module - Error: $psitem" -ForegroundColor Red
      }
   }
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
     $BlockedVersionsForModule = Get-BlockedVersionsForModule -ModuleName $Module -BlockedVersionMap $ModuleVersionBlockList
     $AllVersions = $AllVersions | Sort-Object Version -Descending 
     $MostRecentAllowedEntry = $AllVersions | Where-Object { $BlockedVersionsForModule -notcontains $_.Version.ToString() } | Select-Object -First 1
     if ($MostRecentAllowedEntry) {
       $MostRecentVersion = $MostRecentAllowedEntry.Version
       Write-Host ("Most recent allowed version of $Module is $MostRecentVersion" )
     }
     else {
       Write-Host ("No allowed installed version found for {0}. Only blocked versions are installed or no usable version exists." -f $Module) -ForegroundColor Yellow
       $MostRecentVersion = $null
     }

    # Blocked versions are always removed, even when only one version is installed.
    # Non-blocked versions are reduced to the newest allowed version.
     $RemovedAnyVersion = $false
     ForEach ($Version in $AllVersions) 
     {
       $VersionString = $Version.Version.ToString()
       $IsBlockedVersion = $BlockedVersionsForModule -contains $VersionString
       $IsOlderThanAllowed = $MostRecentVersion -and ($Version.Version -lt $MostRecentVersion)

       If ($IsBlockedVersion -or $IsOlderThanAllowed)
       {
          Write-Host ("Uninstalling version {0} of module {1}" -f $Version.Version, $Module) -foregroundcolor Red
          if (Remove-InstalledModuleFolder -ModuleName $Module -ModuleVersion $Version.Version -ModuleBasePath $Version.ModuleBase) {
            $RemovedModules++
            $RemovedAnyVersion = $true
          }
       }
     }

     if (-not $RemovedAnyVersion) {
       Write-Host ("No earlier or blocked versions of {0} module to remove" -f $Module)
     }
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
Write-Host "########## Check whether a newer Visual Studio Code version is available" -ForegroundColor Cyan
$GitHubversion = $null
$localversion = $null

# Retrieve version (errors are handled inside the function)
$GitHubversion = Get-GitHubVersion -url "https://github.com/microsoft/vscode/releases/latest"

$localversion = (Get-ChildItem "C:\Program Files\Microsoft VS Code\Code.exe").VersionInfo.ProductVersion
if($GitHubversion -eq "ERROR" -or [string]::IsNullOrWhiteSpace($GitHubversion)){
  Write-Host "Failed to retrieve the current version from GitHub. Skipping update."
}
elseif($GitHubversion -ne $localversion) 
{
  Write-Host "Newer version found on GitHub: $GitHubversion"
  Write-Host "Local version: $localversion"
  Write-Host "Updating VS Code..."
  # Define the download URL and the destination
  $Destination = "$temppath\vscode_installer.exe"
  $VSCodeUrl = "https://code.visualstudio.com/sha/download?build=stable&os=win32-x64"

  # Close any running instances of VS Code
  Write-Host "Closing any running instances of VS Code..."
  Get-Process -Name "Code" -ErrorAction SilentlyContinue | ForEach-Object { $_.CloseMainWindow(); Stop-Process -Id $_.Id -Force }

  # Download VSCode installer
  Write-Host "Downloading VS Code..."
  Invoke-WebRequest -Uri $VSCodeUrl -OutFile $Destination -SkipCertificateCheck

  # Install VS Code silently
  Write-Host "Installing VS Code..."
  try 
  {
    Start-Process -FilePath $Destination -ArgumentList '/verysilent /mergetasks=!runcode' -Wait -Passthru
    Write-Host "VS Code installer completed successfully."
  }
  catch 
  {
    Write-Host "Failed to run the VS Code installer."
  }

  # Remove installer
  Write-Host "Removing installation file..."
  Remove-Item $Destination

  Write-Host "VS Code installation completed."
}
else 
{
  Write-Host "Latest VS Code version is installed ($GitHubversion) - no update required." -ForegroundColor Green
}

#endregion vscode

#region vscode extensions
##############################################################################################################################################################################

#endregion vscode extensions
Write-Host "Current VS Code extension versions:"
code --list-extensions --show-versions

Write-Host "Starting VS Code extension update..."
code --update-extensions

#region powershell
##############################################################################################################################################################################

Write-Host  ""
Write-Host  ""
Write-Host "############ Check whether a newer PowerShell version is available..."
# Prefer winget when available.
$WingetCommand = Get-Command -Name "winget.exe" -ErrorAction SilentlyContinue
if ($WingetCommand) {
  Write-Host "winget is available. Checking for a PowerShell upgrade..." -ForegroundColor Cyan

  try {
    Write-Host "Running: winget list --id Microsoft.PowerShell --upgrade-available"
    $WingetListOutput = winget list --id Microsoft.PowerShell --upgrade-available 2>&1 | Out-String
    if (-not [string]::IsNullOrWhiteSpace($WingetListOutput)) {
      Write-Host $WingetListOutput.Trim()
    }
  }
  catch {
    Write-Host "winget check failed. Error: $PSItem" -ForegroundColor Red
    $WingetListOutput = $null
  }

  if ($WingetListOutput -match "Microsoft\.PowerShell") {
    Write-Host "PowerShell upgrade is available. Starting winget upgrade..." -ForegroundColor Yellow
    try {
      Write-Host "Running: winget upgrade --id Microsoft.PowerShell"
      winget upgrade --id Microsoft.PowerShell
      if ($LASTEXITCODE -ne 0) {
        throw "winget upgrade exited with code $LASTEXITCODE"
      }
      Write-Host "PowerShell upgrade via winget completed." -ForegroundColor Green
    }
    catch {
      Write-Host "PowerShell upgrade via winget failed. Error: $PSItem" -ForegroundColor Red
    }
  }
  else {
    Write-Host "No newer PowerShell version found via winget." -ForegroundColor Green
  }
}
else {
  Write-Host "winget is not available. Falling back to the current PowerShell update procedure." -ForegroundColor Yellow

  # Check if a newer PowerShell version is available
  $GitHubversion = $null
  $localversion = $null

  # Retrieve version (errors are handled inside the function)
  $GitHubversion = Get-GitHubVersion -url "https://github.com/PowerShell/PowerShell/releases/latest"

  $localversion = $PSVersionTable.PSVersion.Major.ToString() +"."+ $PSVersionTable.PSVersion.Minor.ToString() +"."+ $PSVersionTable.PSVersion.Patch.ToString()
  if($GitHubversion -eq "ERROR" -or [string]::IsNullOrWhiteSpace($GitHubversion)){
    Write-Host "Failed to retrieve the current version from GitHub. Skipping update."
  }
  elseif($GitHubversion -ne $localversion) 
  {
    Write-Host "Newer version found on GitHub: $GitHubversion"
    Write-Host "Local version: $localversion"
    Write-Host "Updating PowerShell in a separate process..."

    # The update is run in a separate Windows PowerShell process
    # so the current pwsh session does not crash during MSI update.
    $PowerShellUpdateLogPath = Join-Path -Path $Logpath -ChildPath "$(Get-Date -format "yyyyMMdd_HHmmss")_PowerShellUpdate.log"
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

  # Short delay to ensure files/registry are stable after MSI setup.
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
    "-File", $MainScriptPath,
    "-PostPowerShellUpdateCheck",
    "-ExpectedPowerShellVersion", $TargetVersion
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
      "-File", $PowerShellUpdaterScriptPath,
      "-ParentProcessId", $PID,
      "-TargetVersion", $GitHubversion,
      "-LogPath", $PowerShellUpdateLogPath,
      "-MainScriptPath", $PSCommandPath
    )

    Start-Process -FilePath "powershell.exe" -ArgumentList $PowerShellUpdaterArgs -WindowStyle Hidden | Out-Null

    Write-Host "PowerShell update has been started in the background."
    Write-Host "Validation log: $PowerShellUpdateLogPath"
    Write-Host "This script will now exit so the current pwsh instance does not block the update." -ForegroundColor Yellow
    Stop-Transcript
    exit 0
  }
  else
  {
    Write-Host "Latest PowerShell version is installed ($GitHubversion) - no update required." -ForegroundColor Green
  }
}

#endregion powershell

Stop-Transcript