set-location "C:\Scripts\GitHub\M365\Exchange\EWS Usage Report"

$clientid = ""
$tenantid = ""
$clientsecret = ConvertTo-SecureString "" -AsPlainText -Force

# 1 GetAppUsage - checks either Sign-in logs or Audit logs for EWS application usage
#make new folder for output if it doesn't exist
$outputfolder = "C:\temp\$((Get-Date).ToString('yyyyMMdd'))_EWSUsageReport"
if (!(Test-Path -Path $outputfolder)) {
    New-Item -ItemType Directory -Path $outputfolder
}

& "C:\Scripts\GitHub\Exchange-App-Usage-Reporting\Find-EwsUsage.ps1" `
    -StartDate (Get-Date).AddDays(-7) `
    -EndDate (Get-Date) `
    -OutputPath $outputfolder `
    -PermissionType Delegated `
    -Operation GetAppUsage

# 2 GetEwsActivity - checks for ServicePrincipalName activity for applications with EWS permissions
& "C:\Scripts\GitHub\Exchange-App-Usage-Reporting\Find-EwsUsage.ps1" `
    -StartDate (Get-Date).AddDays(-7) `
    -EndDate (Get-Date) `
    -OutputPath $outputfolder `
    -PermissionType Delegated `
    -Operation GetEwsActivity


# 2) Get active applications
# Open a PowerShell session and change to the folder where you downloaded the script. 
# You may need to unblock the files (for example, by using Unblock-File) before execution. 
# The output provides a list of applications with EWS permissions and the last sign-in for the
# associated service principal. A CSV file called App-SignInActivity-yyyyMMddhhmm will be created 
# in the specified output path.
& "C:\Scripts\GitHub\Exchange-App-Usage-Reporting\Find-EwsUsage.ps1" `
    -OutputPath $outputfolder `
    -OAuthClientId $clientid `
    -OAuthTenantId $tenantid `
    -OAuthClientSecret $clientsecret `
    -PermissionType Application `
    -Operation GetEwsActivity


# 3) Get sign-in activity report for a specific application
# Use the AppId/Name from step 2. Run this step once per application.
# Depending on tenant size, reduce the date range and set Interval to 1 hour.
$reportAppId = "9b41a7fb-f046-4807-8b5b-6c6324ea39d6"
$reportAppName = "EWS Test App"

& "C:\Scripts\GitHub\Exchange-App-Usage-Reporting\Find-EwsUsage.ps1" `
    -OutputPath $outputfolder `
    -OAuthClientId $clientid `
    -OAuthTenantId $tenantid `
    -OAuthClientSecret $clientsecret `
    -PermissionType Application `
    -Operation GetAppUsage `
    -QueryType SignInLogs `
    -Name $reportAppName `
    -AppId $reportAppId `
    -StartDate (Get-Date).AddDays(-1) `
    -EndDate (Get-Date) `
    -Interval 4

# Output: <AppId>-SignInEvents-yyyyMMddhhmm.csv in $outputfolder
