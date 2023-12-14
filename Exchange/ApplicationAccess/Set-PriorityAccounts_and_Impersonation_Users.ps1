Connect-ExchangeOnline

# define variables
$phishingpolicyname = "Office365 AntiPhish Default"
$exportfilepath = "c:\temp"
$executivesgroup = "Executives@domain.com"
#set to false when you want to overwrite the impersonation users everytime you run the script
#if set to true the script will add the current list of VIP flagged users to the impersonation users list
#this way you continue to protect your users from impersonation of VIP users after they eventually left the company
$addimpersonationusers = $true

# function for setting VIP Tag to users
function set-viptag($userlist)
{
    foreach($user in $userlist)
    {
        set-user $user.alias -VIP $true -Confirm:$false
        $alias = $user.alias
        $displayname = $user.displayname
        write-host "Set VIP tag to user $displayname with alias $alias"
    }
}

# function for getting the current VIP user list and export the current list
function get-vipuserlist
{
    Param([switch]$export)
    $date = get-date -Format FileDateTime
    $vipusers = Get-User -IsVIP
    if($export -eq $true) { $vipusers | ConvertTo-Json | out-file "$exportfilepath\$date-vip-users.json" }
    return $vipusers
}

# get vip list, output and export
get-vipuserlist -export | Sort-object DisplayName | `
    Select-Object DisplayName,WindowsEmailAddress,UserPrincipalName | `
    Format-Table -AutoSize

# remove VIP tag from current list
foreach($vipuser in $vipusers)
{
    $vipuser | set-user -VIP $false -Confirm:$false
    $alias = $vipuser.alias
    $displayname = $vipuser.displayname
    Write-Host "removed VIP tag from user $displayname with alias $alias"
}

# get new VIP user list from group
$test = Get-DistributionGroup $executivesgroup | `
    foreach-object { Get-DistributionGroupMember $_.name } 

# set vip tags to users
set-viptag -userlist $test

# export the current list of vip users
get-vipuserlist -export

# work on user impersonation
# export current user impersonation list
$date = get-date -Format FileDateTime
$TargetedUsersToProtect_Current = (Get-AntiPhishPolicy $phishingpolicyname).TargetedUsersToProtect
$TargetedUsersToProtect_Current | Out-File "$exportfilepath\$date-impersonation-users.txt"
$TargetedUsersToProtect_Current

# get vip users
$vipusers = get-vipuserlist
# construct string from vip user list
$TargetedUsersToProtect_New = $null
$TargetedUsersToProtect_New = foreach ($vipuser in $vipusers) { $vipuser.DisplayName, $vipuser.UserPrincipalName -join ";" }
$TargetedUsersToProtect_New

if($addimpersonationusers -eq $true)
{
    #add the new users to the protected users list
    $TargetedUsersToProtect_combined = $TargetedUsersToProtect_Current + $TargetedUsersToProtect_New
    #sort out duplicates
    $TargetedUsersToProtect = $TargetedUsersToProtect_combined | sort-object -Unique
    #set users to protect
    Set-AntiPhishPolicy -Identity $phishingpolicyname -TargetedUsersToProtect $TargetedUsersToProtect
    Write-Host "Adding the new users to protect from user impersonation to the curent list:"
    $TargetedUsersToProtect
}
else #overwrite the old list
{
    #set users to protect
    Set-AntiPhishPolicy -Identity $phishingpolicyname -TargetedUsersToProtect $TargetedUsersToProtect_New
    Write-Host "overwriting the curent list of users to protect from user impersonation with the new one:"
    $TargetedUsersToProtect_New
}