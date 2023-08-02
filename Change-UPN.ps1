# export all AD users
$allusers = Get-ADUser -filter * -Properties UserPrincipalName,mail -ResultSetSize $null
$allusers | Export-Csv -NoTypeInformation -Delimiter ";" -Encoding unicode -Path c:\temp\allusers.csv

# Import from CSV file
$importusers = import-csv -Path C:\temp\upn-change.csv -Delimiter ";"
# Filtering the users that need to be changed
$newupnusers = $importusers | Where-Object {$_.newupn -notlike ""} 
foreach ($newupnuser in $newupnusers)
{
    $oldupn = $newupnuser.userprincipalname
    $newupn = $newupnuser.newUPN
    Get-ADUser $newupnuser.ObjectGUID | Set-ADUser -UserPrincipalName $newupn
    write-host "changed $oldupn to $newupn"
}
