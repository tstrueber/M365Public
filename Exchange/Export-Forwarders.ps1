$AllMailboxes = Get-Mailbox -ResultSize Unlimited
$InternalForwarders = $AllMailboxes | where-object { $_.ForwardingAddress -ne $null }
$GenericForwarders = $AllMailboxes | where-object { $_.ForwardingSmtpAddress -ne $null }

write-host "Internal Forwarders:      " $InternalForwarders.Count
write-host "SMTP Address Forwarders:  " $GenericForwarders.Count

$AllForwarders = $InternalForwarders + $GenericForwarders

$export = @()
foreach($forwarder in $AllForwarders)
{
    $psobject = New-Object -TypeName psobject
    
    $resolvedForwardingAddress = $null
    $resolvedForwardingAddress = (Get-Recipient $forwarder.ForwardingAddress).primarysmtpaddress
    
    $psobject | Add-Member -MemberType NoteProperty -Name PrimarySmtpAddress -Value $forwarder.PrimarySmtpAddress
    $psobject | Add-Member -MemberType NoteProperty -Name ForwardingAddress -Value $forwarder.ForwardingAddress
    $psobject | Add-Member -MemberType NoteProperty -Name resolvedForwardingAddress -Value $resolvedForwardingAddress
    $psobject | Add-Member -MemberType NoteProperty -Name ForwardingSmtpAddress -Value $forwarder.ForwardingSmtpAddress
    $psobject | Add-Member -MemberType NoteProperty -Name DeliverToMailboxAndForward -Value $forwarder.DeliverToMailboxAndForward
    $export += $psobject
}
$date = get-date -f yyyy-MM-dd
$filename = $date + "_forwarders.json"
$export `
    | Select-Object PrimarySmtpAddress,ForwardingAddress,resolvedForwardingAddress,ForwardingSmtpAddress,DeliverToMailboxAndForward -Unique `
    | ConvertTo-Json | out-file "C:\temp\$filename"