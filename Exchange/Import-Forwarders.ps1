Connect-ExchangeOnline

$forwarders = `
    get-content "C:temp\forwarders.json" `
    | ConvertFrom-Json

$importlist = $forwarders | Where-Object{$_.PrimarySmtpAddress -match "contoso.com" } # filter the list
$importlist `
    | sort-object PrimarySmtpAddress `
    | select-object PrimarySmtpAddress,resolvedForwardingAddress,ForwardingSmtpAddress,DeliverToMailboxAndForward

foreach($forwarder in $importlist)
{
    if($forwarder.ForwardingSmtpAddress -ne $null)
    {
        # when forwarding smtp address is set
        set-mailbox $forwarder.PrimarySmtpAddress `
            -ForwardingSmtpAddress $forwarder.ForwardingSmtpAddress `
            -DeliverToMailboxAndForward $forwarder.DeliverToMailboxAndForward
        Write-Host "Forwarding" $forwarder.PrimarySmtpAddress "to" $forwarder.ForwardingSmtpAddress "- DeliverToMailboxAndForward:" $forwarder.DeliverToMailboxAndForward
    }
    else
    {
        # when forwarding to a specific object is set
        set-mailbox $forwarder.PrimarySmtpAddress `
            -ForwardingSmtpAddress $forwarder.resolvedForwardingAddress `
            -DeliverToMailboxAndForward $forwarder.DeliverToMailboxAndForward
        Write-Host "Forwarding" $forwarder.PrimarySmtpAddress "to" $forwarder.resolvedForwardingAddress "- DeliverToMailboxAndForward:" $forwarder.DeliverToMailboxAndForward
    }
}