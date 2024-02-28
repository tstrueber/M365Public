Connect-ExchangeOnline

$forwarders = `
    get-content "C:temp\forwarders.json" `
    | ConvertFrom-Json

foreach($forwarder in $forwarders)
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