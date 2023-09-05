#mailbox UPN:
$mailbox = "user@contoso.com"

# set Application (client) ID, tenant Name and secret 
$clientId = ""
$tenantId = ""
$clientSecret = ""

# build request body
$ReqTokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    client_Id     = $clientID
    Client_Secret = $clientSecret
} 

# get token
$TokenResponse = $null
$TokenResponse = Invoke-RestMethod `
    -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
    -Method POST -Body $ReqTokenBody

# request and write the response to a variable
$apiUrl = "https://graph.microsoft.com/v1.0/users/$mailbox/messages/"
$Data = Invoke-RestMethod `
    -Headers @{ Authorization = "Bearer $($Tokenresponse.access_token)" } `
    -Uri $apiUrl -Method Get

# formated output of the mails
$mails = $Data.Value
$mails | select-object receivedDateTime,subject, `
    @{ label='sender'; expression={($_.sender.emailAddress.address)} }