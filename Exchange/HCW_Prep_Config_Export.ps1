$FormatEnumerationLimit=-1 #avoid truncating output

#export onPrem Config
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://server.contoso.com/PowerShell/ -AllowRedirection -Authentication Kerberos
Import-PSSession $Session -DisableNameChecking -ErrorAction Stop

$date = get-date -Format FileDateTimeUniversal
$filepath = "C:\Temp\" + $date + "_HCW_Prep_config_export_EXonPrem"
mkdir $filepath

Get-HybridConfiguration | select * | ConvertTo-Json | Out-File -FilePath "$filepath\Get-HybridConfiguration.json"
Get-RemoteDomain | select * | ConvertTo-Json | Out-File -FilePath "$filepath\Get-RemoteDomain.json"
Get-AcceptedDomain | select * | ConvertTo-Json | Out-File -FilePath "$filepath\Get-AcceptedDomain.json"
Get-EmailAddressPolicy | select * | ConvertTo-Json | Out-File -FilePath "$filepath\Get-EmailAddressPolicy.json"

Get-FederationTrust | select * | ConvertTo-Json | Out-File -FilePath "$filepath\Get-FederationTrust.json"
Get-FederatedOrganizationIdentifier | select * | ConvertTo-Json | Out-File -FilePath "$filepath\Get-FederatedOrganizationIdentifier.json"
Get-OrganizationRelationship | select * | ConvertTo-Json | Out-File -FilePath "$filepath\Get-OrganizationRelationship.json"
Get-AvailabilityAddressSpace | select * | ConvertTo-Json | Out-File -FilePath "$filepath\Get-AvailabilityAddressSpace.json"
Get-IntraOrganizationConnector | select * | ConvertTo-Json | Out-File -FilePath "$filepath\Get-IntraOrganizationConnector.json"

Get-SendConnector | select * | ConvertTo-Json | Out-File -FilePath "$filepath\Get-SendConnector.json"
Get-ReceiveConnector | select * | ConvertTo-Json | Out-File -FilePath "$filepath\Get-ReceiveConnector.json"

Get-AuthConfig | select * | ConvertTo-Json | Out-File -FilePath "$filepath\Get-AuthConfig.json"
Get-AuthServer | select * | ConvertTo-Json | Out-File -FilePath "$filepath\Get-AuthServer.json"
Get-PartnerApplication | select * | ConvertTo-Json | Out-File -FilePath "$filepath\Get-PartnerApplication.json"

#export EXO Config
$FormatEnumerationLimit=-1 #avoid truncating output
connect-exchangeonline

$date = get-date -Format FileDateTimeUniversal
$filepath = "C:\Temp\" + $date + "_HCW_Prep_config_export_EXO"
mkdir $filepath

Get-RemoteDomain  | select * | ConvertTo-Json | Out-File -FilePath "$filepath\Get-RemoteDomain.json"
Get-AcceptedDomain  | select * | ConvertTo-Json | Out-File -FilePath "$filepath\Get-AcceptedDomain.json"
Get-EmailAddressPolicy  | select * | ConvertTo-Json | Out-File -FilePath "$filepath\Get-EmailAddressPolicy.json"

Get-FederationTrust  | select * | ConvertTo-Json | Out-File -FilePath "$filepath\Get-FederationTrust.json"
Get-FederatedOrganizationIdentifier  | select * | ConvertTo-Json | Out-File -FilePath "$filepath\Get-FederatedOrganizationIdentifier.json"
Get-OrganizationRelationship  | select * | ConvertTo-Json | Out-File -FilePath "$filepath\Get-OrganizationRelationship.json"
Get-IntraOrganizationConnector  | select * | ConvertTo-Json | Out-File -FilePath "$filepath\Get-IntraOrganizationConnector.json"

get-InboundConnector  | select * | ConvertTo-Json | Out-File -FilePath "$filepath\get-InboundConnector.json"
get-outboundconnector  | select * | ConvertTo-Json | Out-File -FilePath "$filepath\get-outboundconnector.json"

Get-OnPremisesOrganization  | select * | ConvertTo-Json | Out-File -FilePath "$filepath\Get-OnPremisesOrganization.json"