# Source: https://gsexdev.blogspot.com/2009/10/moving-items-into-their-own-folder-by.html?utm_source=chatgpt.com
$MailboxName = "user@domain.com"
$StartDate = [system.DateTime]::Today.AddDays(-1)
$EndDate = [system.DateTime]::Today

$dllpath = "C:\Program Files\Microsoft\Exchange\Web Services\1.0\Microsoft.Exchange.WebServices.dll"
[void][Reflection.Assembly]::LoadFile($dllpath)
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1)

$windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$sidbind = "LDAP://<SID=" + $windowsIdentity.user.Value.ToString() + ">"
$aceuser = [ADSI]$sidbind

$service.AutodiscoverUrl($aceuser.mail.ToString())

$folderid = new-object  Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)
$InboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
$Sfgt = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $StartDate)
$Sflt = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $EndDate)
$sfCollection = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And);
$sfCollection.add($Sfgt)
$sfCollection.add($Sflt)
$view = new-object Microsoft.Exchange.WebServices.Data.ItemView(2000)
$frFolderResult = $InboxFolder.FindItems($sfCollection,$view)
$NewFolder = new-object Microsoft.Exchange.WebServices.Data.Folder($service)
$NewFolder.DisplayName = $EndDate.ToString("yyyy-MM-dd")
$NewFolder.Save($InboxFolder.Id.UniqueId)
foreach ($miMailItems in $frFolderResult.Items){
   "Moving" + $miMailItems.Subject.ToString()
   [VOID]$miMailItems.Move($NewFolder.Id.UniqueId)
}
