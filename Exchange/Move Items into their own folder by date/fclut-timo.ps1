# Source: https://gsexdev.blogspot.com/2009/10/moving-items-into-their-own-folder-by.html?utm_source=chatgpt.com

# Mailbox SMTP address to process.
$MailboxName = "user@domain.com"

# Path to EWS Managed API DLL.
$DllPath = "C:\Program Files\Microsoft\Exchange\Web Services\1.0\Microsoft.Exchange.WebServices.dll"

# Page size used when querying Inbox items.
$PageSize = 200

function Write-Log {
   param(
      [Parameter(Mandatory = $true)]
      [string]$Message
   )

   Write-Output ("[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message)
}

function Get-OrCreate-InboxSubFolder {
   param(
      [Parameter(Mandatory = $true)]
      [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,

      [Parameter(Mandatory = $true)]
      [Microsoft.Exchange.WebServices.Data.Folder]$InboxFolder,

      [Parameter(Mandatory = $true)]
      [string]$FolderName
   )

   $folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
   $folderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow
   $nameFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo(
      [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,
      $FolderName
   )

   $folderResults = $Service.FindFolders($InboxFolder.Id, $nameFilter, $folderView)
   if ($folderResults.TotalCount -gt 0) {
      Write-Log "Folder '$FolderName' already exists under Inbox."
      return $folderResults.Folders[0]
   }

   Write-Log "Creating folder '$FolderName' under Inbox."
   $newFolder = New-Object Microsoft.Exchange.WebServices.Data.Folder($Service)
   $newFolder.DisplayName = $FolderName
   $newFolder.Save($InboxFolder.Id)
   return $newFolder
}

function Get-OldestYearOlderThanCurrent {
   param(
      [Parameter(Mandatory = $true)]
      [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,

      [Parameter(Mandatory = $true)]
      [Microsoft.Exchange.WebServices.Data.Folder]$InboxFolder,

      [Parameter(Mandatory = $true)]
      [datetime]$CurrentYearStart
   )

   # Fetch the oldest item that is older than Jan 1 of current year.
   $olderThanCurrentFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan(
      [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,
      $CurrentYearStart
   )

   $itemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1)
   $itemView.OrderBy.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, [Microsoft.Exchange.WebServices.Data.SortDirection]::Ascending)

   $results = $InboxFolder.FindItems($olderThanCurrentFilter, $itemView)
   if ($results.TotalCount -eq 0) {
      return $null
   }

   $item = $results.Items[0]
   $item.Load()
   return $item.DateTimeReceived.Year
}

$todayForFile = Get-Date -Format "yyyy-MM-dd"
$transcriptPath = Join-Path -Path $PSScriptRoot -ChildPath ("Move-Inbox-By-Year-{0}.log" -f $todayForFile)

$transcriptStarted = $false

try {
   Start-Transcript -Path $transcriptPath -Append -ErrorAction Stop
   $transcriptStarted = $true

   Write-Log "Starting Inbox archival script for mailbox '$MailboxName'."
   Write-Log "Transcript file: $transcriptPath"

   if (-not (Test-Path -LiteralPath $DllPath)) {
      throw "EWS DLL not found: $DllPath"
   }

   Write-Log "Loading EWS assembly from '$DllPath'."
   [void][Reflection.Assembly]::LoadFile($DllPath)

   Write-Log "Creating EWS service object (Exchange2013_SP1 for Exchange Server SE)."
   $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)

   $windowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
   $sidbind = "LDAP://<SID=" + $windowsIdentity.User.Value.ToString() + ">"
   $aceuser = [ADSI]$sidbind

   Write-Log "Running autodiscover for '$($aceuser.mail.ToString())'."
   $service.AutodiscoverUrl($aceuser.mail.ToString())

   Write-Log "Binding to Inbox for '$MailboxName'."
   $folderId = New-Object Microsoft.Exchange.WebServices.Data.FolderId(
      [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,
      $MailboxName
   )
   $inboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $folderId)

   $currentYear = (Get-Date).Year
   $currentYearStart = Get-Date -Year $currentYear -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0

   Write-Log "Current year is $currentYear. Items from this year stay in Inbox."
   Write-Log "Finding oldest year with messages older than current year in Inbox."
   $oldestYear = Get-OldestYearOlderThanCurrent -Service $service -InboxFolder $inboxFolder -CurrentYearStart $currentYearStart

   if ($null -eq $oldestYear) {
      Write-Log "No items older than $currentYear found in Inbox. Nothing to move."
      return
   }

   Write-Log "Oldest year found: $oldestYear"
   Write-Log "Processing years from $($currentYear - 1) down to $oldestYear."

   for ($year = $currentYear - 1; $year -ge $oldestYear; $year--) {
      try {
         Write-Log "----- Processing year $year -----"

         $yearStart = Get-Date -Year $year -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0
         $yearEnd = $yearStart.AddYears(1)

         $targetFolder = Get-OrCreate-InboxSubFolder -Service $service -InboxFolder $inboxFolder -FolderName $year.ToString()

         $greaterOrEqual = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo(
            [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,
            $yearStart
         )
         $lessThan = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan(
            [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,
            $yearEnd
         )

         $dateRangeFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection(
            [Microsoft.Exchange.WebServices.Data.LogicalOperator]::And
         )
         $dateRangeFilter.Add($greaterOrEqual)
         $dateRangeFilter.Add($lessThan)

         $yearMoveCount = 0

         while ($true) {
            # Always request from offset 0 because moved items disappear from Inbox.
            $itemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView($PageSize)
            $itemView.OrderBy.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, [Microsoft.Exchange.WebServices.Data.SortDirection]::Ascending)
            $findResults = $inboxFolder.FindItems($dateRangeFilter, $itemView)

            if ($findResults.TotalCount -eq 0) {
               break
            }

            foreach ($mailItem in $findResults.Items) {
               try {
                  $subject = $mailItem.Subject
                  if ([string]::IsNullOrWhiteSpace($subject)) {
                     $subject = "<no subject>"
                  }

                  Write-Log "Moving item ($($mailItem.DateTimeReceived.ToString("yyyy-MM-dd"))) '$subject' to folder '$year'."
                  [void]$mailItem.Move($targetFolder.Id)
                  $yearMoveCount++
               }
               catch {
                        Write-Log "ERROR moving one item for year ${year}: $($_.Exception.Message)"
               }
            }
         }

         Write-Log "Year $year completed. Moved $yearMoveCount item(s)."
      }
      catch {
            Write-Log "ERROR while processing year ${year}: $($_.Exception.Message)"
      }
   }

   Write-Log "Script completed. Inbox now keeps only items from year $currentYear and newer."
}
catch {
   Write-Output "FATAL ERROR: $($_.Exception.Message)"
   throw
}
finally {
   if ($transcriptStarted) {
      Stop-Transcript | Out-Null
   }
}
