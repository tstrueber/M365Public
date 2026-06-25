# Move Items Into Yearly Inbox Subfolders

This folder contains a PowerShell script that archives Inbox items into yearly subfolders by using Exchange Web Services (EWS).

## Files

- `fclut-timo.ps1`: Updated script with logging, transcript creation, yearly processing, and error handling.
- `fclut-original.ps1`: Original variant for reference.

## What The Script Does

The script processes one mailbox Inbox and keeps only current-year emails in Inbox.

Procedure:

1. Starts a transcript log file with the current date in the script folder.
2. Loads the EWS Managed API assembly.
3. Creates an EWS service object and runs Autodiscover.
4. Binds to the target mailbox Inbox.
5. Detects the current year and identifies the oldest year still present in Inbox (older than current year).
6. Loops backward year-by-year (for example: 2025, 2024, 2023, ...).
7. For each year:
   - Ensures an Inbox subfolder with the year name exists (creates it if missing).
   - Finds Inbox items received in that year.
   - Moves those items to the matching year folder.
8. Stops when no items older than the current year remain in Inbox.
9. Ends transcript logging.

## Logging And Output

- Console output includes timestamped progress messages.
- A transcript is written to:
  - `Move-Inbox-By-Year-YYYY-MM-DD.log`
- The transcript contains all script output and errors from the run.

## Error Handling

The script uses `try/catch/finally` at multiple levels:

- Global level for startup and overall execution.
- Per-year level so one year failing does not stop all years.
- Per-item level so one message failure does not stop a year batch.

## Requirements

- Exchange environment with EWS access.
- EWS Managed API DLL present (path configured in the script).
- Mailbox permissions to access and move Inbox items.
- PowerShell session running under an account that can complete Autodiscover and mailbox operations.

## Configuration

Edit these values in `fclut-timo.ps1` before running:

- `$MailboxName` (target mailbox SMTP address)
- `$DllPath` (path to Microsoft.Exchange.WebServices.dll)
- `$PageSize` (batch size for item processing)

## Original Source

Original concept and source inspiration:

- Glen Scales blog post: https://gsexdev.blogspot.com/2009/10/moving-items-into-their-own-folder-by.html?utm_source=chatgpt.com
