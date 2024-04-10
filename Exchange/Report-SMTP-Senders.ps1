# Log Parser 2.2 needed
# run from cmd and not ps!
cd "C:\Temp\SMTPLogs"
"C:\Program Files (x86)\Log Parser 2.2\LogParser.exe" "SELECT remote-endpoint,Count(*) as Hits from *.log WHERE data LIKE '%MAIL FROM%' GROUP BY remote-endpoint ORDER BY Hits DESC" -i:CSV -nSkipLines:4 -rtp:-1 > c:\temp\SMTPLogs\exchange.txt

# 1. remove multiple spaces and save to a new file
$inputfile = "C:\Temp\SMTPLogs\exchange.txt"
$normfile = "C:\Temp\SMTPLogs\exchange_norm.txt"
$lpsoutput = get-content $inputfile
$lpsoutputnorm = $lpsoutput -replace '\s+', ' '
$lpsoutputnorm | out-file $normfile

# 2. manual step -> manipulate normfile with notepad: remove second line and custom infos at the end!

# 3. cut the source port from remote-endpoint
$lpstable = Import-Csv $normfile -Delimiter " "
$export = @()
foreach ($line in $lpstable)
{
    $psobject = New-Object -TypeName psobject
    $IP = $line.'remote-endpoint'.Substring(0, $line.'remote-endpoint'.IndexOf(":"))
    $psobject | Add-Member -MemberType NoteProperty -Name IP -Value $IP
    $psobject | Add-Member -MemberType NoteProperty -Name Hits -Value $line.Hits
    $export += $psobject
}

# 4. merge duplicates and sum up the hits
$result = @{}

foreach ($row in $export)
{
    # check if the IP-Address is already in the result-hash
    if ($result.ContainsKey($row.IP))
    {
        # if true, add the value from the column hits to the current value
        $result[$row.IP] += [int]$row.Hits
    } 
    else 
    {
        # if false add the ip address to the list and add the value from the column hits to the result hash
        $result[$row.IP] = [int]$row.Hits
    }
}
# 5. export to a csv file
$result.GetEnumerator() `
    | Sort-Object value -Descending `
    | Select-Object name,value `
    | Export-Csv "C:\Temp\SMTPLogs\exchange.csv"