$filename = "C:\Temp\Confidential Document.docx"
$zip = [System.IO.Compression.ZipFile]::Open($filename, 'Read')
$propsentry = $zip.GetEntry('docProps/custom.xml')
If ($propsentry -ne $null) {
    $stream = $propsentry.Open()
    $reader = New-Object System.IO.StreamReader $stream
    $content = $reader.ReadToEnd()
    $xmldoc = [xml]$content
    $xmldoc.Properties.property | Select-Object name,lpwstr
}
$zip.Dispose()
