
$urlRoot = "https://mdnsw.my.salesforce.com/sfc/servlet.shepherd/version/download"

# $outFolder = "$outRoot\20241017\ContentVersionFile"
$outFolder = "$downloads\ContentDocument"


$ErrorActionPreference = 'Stop'

$d = Import-Csv "$unzippedRoot\ContentVersion.csv" -Encoding UTF8 | 
    where IsLatest -eq 1 | 
    select -first 500 -Skip 5000 

$d.Count

$d | select ContentDocumentId, PathOnClient, Title, LastModifiedDate, 
        @{ n="VersionDataUrl"; e={ "$urlRoot/$($_.Id.substring(0,15))" } },
        @{ n="Extension"; e={ [System.IO.Path]::GetExtension($_.PathOnClient) } } |
    ForEach-Object {
        Start-Process MSEdge $_.VersionDataUrl -Wait -WindowStyle Minimized # the download is async, therefore -Wait doesn't help
        do { # pause until there are no downloads pending
            sleep 1
            $c = Get-ChildItem -Path "$downloads" -Filter '*.crdownload' | select -First 1
            } until ( $c.count -eq 0 )
        # oddly, a few cases where the newest file is NOT the one we just downloaded, but is instead a .tmp file
        # result is that the real file is not moved/renamed 
        $f = Get-ChildItem -Path "$downloads" | sort CreationTime -Descending | select -First 1 # ASSUMES SUCCESSFUL DOWNLOAD
        $f.Name | Out-Host
        $out_path = "$outfolder\$($_.ContentDocumentId)$($_.Extension)"
        #$out_path  | Out-Host
        try { 
            $ef = $null
            # Remove-Item -Path $out_path -ErrorAction Ignore
            Move-Item -Path $f.FullName -Destination $out_path -Force 
            $f = Get-Item -Path $out_path
            $f.LastWriteTimeUtc = [DateTime] $_.LastModifiedDate # need to do this after the other operations
            attrib -p +u $out_path # we can unpin it immediately to free up space
            } catch { $ef = $_ } 
        [pscustomobject] @{ f=$f; err=$ef; cd=$_ }
        } |
    Export-Csv "$outFolder\Export.csv"

   


$nameList = Get-ChildItem $downloads -Filter '*.pdf' | % { $_.Name }
$nameList -contains 'Lok Sum Volunteer_Camp_Carer_-_Position_Description.docx.pdf'

$nameList.Count

$idList = Get-ChildItem $outfolder -Filter '*.pdf' | where LastWriteTime -gt (Get-Date).AddDays(-1) | % {  [System.IO.Path]::GetFileNameWithoutExtension( $_.Name ) }

$idList = Get-ChildItem $downloads -Filter '*.pdf' | where LastWriteTime -gt (Get-Date).AddDays(-1) | % {  [System.IO.Path]::GetFileNameWithoutExtension( $_.Name ) }



$nameList = Get-ChildItem $downloads -Filter '*.docx' | % { $_.Name }
$nameList.Count

$d = Import-Csv "$unzippedRoot\ContentVersion.csv" -Encoding UTF8 | 
    where IsLatest -eq 1 | 
 #   select -first 2500 |
    where { $nameList -contains $_.PathOnClient } 

$d.Count
