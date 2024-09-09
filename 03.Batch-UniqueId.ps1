<#

After uploading each batch of documents, before ingesting incremental xref data into Access

Each staging folder SubsetA..SubsetZ must have fewer than 5000 items, otherwise you get "The attempted operation is prohibited because it exceeds the list view threshold."

We'll run 

#>

# Run this step when doclibs SubsetA..SubsetZ are non-empty (before moving contents over to "Legacy")

<#
   
   To initialise a valid WebRequestSession object:
    In Chrome, go to Developer Tools, find the Network tab
    Open the SPO site url via the address bar
    In the list of requests, right-click the relevant request and select "Copy > Copy as Powershell" from the menu.
    Paste the copied code into a new .PS1 script window
    Select & run the $session statement, and the added cookies
    Then run the code below, including $headers
#>

<#

$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$session.UserAgent = "..."
$session.Cookies.Add((New-Object System.Net.Cookie(...)
...
$session.Cookies.Add((New-Object System.Net.Cookie(...)

#>

$headers = @{
  "authority"="musculardystrophynsw.sharepoint.com"
  "accept"="application/json" # necessary 
  "accept-encoding"="gzip, deflate, br, zstd"
}

# ASSUME: we've partioned the forest into 26 doclibs named SubsetA...SubsetZ

$csvPath ="$downloads\uniqueId.csv"

if ( Test-Path $csvPath ) { Remove-Item $csvPath }


foreach( $a in 0..25 ) {


    $b = [char]($a+[char]'A')

    $uri = "https://musculardystrophynsw.sharepoint.com/sites/SalesForceDocuments/_api/Web/Lists/GetByTitle('Subset$b')/items?`$expand=File&`$select=File&`$filter=FSObjType%20eq%200"

    do {
        $uri

        $x = Invoke-RestMethod -UseBasicParsing -Uri $uri -WebSession $session -Headers $headers 
 
        # $x.value is an array of PSObjects, parsed from JSON the raw http response JSON

        # export the dataset as CSV
        $x.value | % { $_.File } | 
                Select-Object UniqueId, Length, Name, ServerRelativeUrl, TimeLastModified | 
                    % { 
                        $base = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
                        $ext = [System.IO.Path]::GetExtension($_.Name)
                        $obase = $base.Substring(0,$base.Length-20).Trim()
                        $oname = $obase + $ext
                        $id = $base.split('#')[-1] # last token
                        $folder = [System.IO.Path]::GetDirectoryName($_.ServerRelativeUrl)
                        $a = $folder.Split('\')
                        $contactid = $a[-1].Split('#')[-1]
                        $contactname = $a[-1].Split('#')[0].Trim()
                        $category,$subset = $a[4,5]

                        $_ | Add-Member -Name 'SourceName' -Type NoteProperty -Value $oname
                        $_ | Add-Member -Name 'AttachmentId' -Type NoteProperty -Value $id
                        $_ | Add-Member -Name 'Category' -Type NoteProperty -Value $category
                        $_ | Add-Member -Name 'Subset' -Type NoteProperty -Value $subset
                        $_ | Add-Member -Name 'ContactId' -Type NoteProperty -Value $contactid
                        $_ | Add-Member -Name 'ContactName' -Type NoteProperty -Value $contactname
                        $_ | Add-Member -Name 'Folder' -Type NoteProperty -Value $folder

                        $_
                    } |
                    Export-Csv -NoTypeInformation -Path $csvPath -Append

        $uri = $x.'odata.nextLink'

    } while ( $uri )

} # end for loop


# now go to Access and run macro "tblSPFileMaster: Upsert from Staging CSV"

# check rowcounts are as expected. If all good then go to SharePoint UI and move subfolders 'A', 'B'... 'Z' from S0, S1, ...S5 to the corresponding parent subfolder under "Legacy"
# (SP won't merge contents if the same folder already exists in destination)


