<#

Pre-upload - AFTER generating master lookup qryAttachmentScope.csv

Given a complete set of 34 archive zip files (manually downloaded from Salesforce) pick one and unzip its contents 
(If resources are scarce you can break it into smaller batches)
Apply Vertic's file naming standards to the unzipped files, ready for uploading to the SharePoint doclib
Names etc are determined using output from Access stage 1
#>

# SPO Staging folder will contain at least 34 .zip files 


# How to trigger Weekly Data Export: 
# https://help.salesforce.com/s/articleView?id=sf.admin_exportdata.htm&type=5

# https://mdnsw.lightning.force.com/lightning/setup/DataManagementExport/home


# subset of files to be unzipped

function Expand-Archive {
    [CmdletBinding()]
    param ( 
        [string]$Stem = $zipFileStem,  
        [string]$Suffix = 34,
        [string]$ZipPath = $Stem + $Suffix + '.zip',
        [string]$OutPath = $unzippedRoot,
        [switch]$Pin= $false
    )

$shell = New-Object -Com Shell.Application
#unzip all objects into our temporary staging folder
    $zip = $shell.NameSpace($ZipPath)
    Write-Verbose "Downloading $ZipPath"
    $flist = $zip.Items() # will trigger auto-download from SPO, if not already cached
    # BUT lately that doesn't happen on first go. Download hangs and script continues as if all done
    Write-Verbose "Expanding to $OutPath "
    $shell.Namespace($OutPath).CopyHere($flist)

if ( -not $Pin ) {
# default: free up disk space by unpinning the local .zip copy from OneDrive cache
Write-Verbose "Free up disk space: $ZipPath"
attrib -p +u $ZipPath /s
}

} 

    $shell.Namespace($unzippedRoot).CopyHere($flist)


Expand-Archive -Suffix 29 -Pin -Verbose  #29 done, next is 28


function Rename-Attachment { 

<#
.SYNOPSIS
    Checks metadata hashtable $map for an entry with matching AttachmentID. 
    If a matching entry is found and not flagged as 'noise':
        Restores the original extension and filename (suffixed with Salesforce ID)
        Moves it into the designated cache folder for auto-upload to SharePoint doclib
    Else: 
        Moves the file into "skipped" folder

.DESCRIPTION
    Rename-Attachment is a function that moves a file

.PARAMETER File
    A file object, via pipeline


.EXAMPLE
     Rename-Attachment 

.INPUTS
    String

.OUTPUTS
    PSCustomObject

.NOTES
    Author:  Peter Luckock
#>

    [CmdletBinding()]
    param(
      [Parameter(Mandatory, ValueFromPipeline)] [PSObject] $f,
      [Parameter(Mandatory)][int]$suffix # store the batch identifier in place of 'seconds' in the file's timestamp
      )
    # TO DO : begin block: if $map doesn't exist then load it

process {
    $id =  $f.Name # the unzipped raw file item is named as per its case-safe Attachment.Id (with no extension)
    $d = $map.$Id # get the metadata
    if ( $d ) { # if mapping exists
        $c = [char]$d.ContactLastName.Substring(0,1) # first letter of surname also determines which doclib subset
        $subFolder = 'Subset' + $c + '\' + $d.OutCategory + '\' + $c + '\' + $d.ContactFullName + ' #' + $d.ContactID  
        $subFolder = $subFolder -replace '\?', '''' # replace punctuation characters not allowed in a filepath
        $utc = [DateTime]::ParseExact($d.CreatedDate,"d/M/yyyy H:mm:ss",$null)
        $utc = $utc.AddSeconds( $suffix – $utc.Second ) # replace actual seconds with our batch identifier
        $_.CreationTimeUtc = $utc # this works in local filesystem only, not the SharePoint doclib item, sadly
        $_.LastWriteTimeUtc = $utc # successfully propagates to "Modified" datestamp when sync'ed to SPO
        $ts = $utc.ToString("yyyy-MM-dd" ) # not required as we have the separate timestamp attribute
        # filename ends with the old case-safe Attachment ID, to be paired with the uniqueID added by SharePoint when we upload
        $base = [System.IO.Path]::GetFileNameWithoutExtension($d.Name)  -replace '\?', '''' # replace punctuation characters not allowed in a filename
        $ext = [System.IO.Path]::GetExtension($d.Name)

        $fName = $base + ' #' + $id + $ext

        if ( $d.Priority -gt 0 ) {
            New-Item -ItemType Directory -Force -Path "$outRoot\$subFolder" | Out-Null  # -Force adds intermediate subfolders 
            $outPath = "$outRoot\$subFolder\$fName"
        } else {
            $outPath = "$skipFolder\$fName"
            }
        $outPath
        $f.MoveTo($outPath)
    } else {
        $f.MoveTo( "$skipFolder\$id" ) # some Attachment IDs are not in $map eg due to "unresolved email"
        }
} # end process

}



New-Item -ItemType Directory -Force -Path $skipFolder # in case I've deleted it 

# scan the local folder for any unzipped files that have not already been moved+renamed (or skipped)
Get-ChildItem -Path "$unzippedRoot\Attachments" -File | 
    Rename-Attachment


attrib -p +u $outRoot\S2\* /s


# Once syncing is complete, we will use the SharePoint API to generate an incremental CSV file containing the uniqueID and original SF ID
# A macro in Access will then upsert its master table using the incremental data
