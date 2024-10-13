<#

# SPO Staging folder will contain at least 35 .zip files 

# How to trigger Weekly Data Export: 
# https://help.salesforce.com/s/articleView?id=sf.admin_exportdata.htm&type=5

# https://mdnsw.lightning.force.com/lightning/setup/DataManagementExport/home

#>

<# NO LONGER IN USE

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
#>


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
      [Parameter(Mandatory)][int]$suffix, # store the batch identifier in place of 'seconds' in the file's timestamp
      [int] $MaximumNoiseLevel = 9
      )
    # TO DO : begin block: if $attachment hashtable doesn't exist then load it

process {
    $id =  $f.Name # the unzipped raw file item is named as per its case-safe Attachment.Id (with no extension)
    $d = $attachment.$Id # get the metadata
    if ( $d ) { # if mapping exists
        # $d
        if ( [int]$d.noise_level -gt $MaximumNoiseLevel ) {
            $out_folder = "$env:OneDrive\$($d.doclib)`\_NOISE_`\$($d.type)`\$($d.folder)"
        } else {
            $out_folder = "$env:OneDrive\$($d.doclib)`\$($d.type)`\$($d.folder)"
            }
        $out_path = "$out_folder\" + $d.unique_fname
        Write-Verbose $out_path
        $utc = [DateTime] $d.CreatedDate # casting straight from ISO doesn't require ParseExact
        $utc = $utc.AddSeconds( $suffix – $utc.Second ) # replace actual seconds with our batch identifier
        $f.LastWriteTimeUtc = $utc # successfully propagates to "Modified" datestamp when sync'ed to SPO
        New-Item -ItemType Directory -Force -Path $out_folder | Out-Null  # -Force adds intermediate subfolders 
        $f.MoveTo($out_path)
        attrib -p +u $out_path # we can unpin it immediately to free up space

        }
} # end process

}

#-----
# Documents use a different logic for file path

function Rename-Document { 

    [CmdletBinding()]
    param(
      [Parameter(Mandatory, ValueFromPipeline)] [PSObject] $f,
      [Parameter(Mandatory)][int]$suffix # store the batch identifier in place of 'seconds' in the file's timestamp
      )

process {
    $id =  $f.Name # the unzipped raw file item is named as per its case-safe Document.Id (with no extension)
    $d = $document.$Id # get the metadata
    if ( $d ) { # if mapping exists
        $out_folder = "$env:OneDrive\$($d.doclib)`\$($d.folder)"
        $out_path = "$out_folder\" + $d.unique_fname
        Write-Verbose $out_path
        $utc = [DateTime] $d.CreatedDate 
        $utc = $utc.AddSeconds( $suffix – $utc.Second ) # replace actual seconds with our batch identifier
        $f.LastWriteTimeUtc = $utc 
        New-Item -ItemType Directory -Force -Path $out_folder | Out-Null 
        $f.MoveTo($out_path)
        attrib -p +u $out_path
        }
} # end process

}