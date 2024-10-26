
$contentdoc = @{}
Import-Csv "$unzippedRoot\OldOrg\ExportPlus\ContentDocument-Map.csv" -Encoding UTF8 | 
where Folder -ne '' |
ForEach-Object { 
 $_.UniqueFileName =  $_.UniqueFileName -replace '[:\?]', ''
 $_.Folder =  $_.Folder -replace '[:\?]', ''
 $contentdoc[$_.ContentDocumentId] = $_ 
 } 


 $contentdoc.count # 5076 => 5023

#--------------------------------------------------------------------------------------------------

function Rename-ContentDocument { 

    [CmdletBinding()]
    param(
      [Parameter(Mandatory, ValueFromPipeline)] [PSObject] $f,
      [Parameter(Mandatory)][int]$suffix # store the batch identifier in place of 'seconds' in the file's timestamp
      )

process {
    $err = $null
    $id = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
    $d = $contentdoc.$Id # get the metadata
    if ( $d ) { # if mapping exists
        $out_folder = "$env:OneDrive\$($d.doclib)`\$($d.folder)"
        $out_path = "$out_folder\" + $d.UniqueFileName
        Write-Verbose $out_path
#        $utc = [DateTime] $d.CreatedDate 
#        $utc = $utc.AddSeconds( $suffix – $utc.Second ) # replace actual seconds with our batch identifier
#        $f.LastWriteTimeUtc = $utc 
        try { 
            New-Item -ItemType Directory -Force -Path $out_folder | Out-Null  # -Force adds intermediate subfolders 
            $f.MoveTo($out_path)
            attrib -p +u $out_path # we can unpin it immediately to free up space
            } catch { $err = $_ } 
        }
    [pscustomobject]@{ suffix=$suffix; fname=$id; length=$f.length; dest=$d; ts=Get-Date; err=$err }
    } # end process

}

#-----------------------------------------------------------------------------------------------------


$shell = New-Object -Com Shell.Application
$unzipped_namespace = $shell.NameSpace( "$unzippedRoot\ContentDocument" )

# #5 is the first to include Attachments. 
$suffix = '12'
$zip_file_path = "$zipFileStemCD$suffix.zip"
$source_namespace = $shell.NameSpace( $zip_file_path )

# This triggers download of the zip file (if not currently synced)
$from_zip_folder = $source_namespace.Items() 
$unzipped_namespace.CopyHere( $from_zip_folder ) # in Downloads

attrib -p +u $zip_file_path # immediately un-pin the zip file to free up disk space

# scan the local Attachments folder for any newly unzipped files (those not already moved/renamed)
Get-ChildItem -Path "$unzippedRoot\ContentDocument" -File | 
    Rename-ContentDocument -suffix $suffix -Verbose |
    Export-Csv "$unzippedRoot\rename_c_$suffix.csv" -NoTypeInformation

 $contentdoc.'0693b000007DIXiAAO'