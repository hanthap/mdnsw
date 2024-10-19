# Before starting or resuming Stage 2, we (re-)load our enriched metadata into RAM, as a hashtable.

$attachment = @{}
Import-Csv "$unzippedRoot\Attachment-Map.csv" -Encoding UTF8 | 
ForEach-Object { $attachment[$_.Id] = $_ } 

$document = @{}
Import-Csv "$unzippedRoot\Document-Map.csv" -Encoding UTF8 | 
ForEach-Object { $document[$_.Id] = $_ } 

$attachment.Count # 36237
$document.count # 9628

$attachment['00P3b00001T3bu4EAB']


$attachment.'00P3b00001T2Nc5EAF'

$attachment.'00P5c00001qXWxBEAW'
#--------------------------------------------------------------------------------------------

$shell = New-Object -Com Shell.Application
$unzipped_namespace = $shell.NameSpace( $unzippedRoot )

# #5 is the first to include Attachments. 
$suffix = 35
$zip_file_path = "$zipFileStem$suffix.zip"
$source_namespace = $shell.NameSpace( $zip_file_path )

# This triggers download of the zip file (if not currently synced)
$from_zip_folder = $source_namespace.Items() | where Name -eq 'Attachments' 
$unzipped_namespace.CopyHere( $from_zip_folder ) # in Downloads, creates a subfolder called Attachments, complete with all its contents

$from_zip_folder = $source_namespace.Items() | where Name -eq 'Documents' 
$unzipped_namespace.CopyHere( $from_zip_folder ) # in Downloads, creates a subfolder called Documents, complete with all its contents

attrib -p +u $zip_file_path # immediately un-pin the zip file to free up disk space

# scan the local Attachments folder for any newly unzipped files (those not already moved/renamed)
Get-ChildItem -Path "$unzippedRoot\Attachments" -File | 
    Rename-Attachment -suffix $suffix -Verbose |
    Export-Csv "$unzippedRoot\rename_a_$suffix.csv" -NoTypeInformation

# Ditto for Documents
Get-ChildItem -Path "$unzippedRoot\Documents" -File | 
    Rename-Document -suffix $suffix -Verbose |
    Export-Csv "$unzippedRoot\rename_d_$suffix.csv" -NoTypeInformation

Write-Host "Completed extraction from archive: $zip_file_path"

# summary stats
Import-Csv "$unzippedRoot\rename_a_$suffix.csv" | 
    group doclib,type -NoElement
