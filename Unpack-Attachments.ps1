# Before starting or resuming Stage 2, we (re-)load our enriched metadata into RAM, as a hashtable.

$attachment = @{}
Import-Csv "$unzippedRoot\Attachment-Map.csv" -Encoding UTF8 | 
ForEach-Object { $attachment[$_.Id] = $_ } 


$attachment['00PPr0000057BS5MAM']

#--------------------------------------------------------------------------------------------

$shell = New-Object -Com Shell.Application
$unzipped_namespace = $shell.NameSpace( $unzippedRoot )


$suffix = 7
$zip_file_path = "$zipFileStem$suffix.zip"
$source_namespace = $shell.NameSpace( $zip_file_path )

# This triggers download of the zip file (if not currently synced)
$from_zip_folder = $source_namespace.Items() | where Name -eq 'Attachments' 
$unzipped_namespace.CopyHere( $from_zip_folder ) # in Downloads, creates a subfolder called Attachments, complete with all its contents
attrib -p +u $zip_file_path # immediately un-pin the zip file to free up disk space

# scan the local Attachments folder for any newly unzipped files (those not already moved/renamed)
Get-ChildItem -Path "$unzippedRoot\Attachments" -File | #select -first 300 |
    Rename-Attachment -suffix $suffix -Verbose 



#------------------------------------

$onedriveProcess = Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue

$onedrive_namespace = $shell.NameSpace( $env:OneDrive )


$Status = $onedrive_namespace.getDetailsOf( '.', 303 )

$onedriveProcess | get-member


try {

$syncing = $true

while ($syncing) {

    $syncStatus = $onedriveProcess.Process.MainWindowTitle

    if ($syncStatus -eq "OneDrive is up to date") {

        $syncing = $false

    } else {

        Start-Sleep -Seconds 5

      Break

    }

}