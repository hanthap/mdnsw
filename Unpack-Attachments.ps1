# Before starting or resuming Stage 2, we (re-)load our enriched metadata into RAM, as a hashtable.

$attachment = @{}
Import-Csv "$unzippedRoot\Attachment-Map.csv" -Encoding UTF8 | 
ForEach-Object { $attachment[$_.Id] = $_ } 


$attachment['00PPr0000057BS5MAM']

#--------------------------------------------------------------------------------------------


$shell = New-Object -Com Shell.Application

$source_namespace = $shell.NameSpace( $zipFileStem + '34.zip' )
$target_namespace = $shell.Namespace( $unzippedRoot )

# This triggers download of the zip file, if not currently synced
$source_folder = $source_namespace.Items() | where Name -eq 'Attachments' 


# superseded
# $target_namespace.CopyHere( $source_folder ) # in Downloads, creates a subfolder called Attachments, complete with all its contents
# can we now unpin the zip file?
# $target_folder = $target_namespace.Items() | where Name -eq 'Attachments'


# list the Attachments in the zip file, before unzipping
$source_folder.GetFolder().Items() | 
select -first 10 Name, Path, @{n='mapped';e={$attachment[$_.Name]}} |
select Name, Path,
@{n='unique_fname';e={$_.mapped.unique_fname}},
@{n='doclib';e={$_.mapped.doclib}},
@{n='folder';e={$_.mapped.folder}},
@{n='CreatedDate';e={$_.mapped.CreatedDate}},
@{n='noise_level';e={$_.mapped.noise_level}},
@{n='type';e={$_.mapped.type}} |
where noise_level -lt 10



$attachment['00P3b00001eXjJGEA0'] 


