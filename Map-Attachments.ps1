<#

Step 1: prepare a mapping lookup table and save as csv.

Step 2: iterate through zip files, rename & move each unzipped attachment into its pre-mapped local path so that OneDrive can quietly upload in background (and promptly free up local disk space when done)

analyse Attachment.csv to identify commonest rubbish files based on filetype and size
save as an exclusion hashtable

load hashtable with all Attachment Ids in scope (as key) and their SharePoint destination filepath etc. (as value)
exclude noise files, where an image is found to be recurring with more than N times, (based on filetype and body length)


process zip files one by one
for each zip file in folder
 
   download & unzip one archive into unzipped folder
   unpin, to release local cache of zip file 

   for each unzipped file
        
        lookup its target doclib, subfolder & unique filename
        move to that path on local device
        free up space immediaely (No need for API extract with SharePoint ID, Therefore no need for workaround with separate partitions any more.. )



#>

#------------------------------------------------------------

# Identify repeating images and their incidence, so we can choose a cutoff later. Store as a hashtable so we can use it in a left join

$noise_image = @{}
Import-Csv "$unzippedRoot\Attachment.csv" -Encoding UTF8 | 
Where-Object ContentType -Like 'image*' |  #this filter makes the Group run much faster, no idea why.
Group-Object ContentType, BodyLength -NoElement | 
Where-Object Count -gt 1 | 
ForEach-Object { $noise_image[$_.Name] = $_.Count } # add to hashtable

#------------------------------------------------------------

# from Contacts.csv, populate a mapping hashtable with key = id, value = destination folder name : "SURNAME, Firstname #Id" 

$TextInfo = (Get-Culture).TextInfo # for Title Case
$contact = @{}
Import-Csv "$unzippedRoot\Contact.csv" -Encoding UTF8 |
# Append ID so as to avoid combining namesakes "SMITH, John"
Select-Object Id, FirstName, LastName, 
@{n='folder_name'; e={ 
    $s = $_.LastName.ToUpper()+', '+$TextInfo.ToTitleCase($_.FirstName.ToLower())+' #'+$_.id 
    $s[0] + '\' + $s
    }} |
ForEach-Object { $contact[$_.Id] = $_ } 
$contact.Count # 39564 => 39567

$contact.'0033b00002UVlGBAA1'

#------------------------------------------------------------

# Same for Accounts

$account = @{}
Import-Csv "$unzippedRoot\Account.csv" -Encoding UTF8 | 
Select-Object Id, Name, @{ n='folder_name'; e={ $_.Name.ToUpper()+' #'+$_.Id }} |
ForEach-Object { $account[$_.Id] = $_ } 
$account.Count # 4390


#------------------------------------------------------------
# Some tasks only have a What and no Who. Of these, only some have a WhatId that points to a Campaign Id.

$campaign = @{}
Import-Csv "$unzippedRoot\Campaign.csv" -Encoding UTF8 | 
Select-Object Id, Name, @{ n='folder_name'; e={ $_.Name +' #'+$_.Id }} |
ForEach-Object { $campaign[$_.Id] = $_ } 
$campaign.Count #  7771

#------------------------------------------------------------

# Tasks

$task = @{}
Import-Csv "$unzippedRoot\Task.csv" -Encoding UTF8 | 
Select-Object Id, WhoId, WhatId, AccountId, Client_Name__c, Subject, 
@{ n='who'; e={ $contact[$_.WhoId] } }, 
@{ n='what'; e={ $campaign[$_.WhatId] } }, 
@{ n='client'; e={ $contact[$_.Client_Name__c] } }, 
@{ n='account'; e={ $account[$_.AccountId] } } |
ForEach-Object { $task[$_.Id] = $_ } 
$task.Count # 52433 => 52677

#------------------------------------------------------------

# Case Notes

$casenote = @{}
Import-Csv "$unzippedRoot\Case_Note__c.csv" -Encoding UTF8 | 
Select-Object Id, Client_Name__c, Name, Carer_Name__c, Case_Worker__c,
@{ n='client'; e={ $contact[$_.Client_Name__c] } },
@{ n='carer'; e={ $contact[$_.Carer_Name__c] } },
@{ n='case_worker'; e={ $contact[$_.Case_Worker__c] } } |
ForEach-Object { $casenote[$_.Id] = $_ } 
$casenote.Count # 52027 => 52406


#------------------------------------------------------------

# The target document library in SharePoint depends on whether the attachment Owner is (or was) in the 'Service Delivery' team

$sd_user = 
Import-Excel "$env:OneDrive\Documents\Salesforce User ID.xlsx" |  # master file from Gracia (emailed 2024-06-19)
Where-Object Role -eq 'Service Delivery' |
Group-Object Id -AsHashTable
$sd_user.count # 25

# for documents, we will use owner's email address as a folder name
$any_user = 
Import-Csv "$unzippedRoot\User.csv" -Encoding UTF8 | 
Group-Object Id -AsHashTable
$any_user.Count # 146
#------------------------------------------------------------

# To ensure each filename is unique, we insert #[AttachmentId] just before the file extension

function unique_fname( $fn, $id, $ext ) {
    $s = [System.IO.Path]::GetFileNameWithoutExtension($fn)
    $s = $s -replace '[:\?]', '' # replace characters not allowed in the name of a WIndows file system object
    $x = [System.IO.Path]::GetExtension($fn)
    if ( $x -eq '' ) { $x = ".$ext" } # for some Documents, the extension is only in the Type property
    $s = $s.subString(0, [System.Math]::Min(120, $s.Length)) # always chop at this stage 
   return ($s.Trim() + ' #' + $id + $x )
}

#------------------------------------------------------------

# To coalesce a list, and also return the position of the first non-blank element

function coalesce( $a ) {
    $i=1
    foreach( $e in $a ) {
        if ( $e -gt '' ) {
            return @{ name=$e; type=$i}
            }
        $i++
        }
    return $null
}

#------------------------------------------------------------
# ONLY use this for folder paths without filenames. It cuts the string to a max length and might remove a file extension
function clean_path( $s ) {
    $s = $s -replace '[:\?]', '' # replace characters not allowed in the name of a Windows file system object
    # PROBLEM: some filestems include '.' - we can't just assume that whatever comes after the last dot is always a file extension
    $s = $s.subString(0, [System.Math]::Min(120, $s.Length)) 
    return ($s.Trim() + $x ) # trim in case it cuts after a space
}

#------------------------------------------------------------

function type_name( $i ) {
    switch ( $i ) {
        8 { return 'Account' }
        9 { return 'Task' }
        10 { return 'Unresolved' }
        default { return 'Contact' }
        }
}

#------------------------------------------------------------

# Now bring it all together to produce a CSV ready for (re-)loading into a hashtable (stage 2).

Import-Csv "$unzippedRoot\Attachment.csv" -Encoding UTF8 |
select *, 
@{ n='noise_level'; e={ $noise_image[$_.ContentType+', '+$_.BodyLength] } } ,
@{ n='doclib'; e={ if ( $sd_user[$_.OwnerId] ) { 'Service Delivery' } else { 'Other' } } },
@{ n='unique_fname'; e={ unique_fname $_.Name $_.Id } }, # shorten BEFORE adding unique ID suffix
@{ n='out_folder'; e= {
    coalesce @(
        $contact[$_.ParentId].folder_name,
        $casenote[$_.ParentId].client.folder_name,
        $task[$_.ParentId].who.folder_name,
        $task[$_.ParentId].client.folder_name,
        $casenote[$_.ParentId].carer.folder_name,
        $casenote[$_.ParentId].case_worker.folder_name,
        $task[$_.ParentId].account.folder_name,
        $account[$_.AccountId].folder_name,
        $task[$_.ParentId].subject, 
        # now for the really persistent orphans
        ('Object Id #'+$_.ParentId), # parentheses required
        ('AccountId #'+$_.AccountId)
         ) } } | 
select Id, ParentId, AccountId,
    doclib, 
    CreatedDate,
    noise_level, 
    @{ n='folder'; e={ clean_path $_.out_folder.name }}, # some folders are named after Task.Subject, which can be way too long
    @{ n='type'; e={ type_name $_.out_folder.type }},  
    unique_fname | 
Export-Csv -NoTypeInformation -Encoding UTF8 -Path "$unzippedRoot\Attachment-Map.csv"

#------------------------------------------------------------

# Before starting or resuming Stage 2, we need to (re-)load our mapping details into a hashtable.
<#
$attachment = @{}
Import-Csv "$unzippedRoot\Attachment-Map.csv" -Encoding UTF8 | 
ForEach-Object { $attachment[$_.Id] = $_ } 


$attachment['00PPr0000057BS5MAM']


#>

Import-Csv "$unzippedRoot\Document.csv" -Encoding UTF8 | 
select *, 
@{ n='doclib'; e={ if ( $sd_user[$_.AuthorId] ) { 'Service Delivery' } else { 'Other' } } },
@{ n='unique_fname'; e={ unique_fname $_.Name $_.Id $_.Type } },
@{ n='author_email'; e= { $any_user[$_.AuthorId].email } }, 
@{ n='out_folder'; e= { "Document`\$($any_user[$_.AuthorId].email)`\Folder #$($_.FolderId)" } } | 
select Id, FolderId, author_email, Type, 
    doclib, 
    CreatedDate,
    @{ n='folder'; e={ clean_path $_.out_folder }}, 
    unique_fname | 
Export-Csv -NoTypeInformation -Encoding UTF8 -Path "$unzippedRoot\Document-Map.csv"

# 
Import-Csv -Encoding UTF8 -Path "$unzippedRoot\Document-Map.csv" | select -First 1