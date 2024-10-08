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
Import-Csv "$unzippedRoot\Attachment.csv" -Encoding UTF7 | 
Where-Object ContentType -Like 'image*' |  #this filter makes the Group run much faster, no idea why.
Group-Object ContentType, BodyLength -NoElement | 
Where-Object Count -gt 1 | 
ForEach-Object { $noise_image[$_.Name] = $_.Count } # add to hashtable

#------------------------------------------------------------

# from Contacts.csv, populate a mapping hashtable with key = id, value = destination folder name : "SURNAME, Firstname #Id" 

$TextInfo = (Get-Culture).TextInfo # for Title Case
$contact = @{}
Import-Csv "$unzippedRoot\Contact.csv" |
# Append ID so as to avoid combining namesakes "SMITH, John"
Select-Object Id, FirstName, LastName, @{n='folder_name'; e={ $_.LastName.ToUpper()+', '+$TextInfo.ToTitleCase($_.FirstName.ToLower())+' #'+$_.id }} |
ForEach-Object { $contact[$_.Id] = $_ } 
$contact.Count # 39564

#------------------------------------------------------------

# Same for Accounts

$account = @{}
Import-Csv "$unzippedRoot\Account.csv" | 
Select-Object Id, Name, @{ n='folder_name'; e={ $_.Name.ToUpper()+' #'+$_.Id }} |
ForEach-Object { $account[$_.Id] = $_ } 
$account.Count # 4390

#------------------------------------------------------------

# Tasks

$task = @{}
Import-Csv "$unzippedRoot\Task.csv" | 
Select-Object Id, WhoId, AccountId,  Client_Name__c, Subject, 
@{ n='who'; e={ $contact[$_.WhoId] } }, 
@{ n='client'; e={ $contact[$_.Client_Name__c] } }, 
@{ n='account'; e={ $account[$_.AccountId] } } |
ForEach-Object { $task[$_.Id] = $_ } 
$task.Count # 52433

#------------------------------------------------------------

# Case Notes

$casenote = @{}
Import-Csv "$unzippedRoot\Case_Note__c.csv" | 
Select-Object Id, Client_Name__c, Name, 
@{ n='client'; e={ $contact[$_.Client_Name__c] } } |
ForEach-Object { $casenote[$_.Id] = $_ } 
$casenote.Count # 52027

#------------------------------------------------------------

# The target document library in SharePoint depends on whether the attachment Owner is (or was) in the 'Service Delivery' team

$sd_user = 
Import-Excel "$env:OneDrive\Documents\Salesforce User ID.xlsx" |  # master file from Gracia (emailed 2024-06-19)
Where-Object Role -eq 'Service Delivery' |
Group-Object Id -AsHashTable

#------------------------------------------------------------

# To ensure each filename is unique, we insert #[AttachmentId] just before the file extension

function unique_fname( $fn, $id ) {
    $s = [System.IO.Path]::GetFileNameWithoutExtension($fn) 
    $x = [System.IO.Path]::GetExtension($fn)
   return ($s + ' #' + $id + $x )
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

# Now bring it all together to produce a CSV ready for (re-)loading into a hashtable (stage 2).

Import-Csv "$unzippedRoot\Attachment.csv" -Encoding UTF7 |
select *, 
@{ n='noise_level'; e={ $noise_image[$_.ContentType+', '+$_.BodyLength] } } ,
@{ n='doclib'; e={ if ( $sd_user[$_.OwnerId] ) { 'Service Delivery' } else { 'Other' } } },
@{ n='unique_fname'; e={ unique_fname $_.Name $_.Id } }, 
@{ n='out_folder'; e= {
    coalesce @(
        $contact[$_.ParentId].folder_name,
        $casenote[$_.ParentId].client.folder_name,
        $task[$_.ParentId].who.folder_name,
        $task[$_.ParentId].client.folder_name,
        $task[$_.ParentId].account.folder_name,
        $account[$_.AccountId].folder_name,
        ('? ParentId #'+$_.ParentId), # parentheses required
        ('? AccountId #'+$_.AccountId)
         ) } } | 
select Id, ParentId, AccountId,
    doclib, 
    CreatedDate,
    noise_level, 
    @{ n='folder'; e={$_.out_folder.name}}, 
    @{ n='type'; e={$_.out_folder.type}}, 
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
