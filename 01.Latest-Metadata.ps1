<#

You only need to run this code after a fresh set of backup ZIP files from Salesforce Weekly Archive has been copied to Staging doclib in SharePoint.

BEFORE using Access to ingest latest csv files and generate a fresh master lookup qryAttachmentScope.csv

(All joins etc are done in Access.)

#>

$shell = New-Object -Com Shell.Application


$zip = $shell.NameSpace( $zipFileStem + '1.zip' )
    $flist = $zip.Items() # forces download sync if not already cached
    $shell.Namespace($unzippedRoot).CopyHere($flist)

# Some raw csv files have too many columns for Access to handle. We need to extract a subset of columns

Import-Csv -Path "$downloads\Contact.csv" | 
    #select -First 100 | 
        Select-Object Id,Salutation,FirstName,LastName,RecordTypeId,OwnerId,CreatedDate |
            Export-Csv -Path "$downloads\Contact_out.csv" -NoTypeInformation

Import-Csv -Path "$downloads\Case_Note__c.csv" | 
    # select -First 100 | 
        Select-Object Id,OwnerId,Name,CreatedById,Client_Name__c,Carer_Name__c,Case_Worker__c |
            Export-Csv -Path "$downloads\Case_Note.csv" -NoTypeInformation

# now go to Access and repopulate all the tables.. Then export qryAttachmentScope as csv