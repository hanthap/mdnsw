
# for the unzipped files before renaming & moving
$unzippedRoot = "$env:USERPROFILE\Downloads"

$f = "$unzippedRoot\ONEN_Household__c.csv" 

$indata = Get-Content $f -TotalCount 30 | ConvertFrom-Csv 


$indata


# load the map into session memory
$scope_household_id = Get-Content $f -TotalCount 6 | 
    ConvertFrom-Csv | 
        Select-Object Id,Name |
            Group -AsHashTable -Property Id  # retain in memory as a dictionary

#-------------------------------------------------------------------------

function mask( $s ) {
return $s -replace "(\w)(\w)", "X`${2}"
}

#-------------------------------------------------------------------------
function Filter-HashKeys { 


    [CmdletBinding()]
    param(
      [Parameter(Mandatory, ValueFromPipeline)] [PSObject] $obj,
      [string] $Property = 'ContactId',
      [hashtable] $HashTable = @{}, # pipe the object if its [$Property] is found as a key in $HashTable
      $RedactColumns = @()
      )

process {

    if ( $HashTable.Count -eq 0 -or  $HashTable.ContainsKey($obj.$Property) ) { 
        # scramble PII fields
        foreach( $c in $RedactColumns ) {
            $obj.$c = mask $obj.$c
            }

    $obj # send down the pipe

        }
} # end process

} # end function



#-------------------------------------------------------------------------



$f = "$unzippedRoot\ONEN_Household__c.csv" 


$outcsv = "$unzippedRoot\ONEN_Household__c_redacted.csv" 

$PII = @( 'Name', 'MailingStreet__c', 'Unique_Household__c', 'Recognition_Name_Short__c', 'Recognition_Name__c' ) 

Get-Content $f -TotalCount 10000 | 
    ConvertFrom-Csv | 
        Where-Object -Property SystemModstamp -ge -Value '2000-01-01' |
        Filter-HashKeys -RedactColumns $PII |  
            Export-Csv -Path $outcsv -NoTypeInformation


#-------------------------------------------------------------------------


$f = "$unzippedRoot\Account.csv" 

$outcsv = "$unzippedRoot\Account_redacted.csv" 

$PII = @( 'Name', 'BillingStreet', 'ShippingStreet', 'Description', 'Email__c' ) 

Get-Content $f -TotalCount 100000 | 
    ConvertFrom-Csv | 
        Where-Object -Property SystemModstamp -ge -Value '2000-01-01' |
        Filter-HashKeys -RedactColumns $PII |  
            Export-Csv -Path $outcsv -NoTypeInformation

#-------------------------------------------------------------------------

$f = "$unzippedRoot\Contact.csv" 

$outcsv = "$unzippedRoot\Contact_redacted.csv" 

$PII = @( 'FirstName', 'LastName', 'OtherStreet', 'MailingStreet', 'Phone', 'Fax', 
    'MobilePhone', 'HomePhone', 'OtherPhone', 'Email', 'Description', 'Comment__c', 
    'Spouse_Name__c', 'Addressee__c', 'Given_Name__c', 'Emergency_Contact__c', 'Other_Email__c' ) 

Get-Content $f -TotalCount 100000 | 
    ConvertFrom-Csv | 
        Where-Object -Property SystemModstamp -ge -Value '2000-01-01' |
        Filter-HashKeys -RedactColumns $PII |  
            Export-Csv -Path $outcsv -NoTypeInformation

#-------------------------------------------------------------------------

$f = "$unzippedRoot\Opportunity.csv" 

$outcsv = "$unzippedRoot\Opportunity_redacted.csv" 

$PII = @( 'Name', 'Check_Author__c', 'Description' ) 

Get-Content $f -TotalCount 100000 | 
    ConvertFrom-Csv | 
        Where-Object -Property SystemModstamp -ge -Value '2000-01-01' |
        Filter-HashKeys -RedactColumns $PII |  
            Export-Csv -Path $outcsv -NoTypeInformation
            

mask( 'Keith Peter Harry Milvia' )