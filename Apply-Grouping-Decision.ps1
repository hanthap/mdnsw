# Install-Module -Name ImportExcel -Scope CurrentUser

# Export Gracia's decisions into a csv then import that to a lookup hashtable



$d = Import-Excel -Path "$env:OneDrive\Gracia 2024-09-26.xlsx" -WorksheetName Group | 
Where-Object 'Merge?' -eq Combine;
$d.Count # 93

Import-Excel -Path "$env:OneDrive\Gracia 2024-09-26.xlsx" -WorksheetName Group | 
Where-Object 'Merge?' -eq Combine |
ForEach-Object { $a = $_.Id -split '\n'; $v = $a[0]; $a.Trim() } |
ForEach-Object { [pscustomobject] @{ Key = $_; Value = $v } } |
Export-Csv -NoTypeInformation -Path "$unzippedRoot\Merge_ContactId.csv" # also for loading to PowerQuery

# load from csv into a lookup hashtable for use with cmdlets Update-Properties, Where-KeyMatch
$contact_ids_to_merge = @{}
Import-Csv "$unzippedRoot\Merge_ContactId.csv" |
ForEach-Object { $contact_ids_to_merge[$_.Key] = $_.Value }  
$contact_ids_to_merge.Count # 160 - 177 - 192

#-----------------------------------------------------------------------------------------------------------------

function Update-Properties {

    [CmdletBinding()]
    param(
      [Parameter(Mandatory, ValueFromPipeline)] [PSObject] $obj,
      $PropertyList=@('ContactId'), 
      $HashTable = @{},
      [string] $Prefix = 'Original'
      )

begin {
    $bFirst = $true
}

process {

    if ( $bFirst ) { # ensure the backup columns are exported to csv
        foreach( $KeyName in $PropertyList ) {
            Add-Member -InputObject $obj -NotePropertyName "$Prefix`_$KeyName" -NotePropertyValue $null -Force 
            }
        $bFirst = $false
        }

    foreach( $KeyName in $PropertyList ) {
        if ( $HashTable[$obj.$KeyName].Count -gt 0 -and $HashTable[$obj.$KeyName] -ne $obj.$KeyName ) {
            Add-Member -InputObject $obj -NotePropertyName "$Prefix`_$KeyName" -NotePropertyValue $obj.$KeyName -Force # keep the original value
            $obj.$KeyName = $HashTable[$obj.$KeyName] # overwrite the 'live' value
            }
        }

    $obj

    }

}

#--------------------------------------------------------------------------------------------------------------

Import-Csv "$unzippedRoot\Opportunity.csv" | select | 
Update-Properties -PropertyList @( 'npsp__Primary_Contact__c', 'Contact__c', 'ContactId', 'Fundraiser_Name__c' ) -HashTable $contact_ids_to_merge |
Where-Object Original_ContactId -gt '' |
Export-Csv -NoTypeInformation "$unzippedRoot\Opportunity_adjusted.csv" 

Import-Csv  "$unzippedRoot\Opportunity_adjusted.csv" | 
select ContactId, Original_ContactId, Fundraiser_Name__c, Original_Fundraiser_Name__c | 
Where-Object Original_ContactId -gt ''


#--------------------------------------------------------------------------------------------------------------


Import-Csv -Delimiter '|' -Path "$unzippedRoot\Contact_raw_subset.csv" | 
Where-KeyMatch -LookupTable $contact_ids_to_merge -KeyName Id |
Export-Csv -Delimiter '|' -NoTypeInformation -Path "$unzippedRoot\Contacts_to_merge.csv" -Encoding UTF8


$d =  Import-Csv -Delimiter '|' -Path "$unzippedRoot\Contacts_to_merge.csv" 

$d.Count

$d | select -First 1

