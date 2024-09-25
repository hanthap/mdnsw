#--------------------------------------------------------------------

# Create a separate subset of (cleansed) PII columns which we then use for detecting possible duplicates
# To detect more duplicates, it helps to remove irrelevant characters => better signal strength
# $Contact_clean_pii_csv is defined in initialise.ps1

Import-Csv "$unzippedRoot\Contact.csv" | # when checking for duplicates we must always start with raw unredacted data 
    Select-Object Id,
        @{Name='Initials' ;   Expression={ Get-Initials $_.FirstName, $_.LastName } },
        @{Name='FirstName';   Expression={ strim $_.FirstName } },
        @{Name='LastName' ;   Expression={ strim $_.LastName } },
        @{Name='MobilePhone'; Expression={ ntrim $_.MobilePhone } },
        Birthdate,
        Email,
        ONEN_Household__c,
        NDIS_No__c,
        CreatedDate,
        Membership__c |
    ForEach-Object {  
        if ( $_.MobilePhone.Length -gt 9 ) { $_.MobilePhone = $_.MobilePhone.Substring( $_.MobilePhone.Length-9 ) }; # drop ISD prefix
        try { $_.NDIS_No__c = (ntrim($_.NDIS_No__c)).Substring(0,9) } catch { } ; # just in case embedded spaces have circumvented the unique constraint 
        $_ } |
    Export-Csv -Path $Contact_clean_pii_csv -NoTypeInformation # cache this between PS sessions

#--------------------------------------------------------------------

# reload PII from disk cache, EXCEPT records already tagged as Archived (a reasonable assumption)
$Contact_clean_pii = Import-Csv -Path $Contact_clean_pii_csv | Where-Object Membership__c -ne 'Archive'

# now check for plausible duplicates using various criteria 
$dupName   = $Contact_clean_pii | Group LastName,FirstName | Where Count -gt 1
$dupName.Count # 791 => 681 after excluding archived - includes many 'namesakes' (logically distinct people with same name)

# matching by mobile & email detects a few extra pairs that are 'obviously' the same person but have changed surname and/or 'familiar' first name.
# We need to add at least one other field to reduce the noise. (Still this is not immune to the funraisin problem reported by Gracia.)

$dupMobile1 = $Contact_clean_pii | Where MobilePhone -gt '' | Group MobilePhone,FirstName | Where Count -gt 1
$dupMobile1.Count # 55 => 48

$dupMobile2 = $Contact_clean_pii | Where MobilePhone -gt '' | Group MobilePhone,Initials | Where Count -gt 1
$dupMobile2.Count # 95 => 84

$dupEmail1  = $Contact_clean_pii | Where Email -gt '' | Group Email,FirstName | Where Count -gt 1
$dupEmail1.Count # 86 => 84

$dupDOB = $Contact_clean_pii | Where Birthdate -gt '1900' | Group Birthdate,Initials | Where Count -gt 1
$dupDOB.Count # 62 => 52

# there are no duplicate NDIS refs. Probably enforced unique constraint! BUT, one that could be easily fooled by embedded noise eg space characters
#$dupNDIS = $Contact_clean_pii | Where NDIS_No__c -gt '' | Group NDIS_No__c | Where Count -gt 1

<# 
PROBLEM: Gracia reported on 16/9/24 that some seemingly-obvious "duplicates" are spuriously misleading, due to as-yet-unexplained change of name. Found many examples.
Harry suggests maybe an automated data feed (from funraisin?) has been updating Contact records based on {Lastname,Email} and OVERWRITING FirstName (anything else?)
Example: https://mdnsw.lightning.force.com/lightning/r/Contact/00380000012bOj1AAE/view
Description: "9.4.2021 19.04.2021 Fundraisin auto sync BRRS donation from [contact name] into salesforce. Somehow it created new donor "[same name]" and changed [original] client MD record name to anon anon in Mar 2021. 
    GS merge both contact as [original name] as Client MD on 9 Apr 2021"

WE NEED TO BE VERY CAREFUL WHEN DECIDING WHAT GROUPS ARE SAFE TO MERGE

#>

#--------------------------------------------------------------------

<# 
Tag each plausible-looking group with a unique surrogate key (arbitrarily, the ContactId of its newest member)
The 'newest' isn't always the 'best' (eg the one with the largest number of related child records in scope for migration)
To find that Id, we also need the date of their most recent CaseNote,Attachment,Task and/or (non-bounce) Campaign Response 
So the ultimate choice of 'primary host' ('destination' ContactId) has to wait until AFTER all spurious false matches have been flagged manually
Probably best done in Excel with PowerQuery...
#>
 
function GroupId ( $g ) {
     return ( $g | Sort CreatedDate -Descending | Select -First 1 ).Id
 }

#--------------------------------------------------------------------

function Expand-Group { 
    [CmdletBinding()]
    param(
        [Parameter(Mandatory,ValueFromPipeline)] [PSObject] $o, # a reference
        [string] $MatchedBy
        )
process{
    $GroupId = GroupId $_.Group
    $o.Group | % { 
    $_ | select *, @{n='GroupId';e={$GroupId}}, @{n='MatchedBy';e={$MatchedBy}}
    }
    }
}
#----------------------------------------------

$d1 = $dupName    | Expand-Group -MatchedBy 'Name'
$d2 = $dupMobile1 | Expand-Group -MatchedBy 'Mobile'
$d3 = $dupMobile2 | Expand-Group -MatchedBy 'Mobile'
$d4 = $dupDOB     | Expand-Group -MatchedBy 'Birthdate' 
$d5 = $dupEmail1  | Expand-Group -MatchedBy 'Email'

$outcsv = "$unzippedRoot\AllDuplicates.csv" 
$d1 + $d2 + $d3 + $d4 + $d5 | Export-Csv -Path $outcsv -NoTypeInformation

$allDuplicates = Import-Csv $outcsv

$allDuplicates.Count # 2266 => 1958

#--------------------------------------------------------------------

