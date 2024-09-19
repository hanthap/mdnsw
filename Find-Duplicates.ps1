
$Contact_raw_csv = "$unzippedRoot\Contact.csv" 
$Contact_raw_subset_csv = "$unzippedRoot\Contact_raw_subset.csv" 

Get-Content $Contact_raw_csv -TotalCount 2 | ConvertFrom-Csv


Import-Csv $Contact_raw_csv | 
Select-Object Id, # list ALL columns in scope for migration
Salutation,FirstName,LastName,Email,MobilePhone,Birthdate,CreatedDate,
LastActivityDate,Funraisin__Funraisin_Id__c,NDIS_No__c,ONEN_Household__c,RecordTypeId,
MDANSW_ID__c,
MailingCity,
MailingCountry,
MailingPostalCode,
MailingState,
MailingStreet,
Aboriginal__c,
AccountId,
Addressee__c,
Age_Picklist__c,
Allergies__c,
Are_you_a_family_member_carer__c,
Bequest_Prospect__c,
Big_Red_Roll_Stroll__c,
Case_Worker_to_make_contact__c,
Christmas_Appeal_Donor__c,
Christmas_Appeal_Only__c,
Comment__c,
Communication__c,
Consents__c,
Consultation_Prospect__c,
CreatedById,
Cultural_Background__c,
Current_Board_Member__c,
Current_NDIS_Plan_End_Date__c,
Current_Plan_Date__c,
Current_Regular_Donor__c,
Date_Joined__c,
Deceased_Date__c,
Deceased__c,
Department,
Description,
Director__c,
DoNotCall,
Do_Not_Mail__c,
Do_Not_Survey__c,
Do_not_want_membership_renewal_pack__c,
Email_Confirmation__c,
Emergency_Contact__c,
Fax,
Funraisin__Is_Donor__c,
Funraisin__Is_Fundraiser__c,
Funraisin__Optin__c,
Gender__c,
Gifts__c,
Given_Name__c,
HasOptedOutOfEmail,
HasOptedOutOfFax,
Head_Of_Household__c,
Health__c,
HomePhone,
Household_Member_Deceased__c,
Household_Member_has_MD__c,
How_did_you_hear_about_us__c,
Interested_in_receiving_information_abou__c,
Interests__c,
IsDeleted,
Languages__c,
LastModifiedById,
LastModifiedDate,
LeadSource,
MD_Type_Comment__c,
MD_Type__c,
Mailing_State__c,
Major_Donor__c,
March_Appeal_Donor__c,
Marital_Status__c,
Membership__c,
Mobility_requirements__c,
Motor_vehicle_registration__c,
NDIS_Audit_Opt_Out_Date__c,
NDIS_Audit_Opt_Out__c,
Newsletter_Communication__c,
OtherCity,
OtherCountry,
OtherPhone,
OtherPostalCode,
OtherState,
OtherStreet,
Other_Email__c,
Other_del__c,
OwnerId,
Philanthropic_Interests__c,
Phone,
Postcode_Lookup__c,
RT_Name__c,
Receive_Information_Pack__c,
Receive_Newsletter__c,
ReportsToId,
SHARE_DETAILS_TO_OTHER_OR__c,
Safe_Home_Visiting_Date__c,
Safe_Home_Visiting_checklist__c,
Send_EOFY_Receipt__c,
SFSSDupeCatcher__Override_DupeCatcher__c,
Share_Details_opt_out__c,
Spouse_Name__c,
Startup_Audit__c,
Sugar_Free_Donor__c,
Sugar_Free_Fundraiser__c,
SystemModstamp,
Tax_Appeal_Donor__c,
Tax_Appeal_Only__c,
Text_Opt_Out__c,
Title,
Type_of_Member__c,
Volunteer_Interests__c,
WWCC_No__c,
WWCC_Valid_To_Date__c,
WWCC_Verified_Date__c,
Workplace_School_Name__c,
fundraising_and_marketing__c,
npe01__PreferredPhone__c,
npe01__Preferred_Email__c,
npo02__Best_Gift_Year_Total__c,
npo02__Best_Gift_Year__c,
npo02__Soft_Credit_Last_Year__c,
npo02__Soft_Credit_This_Year__c,
npo02__Soft_Credit_Total__c,
npo02__Soft_Credit_Two_Years_Ago__c,
npsp__CustomizableRollups_UseSkewMode__c,
npsp__First_Soft_Credit_Amount__c,
npsp__First_Soft_Credit_Date__c,
npsp__Largest_Soft_Credit_Amount__c,
npsp__Largest_Soft_Credit_Date__c,
npsp__Last_Soft_Credit_Amount__c,
npsp__Last_Soft_Credit_Date__c,
npsp__Number_of_Soft_Credits_Last_N_Days__c,
npsp__Number_of_Soft_Credits_Last_Year__c,
npsp__Number_of_Soft_Credits_This_Year__c,
npsp__Number_of_Soft_Credits_Two_Years_Ago__c,
npsp__Number_of_Soft_Credits__c,
npsp__Soft_Credit_Last_N_Days__c,
receive_Talking_Point__c,
volunteering__c |
Export-Csv -Path $Contact_raw_subset_csv -NoTypeInformation # cache this between PS sessions

Get-Content $Contact_raw_subset_csv -TotalCount 5 | ConvertFrom-Csv

#--------------------------------------------------------------------

function ntrim( $s ) {
try { 
return $s -replace '\D', '' 
} catch { return '' }
}

function strim( $s ) {
try {
return $s.ToUpper()  -replace '[^A-Z]', '' 
} catch { return '' }
}
<#
#--------------------------------------------------------------------
function Append-Score() {
    [CmdletBinding()]

    param(
        [Parameter(Mandatory,ValueFromPipeline)] [PSObject] $o,
        [hashtable] $coeff = @{}
        )

  process{
    $n= 0
      $x = $o.PSObject.Properties | Where Value -eq  '' | % {
        $n += $coeff[$_.Name]
        }
      $o | Add-Member 'GapSize' $n -Force
      $o
  }
}

#--------------------------------------------------------------------

# weighting coefficients for relevance score, 

$w = @{ 
    Birthdate =8;
    FirstName =8;
    MobilePhone =4;
    NDIS_No__c =4;
    MD_Type__c = 4;
    WWCC_No__c = 4;
    Email =4;
    MailingStreet =4
    ONEN_Household__c =2;
    Phone =1;
}
#>

#--------------------------------------------------------------------

# A subset of CLEANSED PII columns that we will use for detecting possible duplicates
# Sorting (eg by last activity date) will be done later in PowerQuery after joining with other csv datasets and removing 'Archived' Contact records

$Contact_raw_csv = "$unzippedRoot\Contact.csv" 
$Contact_clean_subset_csv = "$unzippedRoot\Contact_clean_subset.csv" 

Import-Csv -LiteralPath $Contact_raw_csv | select -first 10

Import-Csv $Contact_raw_csv |
    Select-Object Id,
        FirstName,
        LastName,
        DepartmentGroup,
        Birthdate,
        Email,
        MobilePhone,
        ONEN_Household__c,
        CreatedDate | 
    ForEach-Object {  # improve signal strength in some group-by fields
    $_.FirstName=strim($_.FirstName);
    $_.LastName=strim($_.LastName);
    try { $_.DepartmentGroup=$_.FirstName.SubString(0,1)+$_.LastName.SubString(0,1) } catch { };  # re-purposed to store initials
    $_.MobilePhone=ntrim($_.MobilePhone); 
    if ( $_.MobilePhone.Length -gt 9 ) { $_.MobilePhone = $_.MobilePhone.Substring( $_.MobilePhone.Length-9 ) }; # drop ISD prefix
    $_ } |
    Export-Csv -Path $Contact_clean_subset_csv -NoTypeInformation # cache this between PS sessions

#--------------------------------------------------------------------

$indata = Import-Csv -Path $Contact_clean_subset_csv # reload from disk cache

# now check for potential duplicates 
$dupName   = $indata | Group LastName,FirstName | Where Count -gt 1
$dupName.Count # 791 - includes many 'namesakes' (logically distinct people with same name)

# matching by mobile & email detects a few extra pairs that are 'obviously' the same person (but use distinct maiden? surname and/or 'familiar' first name) 
# We need to add at least one other field to reduce the noise. (Still this is not immune to the problem reported by Gracia.)

$dupMobile1 = $indata | Where MobilePhone -gt '' | Group MobilePhone,FirstName | Where Count -gt 1
$dupMobile1.Count # 55

# NOTE: the result is a structure of POINTERS. So everything is done by reference. Adding a property (for example) can have spooky side-effects

$dupMobile2 = $indata | Where MobilePhone -gt '' | Group MobilePhone,DepartmentGroup | Where Count -gt 1
$dupMobile2.Count # 95

$dupEmail1  = $indata | Where Email -gt '' | Group Email,FirstName | Where Count -gt 1
$dupEmail1.Count # 86

$dupDOB = $indata | Where Birthdate -gt '1900' | Group Birthdate,DepartmentGroup | Where Count -gt 1
$dupDOB.Count # 62


# MAJOR PROBLEM: Gracia reported on 16/9/24 that some "seemingly-obvious" duplicates are spuriously misleading, due to as-yet-unexplained change of name. Many examples.
# Harry suggests an automated data update (from funraisin?) may have been updating Contact record based on {Lastname,Email} and OVERWRITING FirstName (anything else?)
# THIS MEANS WE NEED TO BE VERY CAREFUL WHEN DECIDING WHAT IS SAFE TO MERGE

# there are no duplicate NDIS refs. Must be a unique constraint!
# $dupNDIS = $indata | Where NDIS_No__c -gt '' | Group NDIS_No__c | Where Count -gt 1

#--------------------------------------------------------------------

# Tag each group with a 'default ID' (as the ID of its most recently-created Contact)
# This is an 'interim' ID since the 'newest' isn't always the 'best' (most recently ACTIVE)
# To find that, we also need the date of their most recent CaseNote,Attachment,Task and/or (non-bounce) Campaign Response 
# Therefore the ultimate choice of destination ContactId has to wait until AFTER spurious false matches have been flagged manually
# i.e. it has to be done with PowerQuery & Excel, not PowerShell.

 
function GroupId ( $g ) {
     return ( $g | Sort CreatedDate -Descending | Select -First 1 ).Id
 }

#--------------------------------------------------------------------


# HOW TO copy the referenced input, then add a property to the copy and send THAT down the pipe?
function Expand-Group { 
    [CmdletBinding()]
    param(
        [Parameter(Mandatory,ValueFromPipeline)] [PSObject] $o, # a reference
        [string] $MatchedBy
        )
process{
    $GroupId = GroupId $_.Group
    $o.Group | % { 
    $_ | select *, @{n='GroupId';e={$GroupId}}, @{n='MatchedBy';e={$MatchedBy}} # don't use Add-Member!
    }
    }
}
#----------------------------------------------


$d1 = $dupName    | Expand-Group -MatchedBy 'Name'
$d2 = $dupMobile1 | Expand-Group -MatchedBy 'Mobile'
$d3 = $dupMobile2 | Expand-Group -MatchedBy 'Mobile'
$d4 = $dupDOB     | Expand-Group -MatchedBy 'Birthdate' 
$d5 = $dupEmail1  | Expand-Group -MatchedBy 'Email'

<#
$d3[1] | select GroupId,Id,FirstName,LastName,Email,MobilePhone,MatchedBy -First 10 | ft

$dupName | select -First 10


$dupName[1].Group

$d3 | where GroupId -eq '0033b00002gJqcNAAS'
#>

$outcsv = "$unzippedRoot\AllDuplicates.csv" 
$d1 + $d2 + $d3 + $d4 + $d5 | Export-Csv -Path $outcsv -NoTypeInformation

$allDuplicates = Import-Csv $outcsv
#$allDuplicates | where GroupId -eq '0033b00002SZTpJAAX' |select Id,FirstName,LastName,Email,MobilePhone,MatchedBy -First 30 | ft

$allDuplicates.Count # 2266

#--------------------------------------------------------------------

