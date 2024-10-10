


# Create lookup hashtables used as scoping filters

$contact_in_scope = @{}
Import-Csv  $Contact_clean_pii_csv |
Where-Object Membership__c -ne 'Archive'  | # explicit assumption
Select-Object Id, FirstName, LastName |
ForEach-Object { $contact_in_scope[$_.Id] = $_.FirstName + ' ' + $_.LastName }  # TO DO - value could be the post-merge 'primary' Contact Id 
$contact_in_scope.Count # 36404 => 36417 => 36427

# Many ONEN_Household records are obviously duplicates.
# 99% of these are resolved by excluding hholds with no Contacts in scope

$hhold_in_scope = Import-Csv  $Contact_clean_pii_csv | 
Where-Object Membership__c -ne 'Archive'  | # explicit assumption
Select-Object ONEN_Household__c |
Group-Object ONEN_Household__c -AsHashTable
$hhold_in_scope.count # 1313 => 1313


#--------------------------------------------------------------------
# ACCOUNTS

Import-Csv "$unzippedRoot\Account.csv" -Encoding UTF8 | # "Bush Rock Café" 
Where-Object Membership__c -ne 'Archive' |
# Redact-Columns -ColumnNames @( 'BillingStreet', 'ShippingStreet', 'Description', 'Email__c', 'Phone' )  | # Keith asked for Name to be unmasked, 30/9/24
# Update-Properties -PropertyList @( ) -HashTable $contact_ids_to_merge |
Select-Object Id,
RecordTypeId,
Type,
Name,
CreatedDate,
Description,
ParentId,
Area__c,
Collector__c,
CreatedById,
Donations_this_FY__c,
Email__c,
Fax,
First_Donation__c,
Largest_Donation__c,
Membership__c,
Most_Recent_Donation__c,
Most_Recent_Opportunity_Date__c,
npo02__Best_Gift_Year__c,
npo02__Best_Gift_Year_Total__c,
npsp__All_Members_Deceased__c,
npsp__CustomizableRollups_UseSkewMode__c,
npsp__Sustainer__c,
npsp__Undeliverable_Address__c,
Number_of_Donations__c,
NumberOfEmployees,
OwnerId,
Phone,
Total_Donations__c,
Website, 
# BillingStreet, BillingCity, BillingState, BillingPostalCode, BillingCountry, 
# ShippingStreet, ShippingCity, ShippingState, ShippingPostalCode, ShippingCountry,
@{Name='BillingAddress'; Expression={ Clean-Address $_.BillingStreet, $_.BillingCity, $_.BillingState, $_.BillingPostalCode, $_.BillingCountry } }, 
@{Name='ShippingAddress'; Expression={ Clean-Address $_.ShippingStreet, $_.ShippingCity, $_.ShippingState, $_.ShippingPostalCode, $_.ShippingCountry } }  |
Export-Csv -NoTypeInformation -Delimiter '|' -Encoding UTF8 -Path "$unzippedRoot\Account_subset.csv" 

#--------------------------------------------------------------------
# CONTACTS


Import-Csv "$unzippedRoot\Contact.csv" -Encoding UTF8 | # for Contact ID = "0033b00002T0K03AAF", UTF7 reads [MobilePhone] as "㗓蹴�" whereas UTF8 agrees with UI ("+61410450243") 
Where-KeyMatch -KeyName Id -LookupTable $contact_in_scope |
# Redact-Columns -ColumnNames @( 'FirstName', 'LastName', 'OtherStreet', 'MailingStreet', 'Description', 'Email', 'MobilePhone', 'Phone', 'Other_Email__c', 'Spouse_Name__c' )  |
Select-Object Id, # ALL non-trivial columns in scope for migration
Salutation,FirstName,LastName,Email,MobilePhone,Birthdate,CreatedDate,
LastActivityDate,Funraisin__Funraisin_Id__c,NDIS_No__c,ONEN_Household__c,RecordTypeId,
MDANSW_ID__c,
# MailingCity,MailingCountry,MailingPostalCode,MailingState,MailingStreet,
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
# Email_Confirmation__c, # all blank
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
#IsDeleted,
Languages__c,
# LastModifiedById,
# LastModifiedDate,
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
# OtherCity,OtherCountry,OtherPhone,OtherPostalCode,OtherState,OtherStreet,
Other_Email__c,
Other_del__c,
OwnerId,
Philanthropic_Interests__c,
Phone,
Postcode_Lookup__c,
RT_Name__c,
Receive_Information_Pack__c,
# Receive_Newsletter__c, all 0
ReportsToId,
SHARE_DETAILS_TO_OTHER_OR__c,
Safe_Home_Visiting_Date__c,
Safe_Home_Visiting_checklist__c,
Send_EOFY_Receipt__c,
# SFSSDupeCatcher__Override_DupeCatcher__c, all 0
Share_Details_opt_out__c,
Spouse_Name__c,
Startup_Audit__c,
Sugar_Free_Donor__c,
Sugar_Free_Fundraiser__c,
# SystemModstamp,
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
# npsp__CustomizableRollups_UseSkewMode__c, all 0
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
volunteering__c,
@{Name='MailingAddress'; Expression={ Clean-Address $_.MailingStreet, $_.MailingCity, $_.MailingState, $_.MailingPostalCode, $_.MailingCountry } },
@{Name='OtherAddress'  ; Expression={ Clean-Address $_.OtherStreet,   $_.OtherCity,   $_.OtherState,   $_.OtherPostalCode,   $_.OtherCountry   } } | 
ForEach-Object { # a bit of scrubbing 
    $_.NDIS_No__c = ntrim( $_.NDIS_No__c )
    $_
    } |
# ReportsToId is the only Lookup(Contact). There are 11 Contacts in scope with a non-blank ReportsToId
# Update-Properties -HashTable $contact_ids_to_merge -PropertyList @( 'ReportsToId' ) |
Export-Csv -NoTypeInformation -Delimiter '|' -Encoding UTF8 -Path "$unzippedRoot\Contact_raw_subset.csv"

#-------------------------------------------------------------------------
# CAMPAIGN MEMBERS

Import-Csv "$unzippedRoot\CampaignMember.csv" -Encoding UTF8 | 
# Easier in PowerQuery : DO NOT MIGRATE Campaign member record type.Name in ( 'Care for Carers', 'Tour Duchenne' )
# Where-Object Funraisin__Is_Archived__c -ne 'Y' |  
Where-KeyMatch -KeyName ContactId -LookupTable $contact_in_scope |
Select-Object Id,
#IsDeleted,
IsPrimary,
CampaignId,
ContactId,
CreatedById,
CreatedDate,
FirstRespondedDate,
Fundraiser_URL__c,
Funraisin__Fundraising_Target__c,
Funraisin__Funraisin_Id__c,
Funraisin__History_Type__c,
Funraisin__Is_Active__c,
Funraisin__Is_Archived__c,
Funraisin__Is_Paid__c,
Funraisin__Number_Seats__c,
# Funraisin__Seat_Number__c,
# Funraisin__Team__c,
# Handicap__c,
HasResponded,
LastModifiedById,
LastModifiedDate,
# LeadId,
Status,
# SystemModstamp,
# Team_Captain_name__c,
# Team_Name__c,
# Team_URL__c,
wbsendit__Activity_Date__c,
wbsendit__Activity__c,
wbsendit__Clicks__c,
wbsendit__Opens__c,
# What_type_of_team_are_you_registering_as__c,
# Why_did_you_choose_to_enter_this_challen__c,
# Workplace_School_Name__c 
# 27/9/24 MH requested add next 2:
Number_Attending__c, 
How_did_you_hear_about_this_challenge__c |
Update-Properties -PropertyList @( 'ContactId' ) -HashTable $contact_ids_to_merge |
Export-Csv -NoTypeInformation -Delimiter '|' -Encoding UTF8 -Path "$unzippedRoot\CampaignMember_subset.csv"

#-------------------------------------------------------------------------

# CAMPAIGNS

Import-Csv "$unzippedRoot\Campaign.csv" -Encoding UTF8 | 
# Where-Object Funraisin__Is_Archived__c -ne 'Y' |  #TO DO : CRITERIA?
Where-Object { $_.CampaignMemberRecordTypeId -ne '' -or $_.Id -eq '701800000006lLHAAY' } | # workaround: keep 'General Donations' even though it has no CampaignMemberRecordTypeId (27/9/24)
#Redact-Columns -ColumnNames @( 'Description', 'Name' ) |
Select-Object Id, 
Name,
ActualCost,
AmountAllOpportunities,
AmountWonOpportunities,
BudgetedCost,
CampaignMemberRecordTypeId,
# Campaign_External_ID__c,
<#
CICD__Child_Campaign_Order__c
CICD__Date_of_death__c
CICD__Donation_Envelopes__c
CICD__Donation_Image_External_URL__c
CICD__Donation_Image_Id__c
CICD__Donation_Page_Description__c
CICD__Donation_Settings__c
CICD__Donation_Target__c
CICD__Donation_Title__c
CICD__Donation_URL_Name__c
CICD__Fundraiser_GUID__c
CICD__Fundraiser_Photo__c
CICD__Fundraiser__c
CICD__Fundraise_For__c
CICD__Fundraise_Image_External_URL__c
CICD__Fundraise_Image_Id__c
CICD__Fundraise_Message__c
CICD__Fundraise_Page_Description__c
CICD__Fundraise_Promote__c
CICD__Fundraise_Setting__c
CICD__Fundraise_Short_Desc__c
CICD__Fundraise_Target__c
CICD__Fundraise_Thank_You__c
CICD__Fundraise_Title__c
CICD__Fundraise_Type__c
CICD__Fundraise_URL_Name__c
CICD__Fundraising_Donation_Campaign__c
CICD__Fundraising_Team_Photo__c
CICD__Fundraising_Team__c
CICD__HTML_Head__c
CICD__InMemoriam_Tribute_Page__c
CICD__Next_of_Kin__c
CICD__Recurring_Campaign__c
CICD__Send_Fundraiser_Email__c
CICD__Source_Opportunity__c
CICD__Video_URL__c
CICI__Contact__c
CICI__Event__c
CICI__Fundraiser_ID__c
CICM__Membership_Drive_Image_External_URL__c
CICM__Membership_Drive_Image_Id__c
CICM__Membership_Drive_Page_Description__c
CICM__Membership_Drive_Setting__c
CICM__Membership_Drive_Title__c
CICM__Membership_Drive_URL_Name__c
CIC_External_ID__c
#>
CreatedById,
CreatedDate,
Description,
Elapsed_Time__c,
EndDate,
#EveryDayHero_URL__c,
ExpectedResponse,
ExpectedRevenue,
<#
Funraisin__Allow_Entries__c
Funraisin__Allow_Tables__c
Funraisin__Campaign_Type__c
Funraisin__City__c
Funraisin__Country__c
Funraisin__Created_By_Fundraiser__c
Funraisin__Entry_Fee__c
Funraisin__Entry_Limit__c
Funraisin__Entry_Type__c
Funraisin__Event_Target__c
Funraisin__Event_Type__c
Funraisin__Funraisin_Id__c
Funraisin__Is_Fundraising_Event__c
Funraisin__Maximum_Tickets__c
Funraisin__Minimum_Tickets__c
Funraisin__Postcode__c
Funraisin__Seats_Per_Table__c
Funraisin__State__c
Funraisin__Street_Address__c
Funraisin__Table_Price__c
Funraisin__Ticket_Price__c
#>
HierarchyActualCost,
HierarchyAmountAllOpportunities,
HierarchyAmountWonOpportunities,
HierarchyBudgetedCost,
HierarchyExpectedRevenue,
HierarchyNumberOfContacts,
HierarchyNumberOfConvertedLeads,
HierarchyNumberOfLeads,
HierarchyNumberOfOpportunities,
HierarchyNumberOfResponses,
HierarchyNumberOfWonOpportunities,
HierarchyNumberSent,
IsActive,
#IsDeleted
#LastActivityDate
#LastModifiedById
#LastModifiedDate
Membership_Donation_Email_Template__c,
# Membership_renewal_template__c,
New_Member_template__c,
NSWMD_ID__c,
NumberOfContacts,
# NumberOfConvertedLeads,
# NumberOfLeads,
NumberOfOpportunities,
NumberOfResponses,
NumberOfWonOpportunities,
NumberSent,
OwnerId,
ParentId,
RecordTypeId,
Single_Donation_template__c,
StartDate,
Status,
# SystemModstamp,
Type,
wbsendit__Campaign_Monitor_Id__c,
wbsendit__Email_Text_Version__c,
wbsendit__Email_Web_Version__c,
wbsendit__Num_Bounced__c,
wbsendit__Num_Clicks__c,
wbsendit__Num_Forwards__c,
# wbsendit__Num_Likes__c,
# wbsendit__Num_Mentions__c,
wbsendit__Num_Opens__c,
wbsendit__Num_Recipients__c,
wbsendit__Num_Spam_Complaints__c,
wbsendit__Num_Unique_Opens__c,
wbsendit__Num_Unsubscribed__c,
wbsendit__Tags__c,
wbsendit__World_View_Email_Tracking__c |
Export-Csv -NoTypeInformation -Delimiter '|' -Encoding UTF8 -Path "$unzippedRoot\Campaign_subset.csv"

#-------------------------------------------------------------------------
# HOUSEHOLDS

Import-Csv "$unzippedRoot\ONEN_Household__c.csv" -Encoding UTF8 |
Where-KeyMatch -LookupTable $hhold_in_scope | # skip any households that have no current in-scope contacts. (This removes 99% of duplicated households.)
# Redact-Columns -ColumnNames @( 'Name', 'MailingStreet__c', 'Unique_Household__c', 'Recognition_Name_Short__c', 'Recognition_Name__c' ) |
# there are no Lookup(Contact) ids to be remapped
Select-Object *, 
@{Name='MailingAddress'; Expression={ Clean-Address $_.MailingStreet__c, $_.MailingCity__c, $_.MailingState__c, $_.MailingPostalCode__c, $_.MailingCountry__c } } |
Export-Csv -NoTypeInformation -Delimiter '|' -Encoding UTF8 -Path "$unzippedRoot\ONEN_Household__c_subset.csv"


#------------------------------------------------------------------------
# OPPORTUNITIES

# Business rule: exclude RecordTypes with these IDs as specified by Gracia. 

$ExcludeRecordTypeIdList = @(
'012800000003Z82', # Collection Box - MDF
'0128000000022bT', # Donation DO NOT USE
'012800000003Z88', # Grant - MDF
'0123b0000007ygA', # Membership Donation
'0123b0000007ygB'  # Membership Registration
)

Import-Csv "$unzippedRoot\Opportunity.csv" -Encoding UTF8 | 
Where-Object RecordTypeId -NotIn $ExcludeRecordTypeIdList |
Where-KeyMatch -KeyName ContactId -LookupTable $contact_in_scope |
# Redact-Columns -ColumnNames @( 'Name', 'Check_Author__c', 'Description' )  |
Select-Object Id,
Name,
ContactId,
CreatedDate,
RecordTypeId,
Receipt_number__c,
AccountId,
Amount,
Description,
CampaignId,
Check_Author__c,
Check_Bank__c,
Check_Branch__c,
Check_Date__c,
Check_Number__c,
CloseDate,
Contact__c,
CreatedById,
ExpectedRevenue,
ForecastCategory,
Fundraiser_Name__c,  # new target field = Funraisin__Fundraiser__c ; Points to a Contact, so will need adjusting?
Funraisin__Donation_Type__c,
Funraisin__Fundraiser__c, # Points to a Contact, so will need adjusting?
Funraisin__Funraisin_Id__c,
Funraisin__OrderNumber__c,
Funraisin__Payment_Date__c,
Funraisin__Payment_Method__c,
Funraisin__Primary_Contact__c,
Funraisin__Source_Opportunity__c,
Funraisin__TrackingNumber__c,
IsClosed,
IsWon,
LastActivityDate,
LastStageChangeDate,
LeadSource,
OwnerId,
Payment_ID__c,
Payment_Type__c,
Pricebook2Id,
Probability,
StageName,
StageSortOrder,
Status__c,
Tax_Code__c,
Type |
# skip these as they're all empty
# npsp__Batch_Number__c, 
# npsp__CommitmentId__c,
# npsp__Grant_Period_End_Date__c,
# npsp__Grant_Period_Start_Date__c,
# npsp__Grant_Program_Area_s__c,
# npsp__Grant_Requirements_Website__c, 
# npsp__Primary_Contact__c,
# npsp__Requested_Amount__c 
Update-Properties -HashTable $contact_ids_to_merge -PropertyList @( 'ContactId', 'Contact__c', 'Fundraiser_Name__c', 'Funraisin__Fundraiser__c', 'Funraisin__Primary_Contact__c' ) | 
Export-Csv -Delimiter '|' -NoTypeInformation -Encoding UTF8 -Path "$unzippedRoot\Opportunity_subset.csv" 

#-------------------------------------------------------------------------------
# CASE NOTES

Import-Csv "$unzippedRoot\Case_Note__c.csv"  -Encoding UTF8 | 
Where-Object LastModifiedDate -ge '2022' | # Only migrate case notes from 2022-2024 (Jess, 30/9/24)
Where-KeyMatch -KeyName Client_Name__c -LookupTable $contact_in_scope |
# Redact-Columns -ColumnNames @( 'Name' ) |  #  Keith asked for Action_c to be unmasked 30/9/24
Select-Object Id,
Name,
Action__c, # do not mask this
Carer_Name__c, #  Lookup(Contact) 
Case_Worker__c, # Lookup(User)
Client_Name__c, #  Lookup(Contact)
CreatedById,
CreatedDate,
Date__c,
Elapsed_Time__c,
# IsDeleted,
# LastActivityDate,
#SystemModstamp 
LastModifiedById,
LastModifiedDate,
Location_of_service_delivery__c,
NDIS_Billable__c,
#  These 2 HTML columns are imported directly to accdb. (Access can't import csv records longer than 32k)
# Action_Detail__c,
# Notes__c, 
OwnerId |
# Currently no (in scope) case notes are related to merged contacts, but just in case...
Update-Properties -PropertyList @( 'Client_Name__c', 'Carer_Name__c' ) -HashTable $contact_ids_to_merge | 
Export-Csv -Delimiter '|' -NoTypeInformation -Encoding UTF8 -Path "$unzippedRoot\Case_Note__c_subset.csv"


#-------------------------------------------------------------------------------

#TASKS

# Where-KeyMatch only supports a single KeyName. We use '+' to concatenate 2 filtered resultsets. Therefore need to dedupe later.
( Import-Csv "$unzippedRoot\Task.csv" -Encoding UTF8 | Where-Object IsArchived -ne 1 | Where-KeyMatch -KeyName Client_Name__c -LookupTable $contact_in_scope ) + 
( Import-Csv "$unzippedRoot\Task.csv" -Encoding UTF8 | Where-Object IsArchived -ne 1 | Where-KeyMatch -KeyName WhoId -LookupTable $contact_in_scope ) |
# Redact-Columns -ColumnNames @( 'Subject', 'Description' )  | # Keith wants unmasked data so he can assign to the correct sub-type in new org
Select-Object Id,
Client_Name__c, # Lookup(Contact)
# Type, # all blank (Some were 'Email' but they all have IsArchived=1)
WhoId, # Lookup(Contact or User)
Subject,
Description, # long text including some html tags and non-visible bytes - better save directly to accdb, Excel doesn't export the whole text.
Action__c,  #Non-billable etc
OwnerId,
Priority,  #High / Normal
AccountId,
ActivityDate,
ActivityOriginType, # 1,2, 5
CallDisposition, # free text name
# CallDurationInSeconds,  0
# CallObject,
# CallType,
CompletedDateTime,
CreatedById,
CreatedDate,
Elapsed_Time__c,
EmailMessageId, # points to another table not in scope!
# IsArchived, Since we're filtering on this column
IsClosed,
# IsDeleted,
# IsRecurrence,
IsReminderSet,
#  IsVisibleInSelfService,
LastModifiedById,
LastModifiedDate,
# npsp__Engagement_Plan_Task__c,
# npsp__Engagement_Plan__c,
# Number__c,
# RecurrenceActivityId,
# RecurrenceDayOfMonth,
# RecurrenceDayOfWeekMask,
# RecurrenceEndDateOnly,
# RecurrenceInstance,
# RecurrenceInterval,
# RecurrenceMonthOfYear,
# RecurrenceRegeneratedType,
# RecurrenceStartDateOnly,
# RecurrenceTimeZoneSidKey,
# RecurrenceType,
ReminderDateTime,
Status,
#SystemModstamp,
# Time__c,
# wbsendit__Smart_Email_Id__c,
# wbsendit__Smart_Email_Message_Id__c,
# wbsendit__Smart_Email_Recipient__c,
# wbsendit__Smart_Email_Status__c,
WhatCount,
WhatId, # Lookup(*anything*)
WhoCount |
Update-Properties -PropertyList @( 'Client_Name__c', 'WhoId' ) -HashTable $contact_ids_to_merge |
Export-Csv -Delimiter '|' -NoTypeInformation -Encoding UTF8  -Path "$unzippedRoot\Task_subset.csv"


#-------------------------------------------------------------------------------

# TASK RELATION
Import-Csv -Encoding UTF8 -Path "$unzippedRoot\TaskRelation.csv"  | 
Where-Object LastModifiedDate -gt '2022' |
Where-KeyMatch -KeyName RelationId -LookupTable $contact_in_scope | 
Update-Properties -PropertyList @('RelationId') -HashTable $contact_ids_to_merge | 
Export-Csv -NoTypeInformation -Encoding UTF8 -Path "$unzippedRoot\TaskRelation_subset.csv"
