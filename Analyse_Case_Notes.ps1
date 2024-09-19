$raw_csv = "$unzippedRoot\Case_Note__c.csv" 

Get-Content $raw_csv -TotalCount 3 | ConvertFrom-Csv

$clean_csv = "$unzippedRoot\Case_Note__c_subset.csv" 

Import-Csv $raw_csv |
    Select-Object Client_Name__c, Id,OwnerId,CreatedDate,Carer_Name__c, Case_Worker__c, Name |
    Export-Csv -Path $clean_csv -NoTypeInformation # save this between PS sessions

$indata = Import-Csv -Path $clean_csv # reload from disk

# now check for potential duplicates 


$latest_casenote_csv = "$unzippedRoot\LatestCaseNotePerClientName.csv" 

$indata | Group Client_Name__c | Foreach-Object { $Count=$_.Count; $_.Group | Sort-Object CreatedDate | Select-Object *, @{n='Count';e={$Count}} -Last 1 } | 
    Export-Csv -NoTypeInformation -Path $latest_casenote_csv

$LatestCaseNotePerClientName = Import-Csv -Path $latest_casenote_csv

$LatestCaseNotePerClientName.Count # 1047

$LatestCaseNotePerClientName | select -first 20 | ft

#------------------------------------------------------------------------------------------------------------


$raw_csv = "$unzippedRoot\Attachment.csv" 

Get-Content $raw_csv -TotalCount 3 | ConvertFrom-Csv

$clean_csv = "$unzippedRoot\Attachment_subset.csv" 

Import-Csv $raw_csv |
    Select-Object ParentId, Id,OwnerId,AccountId,CreatedDate,Name |
    Export-Csv -Path $clean_csv -NoTypeInformation # save this between PS sessions

$indata = Import-Csv -Path $clean_csv # reload from disk


$latest_attachment_csv = "$unzippedRoot\LatestAttachmentPerParent.csv" 

$indata | Group ParentId | Foreach-Object { $Count=$_.Count; $_.Group | Sort-Object CreatedDate | Select-Object *, @{n='Count';e={$Count}} -Last 1 } | 
    Export-Csv -NoTypeInformation -Path $latest_attachment_csv

$LatestAttachmentPerParent = Import-Csv -Path $latest_attachment_csv

$LatestAttachmentPerParent.Count # 8088


#------------------------------------------------------------------------------------------------------------


$raw_csv = "$unzippedRoot\Opportunity.csv" 

Get-Content $raw_csv -TotalCount 3 | ConvertFrom-Csv

$clean_csv = "$unzippedRoot\Opportunity_subset.csv" 

Import-Csv $raw_csv |
    Select-Object ContactId,Id,CampaignId,AccountId,Payment_Id__c,MDNSW_PaymentID__c,CreatedDate,Type,Name,Amount |
    Export-Csv -Path $clean_csv -NoTypeInformation # save this between PS sessions

$indata = Import-Csv -Path $clean_csv # reload from disk


$latest_opportunity_csv = "$unzippedRoot\LatestOpportunityPerContact.csv" 

$indata | Group ContactId | Foreach-Object { $Count=$_.Count; $_.Group | Sort-Object CreatedDate | Select-Object *, @{n='Count';e={$Count}} -Last 1 } | 
    Export-Csv -NoTypeInformation -Path $latest_opportunity_csv

$LatestOpportunityPerContact = Import-Csv -Path $latest_opportunity_csv

$LatestOpportunityPerContact.Count # 28275


$LatestOpportunityPerContact | select -first 30 | ft

#------------------------------------------------------------------------------------------------------------


$raw_csv = "$unzippedRoot\Task.csv" 

Get-Content $raw_csv -TotalCount 3 | ConvertFrom-Csv


$clean_csv = "$unzippedRoot\Task_subset.csv" 

Import-Csv $raw_csv |
    Select-Object Client_Name__c, WhoId, Id,OwnerId,AccountId,CreatedDate,Subject |
    Export-Csv -Path $clean_csv -NoTypeInformation # save this between PS sessions

$indata = Import-Csv -Path $clean_csv # reload from disk


$latest_task_who_csv = "$unzippedRoot\LatestTaskPerWho.csv" 

$indata | Group WhoId | Foreach-Object { $Count=$_.Count; $_.Group | Sort-Object CreatedDate | Select-Object *, @{n='Count';e={$Count}} -Last 1 } | 
    Export-Csv -NoTypeInformation -Path $latest_task_who_csv

$LatestTaskPerWho = Import-Csv -Path $latest_task_who_csv

$LatestTaskPerWho.Count # N=6332


$latest_task_client_csv = "$unzippedRoot\LatestTaskPerClient.csv" 

$indata | Group Client_Name__c | Foreach-Object { $Count=$_.Count; $_.Group | Sort-Object CreatedDate | Select-Object *, @{n='Count';e={$Count}} -Last 1 } | 
    Export-Csv -NoTypeInformation -Path $latest_task_client_csv

$LatestTaskPerClient = Import-Csv -Path $latest_task_client_csv

$LatestTaskPerClient.Count # N=81


#------------------------------------------------------------------------------------------------------------


$raw_csv = "$unzippedRoot\CampaignMember.csv" 

Get-Content $raw_csv -TotalCount 3 | ConvertFrom-Csv

$clean_csv = "$unzippedRoot\CampaignMember_subset.csv" 

Import-Csv $raw_csv |
    Select-Object ContactId,Id,CampaignId,CreatedDate,FirstRespondedDate,Status |
    Export-Csv -Path $clean_csv -NoTypeInformation # save this between PS sessions

$indata = Import-Csv -Path $clean_csv # reload from disk

$latest_CampaignMemberCreated_csv = "$unzippedRoot\LatestCampaignMemberCreatedPerContact.csv" 

$indata | Group ContactId | Foreach-Object { $Count=$_.Count; $_.Group | Sort-Object CreatedDate | Select-Object *, @{n='Count';e={$Count}} -Last 1 } | 
    Export-Csv -NoTypeInformation -Path $latest_CampaignMemberCreated_csv

$LatestCampaignMemberCreatedPerContact = Import-Csv -Path $latest_CampaignMemberCreated_csv

$LatestCampaignMemberCreatedPerContact.Count # N=24517


$latest_CampaignMemberResponded_csv = "$unzippedRoot\LatestCampaignMemberRespondedPerContact.csv" 

$indata | Where-Object FirstRespondedDate -gt '' | Group ContactId | Foreach-Object { $Count=$_.Count; $_.Group | Sort-Object FirstRespondedDate | Select-Object *, @{n='Count';e={$Count}} -Last 1 } | 
    Export-Csv -NoTypeInformation -Path $latest_CampaignMemberResponded_csv

$LatestCampaignMemberRespondedPerContact = Import-Csv -Path $latest_CampaignMemberResponded_csv

$LatestCampaignMemberRespondedPerContact.Count # N=4205


