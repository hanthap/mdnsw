



function ntrim( $s ) {
return $s -replace '\D', '' 
}


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

#--------------------------------------------------------------------

$raw_csv = "$unzippedRoot\Contact.csv" 

Get-Content $raw_csv -TotalCount 2 | ConvertFrom-Csv


$clean_csv = "$unzippedRoot\Contact_clean_subset.csv" 

Get-Content $raw_csv |# -TotalCount 10 | 
              ConvertFrom-Csv |  
              Select-Object Id,
                  LastName,FirstName,
                  Birthdate,
                  Email,
                  MobilePhone,Phone,
                  AccountId,ONEN_Household__c,
                  CreatedDate,LastModifiedDate,SystemModstamp,
                  npsp__Last_Soft_Credit_Date__c,npsp__Last_Soft_Credit_Amount__c,
                  MDANSW_ID__c,New_MDANSW_ID__c,NDIS_No__c,WWCC_No__c,
                  MailingStreet,MailingCity,MailingPostalCode,
                  OtherStreet,OtherCity,OtherPostalCode,
                  Head_Of_Household__c,MD_Type__c,Household_Member_has_MD__c,
                  Funraisin__Is_Donor__c,Funraisin__Is_Fundraiser__c,Major_Donor__c,
                  Funraisin__Funraisin_Id__c,RecordTypeId, DepartmentGroup | 
              ForEach-Object { $_.MobilePhone=ntrim($_.MobilePhone); $_.Phone=ntrim($_.Phone); $_.NDIS_No__c=ntrim($_.NDIS_No__c); $_.DepartmentGroup=$_.LastName.SubString(0,1)+$_.FirstName.SubString(0,1); $_ } |
              Export-Csv -Path $clean_csv -NoTypeInformation # save this between PS sessions

#--------------------------------------------------------------------

$indata = Import-Csv -Path $clean_csv # reload from disk

# now check for potential duplicates 
# most are obvious 
$dupName   = $indata | Group LastName,FirstName | Where Count -gt 1

# most cases of same email are distinct persons. 
$dupEmail1  = $indata | Where Email -gt '' | Group Email,FirstName | Where Count -gt 1

# mobile number reveals a few extra pairs that are logically the same person (but use distinct maiden surname and/or common shortening of first name ) 

$dupMobile1 = $indata | Where MobilePhone -gt '040001' | Group MobilePhone,FirstName | Where Count -gt 1
$dupMobile2 = $indata | Where MobilePhone -gt '040001' | Group MobilePhone,DepartmentGroup | Where Count -gt 1



$dupDOB = $indata | Where Birthdate -gt '1900' | Group Birthdate,DepartmentGroup | Where Count -gt 1


# there are no duplicate NDIS refs. Must be a unique constraint!
# $dupNDIS = $indata | Where NDIS_No__c -gt '' | Group NDIS_No__c | Where Count -gt 1

#--------------------------------------------------------------------

# per group, tag with a 'default best candidate ID' 
# this could be further improved by looking at volume of recent casenotes
 
function BestId ( $g ) {
     return ( $g | Sort NDIS_No__c,CreatedDate -Descending | Select -First 1 ).Id
 }

#--------------------------------------------------------------------

$outcsv = "$unzippedRoot\DupName.csv" 

  
$dupName | % { 
    $BestId = BestId $_.Group
    $_.Group | % {
        $_ | Add-Member 'MergeWithId' $BestId -Force
        $_
        }
     } |
    Export-Csv -Path $outcsv -NoTypeInformation


$d = $indata | Select-Object -First 3

$d.Count

#--------------------------------------------------------------------


$outcsv = "$unzippedRoot\DupMobile1.csv" 

  
$dupMobile1 | % { 
    $BestId = BestId $_.Group
    $_.Group | % {
        $_ | Add-Member 'MergeWithId' $BestId -Force
        $_
        }
     } |
    Export-Csv -Path $outcsv -NoTypeInformation



#--------------------------------------------------------------------

$outcsv = "$unzippedRoot\DupMobile2.csv" 
  
$dupMobile2 | % { 
    $BestId = BestId $_.Group
    $_.Group | % {
        $_ | Add-Member 'MergeWithId' $BestId -Force
        $_
        }
     } |
    Export-Csv -Path $outcsv -NoTypeInformation


#--------------------------------------------------------------------

$outcsv = "$unzippedRoot\DupEmail1.csv" 

  
$dupEmail1 | % { 
    $BestId = BestId $_.Group
    $_.Group | % {
        $_ | Add-Member 'MergeWithId' $BestId -Force
        $_
        }
     } |
    Export-Csv -Path $outcsv -NoTypeInformation


#--------------------------------------------------------------------

$outcsv = "$unzippedRoot\DupDOB.csv" 

  
$dupDOB | % { 
    $BestId = BestId $_.Group
    $_.Group | % {
        $_ | Add-Member 'MergeWithId' $BestId -Force
        $_
        }
     } |
    Export-Csv -Path $outcsv -NoTypeInformation