<# 

PRECONDITIONS

Wait for notification that maica__Client_Note__c has been (re-)populated (so that new Ids are available)

Login to workbench (new org) and run this SOQL query: 

    SELECT Id,Legacy_Case_Note_ID__c FROM maica__Client_Note__c

Then download the results as csv
Rename fresh csv file as "maica__Client_Note__c_map.csv"

#>

# load the new org ID mapping into a hashtable
$maica__Client_Note__c_map = @{}
Import-Csv "$unzippedRoot\maica__Client_Note__c_map.csv" | where Legacy_Case_Note_ID__c -gt '' | 
ForEach-Object { $maica__Client_Note__c_map[$_.Legacy_Case_Note_ID__c] = $_.Id }

$maica__Client_Note__c_map.Count # 46789 => 15455



#---------------------------------------------------------------------------------------------
<#
    We can bypass Access accdb and export a csv directly from PowerShell, for updating via Data Loader. 
#>


function Trim-Html {

    [CmdletBinding()]
    param(
      [Parameter(Mandatory, ValueFromPipeline)] [PSObject] $obj,
      $PropertyList=@('Notes__c'),
      $MaxLength =  131072 , # increased from 32768
      [string] $ReplaceWithText = '<H3 style="color: red">[Embedded image removed due to 131 kB character limit]</H3>',
      [string] $TruncatedWarning = '<H3 style="color: red">[End of text not included due to 131 kB character limit]</H3>',
      [System.Text.RegularExpressions.RegexOptions] $RegexOptions = ([System.Text.RegularExpressions.RegexOptions]::RightToLeft),
      [switch] $ToPlainText = $false
      )

begin {


    $rxFooter = New-Object -TypeName regex -ArgumentList '\<img alt="email footer".*\<\/img\>', ($RegexOptions)
    $rxAny = New-Object -TypeName regex -ArgumentList '\<img.*\<\/img\>', ($RegexOptions)
    [regex] $rxRubbish = '�| style=""'

}

process {


    foreach( $PropName in $PropertyList ) {

            $s = $obj.$PropName 

            if ( $s.Length -gt 1 ) {
                $s = $s.Substring(1) # skip zero byte
                #remove all "email footer" images with no warning text
                [void]$s.Replace( '<span.*></span>', '' ) # junk span tags
                $s = $rxFooter.Replace( $s, '' ) 
                $s = $rxRubbish.Replace( $s, '' ) 

                if ( $ToPlainText ) {
                    $s = $s -replace '<br>', "`n" # we want to preserve linefeeds
                    $s = $s -replace '<[^>]*>', '' # remove all element tags
                    $s = [System.Web.HttpUtility]::HtmlDecode($s) # replace &amp; &gt; etc.
                    }
    
                #  insert warning message in place of the last (up to N remaining) images
                for ( $i=0; $i -lt 10 -and $s.Length -gt $MaxLength; $i++ ) {
                    #remove all footer images 
                     $s = $rxAny.Replace( $s, $ReplaceWithText, 1 ) 
                }

                if ( $s.Length -gt $MaxLength ) { # still too big
                    $s = ' ' + $TruncatedWarning + $s.Substring(1) # extra space at start?
                    $s = $s.Substring(0,$MaxLength)
                }
            }

            $obj.$PropName = $s
        }


    $obj

    }

}


#---------------------------------------------------------------------------------------------


function Append-Accdb {

param( 
    [parameter(mandatory=$true)] [string] $Path,
    [Parameter(ValueFromPipeline)] [psobject] $o
)

begin {

    $cn = New-Object System.Data.OleDb.OleDbConnection
    $cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$Path;Mode=Share Deny None;"
    [void] $cn.Open() 

    $cmd = New-Object System.Data.OleDb.OleDbCommand
    $cmd.Connection = $cn

    $p = $cmd.Parameters
    foreach( $i in 1..3 ) { 
        [void]$p.Add( ( New-Object System.Data.OleDb.OleDbParameter )  )
        }
    
    $sql = "insert into Case_Note__c ( Id, Notes__c, Action_Detail__c ) values ( ?, ?, ? )"
    $cmd.CommandText = $sql

}

process {
    $p[0].Value = $o.Id
    [void] $cmd.ExecuteNonQuery()  #returns number of rows
}

end {
    $cn.Close()
    }

}

#---------------------------------------------------------------------------------------------

<# DEPRECATED - Keith still has trouble exporting long text from Access to csv - so accdb doesn't help
Import-Csv "$unzippedRoot\Case_Note__c.csv" -Encoding UTF8 | 
Where-Object LastModifiedDate -ge '2022' | # Only migrate case notes from 2022-2024 (Jess, 30/9/24)
Where-KeyMatch -KeyName Client_Name__c -LookupTable $contact_in_scope |
#Select-Object -First 1000 |
Trim-Html -PropertyList @( 'Notes__c', 'Action_Detail__c' ) |
Append-Accdb -Path "$unzippedRoot\Case_Note__c.accdb"
#>

#---------------------------------------------------------------------------------------------

# Instead we send Vertic a pre-mapped "update" csv dataset ready for SF Data Loader

# PRECONDITIONS
$contact_in_scope.count # 36427
$maica__Client_Note__c_map.count # 15455

Import-Csv "$unzippedRoot\Case_Note__c.csv" -Encoding UTF8 | 
Where-Object LastModifiedDate -ge '2022' | 
Where-KeyMatch -KeyName Client_Name__c -LookupTable $contact_in_scope |
Select-Object *,
    @{ n='Legacy_Case_Note_ID__c'; e={ $_.Id } }, 
    @{ n='maica__Client_Note__c.Id'; e={ $maica__Client_Note__c_map[$_.Id] } } |
ForEach-Object { # keep multi-picklist values just for posterity
    if ( $_.Action_Detail__c -gt '' ) { $_.Action__c += '<br>' }
    $_.Action_Detail__c = " Original type: $($_.Action__c)", $_.Action_Detail__c | Out-String
    $_
    } |
Trim-Html -PropertyList @( 'Notes__c', 'Action_Detail__c' ) |
# Trim-Html -PropertyList @( 'Action_Detail__c' ) -MaxLength 32768 -ToPlainText | # Keith has enabled HTML & increased the character limit 15/10/24
select maica__Client_Note__c.Id, Notes__c, Action_Detail__c |
Export-Csv -NoTypeInformation -Encoding UTF8 -Path "$unzippedRoot\maica__Client_Note__c_RTF.csv"

# Import-Csv -Encoding UTF8 -Path "$unzippedRoot\maica__Client_Note__c_RTF.csv" | where { $_.Action_Detail__c.length -gt 200 } |  select -first 10 | fl 'maica__Client_Note__c.Id', Action_Detail__c | fl

