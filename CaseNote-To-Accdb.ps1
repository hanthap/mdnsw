<# 

PRECONDITIONS

In workbench run this SOQL query: 

SELECT Id,Legacy_Case_Note_ID__c FROM maica__Client_Note__c

Then download result as csv
Rename csv as maica__Client_Note__c_map.csv

#>

# load the new org ID mapping into a hashtable
$maica__Client_Note__c_map = @{}
Import-Csv "$unzippedRoot\maica__Client_Note__c_map.csv" | where Legacy_Case_Note_ID__c -gt '' | 
ForEach-Object { $maica__Client_Note__c_map[$_.Legacy_Case_Note_ID__c] = $_.Id }

$maica__Client_Note__c_map.Count # 46789



#---------------------------------------------------------------------------------------------
<#
    Update as of 4/10/24: Data Loader can import long text > 32kB but Data Import Wizard truncates regardless.  Keith says that's not a problem with Data Loader.
    I've asked Keith to try importing a single row including "style-conscious" HTML just to find out whether the fonts & colours are rendered in browser or (b) removed during import (which is what we see so far).
    10/10/24 Harry says limit has been increased to 131072 bytes
   
#>


function Trim-Html {

    [CmdletBinding()]
    param(
      [Parameter(Mandatory, ValueFromPipeline)] [PSObject] $obj,
      $PropertyList=@('Notes__c'),
      $MaxLength =  131072 , # increased from 32768
      [string] $ReplaceWithText = '<H3 style="color: red">[Embedded image removed due to 131 kB character limit]</H3>',
      [string] $TruncatedWarning = '<H3 style="color: red">[End of text not included due to 131 kB character limit]</H3>',
      [System.Text.RegularExpressions.RegexOptions] $RegexOptions = ([System.Text.RegularExpressions.RegexOptions]::RightToLeft)
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

Import-Csv "$unzippedRoot\Case_Note__c.csv" -Encoding UTF8 | 
Where-Object LastModifiedDate -ge '2022' | 
Where-KeyMatch -KeyName Client_Name__c -LookupTable $contact_in_scope |
Select-Object -First 1000 *,
    @{ n='Legacy_Case_Note_ID__c'; e={ $_.Id } }, 
    @{ n='maica__Client_Note__c.Id'; e={ $maica__Client_Note__c_map[$_.Id] } } |
Trim-Html -PropertyList @( 'Notes__c', 'Action_Detail__c' ) |
ForEach-Object { # keep multi-picklist values just for posterity
    if ( $_.Action_Detail__c -gt '' ) { $_.Action__c += '<br>' }
    $_.Action_Detail__c = "Original type: $($_.Action__c)", $_.Action_Detail__c | Out-String
    $_
    } |
select maica__Client_Note__c.Id, Legacy_Case_Note_ID__c, Notes__c, Action_Detail__c |
Export-Csv -NoTypeInformation -Encoding UTF8 -Path "$unzippedRoot\maica__Client_Note__c_RTF.csv"

