#---------------------------------------------------------------------------------------------


function Trim-Html {

    [CmdletBinding()]
    param(
      [Parameter(Mandatory, ValueFromPipeline)] [PSObject] $obj,
      $PropertyList=@('Notes__c'),
      $MaxLength = 32768,
      [string] $ReplaceWithText = '<H3 style="color: red">[Embedded image removed due to Maica 32k character limit]</H3>',
      [string] $TruncatedWarning = '<H3 style="color: red">[End of text not included due to Maica 32k character limit]</H3>',
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
                #$s = $s.Substring(1) # skip zero byte
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
    # skip null byte at start, Access inteprets as 'end of string'
    try { $p[1].Value = $o.Notes__c.Substring(1) } catch { $p[1].Value = '' } 
    try { $p[2].Value = $o.Action_Detail__c.Substring(1) } catch { $p[2].Value = '' } 
    [void] $cmd.ExecuteNonQuery()  #returns number of rows
}

end {
    $cn.Close()
    }

}

#---------------------------------------------------------------------------------------------


Import-Csv "$unzippedRoot\Case_Note__c.csv" -Encoding UTF8 | # NOT UTF7!
Where-Object LastModifiedDate -ge '2022' | # Only migrate case notes from 2022-2024 (Jess, 30/9/24)
Where-KeyMatch -KeyName Client_Name__c -LookupTable $contact_in_scope |
Select-Object -First 1000 |
Trim-Html -PropertyList @( 'Notes__c', 'Action_Detail__c' ) |
Append-Accdb -Path "$unzippedRoot\Case_Note__c.accdb"

