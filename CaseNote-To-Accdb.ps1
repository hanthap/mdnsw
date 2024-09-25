
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

Import-Csv "$unzippedRoot\Case_Note__c.csv" | 
Where-KeyMatch -KeyName Client_Name__c -LookupTable $contact_in_scope |
Select-Object -First 1000 |
# surgically excise large embedded graphic only if it's an email footer. (Other embedded graphics should be preserved in case they're important documents or photos.)
ForEach { $_.Notes__c = $_.Notes__c -replace '\<img alt="email footer".*\<\/img\>', '<img alt="email footer">FOOTER IMAGE REMOVED</img>'; $_ } |
#Redact-Columns -ColumnNames @( 'Notes__c',  'Action_Detail__c' )  | # this would mess up the HTML tags
Append-Accdb -Path "$unzippedRoot\Case_Note__c.accdb"

