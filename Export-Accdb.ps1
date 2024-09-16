
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
    # skip null bytes
    try { $p[1].Value = $o.Notes__c.Substring(1) } catch { $p[1].Value = '' } 
    try { $p[2].Value = $o.Action_Detail__c.Substring(1) } catch { $p[2].Value = '' } 

    [void] $cmd.ExecuteNonQuery()  #returns number of rows

}

end {
    $cn.Close()
    }

}

$csvPath = 'C:\Users\PeterLuckock\Downloads\Case_Note__c.csv'
$accdbPath = 'C:\Users\PeterLuckock\Downloads\test.accdb'

Get-Content $csvPath -TotalCount 10 | 
    ConvertFrom-Csv | 
        Append-Accdb -Path $accdbPath

<#

$fields = {
“Date” = “06/19/2014”;
“Name” = “Nicolas1847”
“Spiceworks” = “Rulez”
}


process {
foreach ($field in $fields.keys)
{
$recordset.Fields.item($field) = $($fields.$field)
}
}


end {

}

#>


$cmd = New-Object System.Data.OleDb.OleDbCommand

foreach( $i in 1..3 ) { 

[void]$cmd.Parameters.Add( ( New-Object System.Data.OleDb.OleDbParameter )  )

 }


 $cmd.Parameters[0] = "avc"