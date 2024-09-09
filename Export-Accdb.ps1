
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

    $parm1 = New-Object System.Data.OleDb.OleDbParameter
    $parm2 = New-Object System.Data.OleDb.OleDbParameter
    $parm3 = New-Object System.Data.OleDb.OleDbParameter

    [void] $cmd.Parameters.Add( $parm1 )
    [void] $cmd.Parameters.Add( $parm2 )
    [void] $cmd.Parameters.Add( $parm3 )

    $sql = "insert into Case_Note__c ( Id, Notes__c, Action_Detail__c ) values ( ?, ?, ? )"
    $cmd.CommandText = $sql

}

process {

    $parm1.Value = $o.Id
    # skip null bytes
    try { $parm2.Value = $o.Notes__c.Substring(1) } catch { $parm2.Value = '' } 
    try { $parm3.Value = $o.Action_Detail__c.Substring(1) } catch { $parm3.Value = '' } 

    [void] $cmd.ExecuteNonQuery()  #returns number of rows

}

end {
    $cn.Close()
    }

}

$csvPath = 'C:\Users\PeterLuckock\Downloads\Case_Note__c.csv'
$accdbPath = 'C:\Users\PeterLuckock\Downloads\test.accdb'

Get-Content $csvPath -TotalCount 1000 | 
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