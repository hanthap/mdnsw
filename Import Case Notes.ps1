<#
$f = 'C:\Users\PeterLuckock\Downloads\Case_Note__c.csv'


$indata = Get-Content $f -TotalCount 2 | ConvertFrom-Csv 

$indata

#----------------------------- ##>

    
function Import-Accdb {
param( 
    [parameter(mandatory=$true)] [string] $ViewName,
    [parameter(mandatory=$true)] [string] $Path,
    $SQL = "select * from $ViewName",
    $CountSQL = "select count(*) as N from ( $SQL )",
    $Caption = "Processing $ViewName in $Path..."
)

begin {

    $cn = New-Object System.Data.OleDb.OleDbConnection
    $cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$Path;Mode=Share Deny Write;"
    [void] $cn.Open() # seems to work here inside function, but not in open code

    $cmd = New-Object System.Data.OleDb.OleDbCommand
    $cmd.Connection = $cn

    if ( $CountSQL ) {
        try {
        Write-Host "Checking batch size"
        $cmd.CommandText = $countSQL
        $rs = $cmd.ExecuteReader()
        [void]$rs.Read()
        $nTotalCount = $rs.GetValue(0)
        Write-Host "Total batch count $nTotalCount"
        } catch { }
    }

    $rs.Close()

    $cmd.CommandText = $sql
    $rs = $cmd.ExecuteReader()
    $nDocsProcessed = 0
    Write-Progress -PercentComplete 0 -Activity $caption -Status "Processed $nDocsProcessed of $nTotalCount"

    while ( $rs.Read() ) {
        $ht = @{}
        for ( $i=0; $i -lt $rs.FieldCount; $i++ )
        {
            if ( $rs.GetValue($i) -ne $null ) { $ht.Add( $rs.GetName($i), $rs.GetValue($i) ) }
        } 
        [pscustomobject] $ht

        $nDocsProcessed++
        $pctDone = $nDocsProcessed / $nTotalCount * 100
        Write-Progress -PercentComplete $pctDone -Activity $caption -Status "Processed $nDocsProcessed of $nTotalCount."
        }
    $rs.Close()

    $cn.Close()

    }  # end begin


} # end function


$accdb = 'C:\Users\PeterLuckock\Downloads\test.accdb'

$d = Import-Accdb -ViewName Campaign -Path $accdb 

$d 