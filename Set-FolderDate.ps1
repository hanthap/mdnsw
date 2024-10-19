

function Set-FolderDate { 

    [CmdletBinding()]
    param(
      [Parameter(Mandatory, ValueFromPipeline)] [PSObject] $f
      )

process {
    $ts = $f | Get-ChildItem | Sort LastWriteTime | Select-Object -Last 1 LastWriteTime
    $f.LastWriteTime = $ts 
    } 

}

$d = Get-ChildItem "$env:OneDrive\Other\_NOISE_\Contact" | group Parent | fl | % { $_.Group | sort LastWriteTime | select -last 1 } | select LastWriteTime, FullName | fl

$d[0].Parent.FullName


$d | % { $_.Group } | select FullName, LastWriteTime | 