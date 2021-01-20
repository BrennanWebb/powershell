param(
    [string]$server = "hulk",
    [string]$db = "adhocdata",
    [string]$file = "d:\humana\powershell\output\test.csv",
    [string]$sql = "select getdate() loadtime, * from adhocdata.sys.schemas ;"
)
$conn="data source=$server;initial catalog=$db;integrated security=true;"
invoke-sqlcmd -query $sql -connectionstring $conn | export-csv -path $file -notypeinformation