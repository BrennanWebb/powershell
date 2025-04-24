If ($host.Version.Major -lt 5) {
    Write-Host "This script requires version 5 or higher.  See Humana App Shop for Powershell 5 or higher install.";
    Exit;
};

[array]$module ="ImportExcel","SqlServer"
foreach ($item in Get-InstalledModule) {
    $name=$item.name;
    if ($name -in $module) {
        Write-Host "Module "$name" exists.";
    } 
    else {
        Write-Host "Module "$name" does not exist.  Install starting.";
        Install-module $Module -scope currentuser -Force;
    } ;  
};

#Select Data from Server and write results to window (host)
$Conn ='data source="HULK";initial catalog="ADHOCDATA";integrated security=true;';
$sql ='SELECT * From AdhocData.PSLAO.DPM_ROSTER_GROUP;';

$psTable=invoke-sqlcmd -query $sql -connectionstring $Conn;
$psTable |format-table -AutoSize| out-string|Write-Host


