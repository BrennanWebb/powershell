#check if a version is above a certain level.
If ($host.Version.Major -lt 5) {
    Write-Host "This script requires version 5 or higher.  See Humana App Shop for Powershell 5 or higher install.";
    Exit;
};

#Find what modules are installed
Get-InstalledModule|format-table -AutoSize|out-string|Write-Host

#if modules are not installed, install them. Else update them.
[array]$module = "ImportExcel","SqlServer";
foreach ($item in $module){
    If ($item -in (Get-InstalledModule).name){
        write-host "module "$item" exist";
    }
    else {
        install-module $item -scope CurrentUser -force;
        write-host "module "$item" installed";
    }
};

#Uninstall all modules installed by PS.  Caution!
foreach ($item in (Get-InstalledModule).name) {
    Write-host "Uninstalling " $item;
    UnInstall-module $Item -Force; 
    Write-host "Uninstall Complete";
};

#Find functions and cmdlets in a module
Get-Command -Module ImportExcel

#get help for modules, cmdlets
Get-Help Export-Excel

#uninstall a module
Uninstall-Module ImportExcel

#open webpage
start-process https://github.com/dfinke/ImportExcel/tree/master/Examples

#Select Data from Server and write results to window (host)
$Conn ='data source="HULK";initial catalog="ADHOCDATA";integrated security=true;';
$sql ='SELECT * From AdhocData.PSLAO.DPM_ROSTER_GROUP;';
$psTable=invoke-sqlcmd -query $sql -connectionstring $Conn;
$psTable |format-table -AutoSize| out-string|Write-Host;

#export SQL data set to excel table
$pstable | Export-Excel -ExcludeProperty ItemArray, RowError, RowState, Table, HasErrors; 

#insert records from table or excel file to sql server (server write permissions must exist)
$servertable = 'DPM_PS_TEST'
$sql = 'Drop Table '+ $Servertable
invoke-sqlcmd -query $sql -connectionstring $Conn; #make sure target object does not exist.
$psTable | Write-SqlTableData -ServerInstance 'HULK' -DatabaseName 'AdhocData' -SchemaName 'PSLAO' -TableName $servertable -Force -Timeout 0;

