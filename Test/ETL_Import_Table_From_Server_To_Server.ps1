
<#Overview:
The following code will copy results from server table to another server table.
The $drop variable will drop the table before load.
Powershell 5.0, 7.1.4, 7.2.5 tested.
Required Modules: Install-Module -Name SqlServer -Scope CurrentUser -Force -ErrorAction Stop;
#>


#Source Vars
$SourceServer ="warehouse.selectquote.com";
$SourceDb="test_Sandbox";
$SourceSchema ="dbo";
$SourceTable ="CallDetail_Restate_DeleteAnyTime"

#Target Vars
$TargetServer ="warehouse-UAT.selectquote.com";
$TargetDb="test_Sandbox";
$TargetSchema ="dbo";
$TargetTable ="CallDetail_Restate_DeleteAnyTime"
$TargetAction = "Drop" #Drop, Trunc

$TopN = 0  # 0 means all
function ReadWrite {
    if ($TopN -gt 0) {
        $Query = "SELECT TOP($TopN) * FROM [$SourceSchema].[$SourceTable]"
        $Data = Invoke-Sqlcmd -Query $Query -ServerInstance $SourceServer -Database $SourceDb
    } else {
        $Data = Read-SqlTableData -ServerInstance $SourceServer -Database $SourceDb -SchemaName $SourceSchema -TableName $SourceTable
    }
    $Data | Write-SqlTableData -ServerInstance $TargetServer -Database $TargetDb -SchemaName $TargetSchema -TableName $TargetTable
}

try {
    Switch($TargetAction)
    {
        Drop {
            $Query = "DROP TABLE IF EXISTS [$TargetSchema].[$TargetTable]"
            Invoke-Sqlcmd -Query $Query -ServerInstance $TargetServer -Database $TargetDb
            ReadWrite
        }
        Trunc {
            $Query = "TRUNCATE TABLE [$TargetSchema].[$TargetTable]"
            Invoke-Sqlcmd -Query $Query -ServerInstance $TargetServer -Database $TargetDb
            ReadWrite
        }
    }
}
catch {
    Write-Error "Error occurred: $_"
}



