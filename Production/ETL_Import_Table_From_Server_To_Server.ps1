
<#Overview:
The following code will copy results from server table to another server table.
The $drop variable will drop the table before load.
Powershell 5.0, 7.1.4, 7.2.5 tested.
Required Modules: Install-Module -Name SqlServer;
#>


#Source Vars
$SourceServer ="warehouse.selectquote.com";
$SourceDb="Commissions_SQS";
$SourceSchema ="dbo";
$SourceTable ="COMP_CoreCompReportingMonths"

#Target Vars
$TargetServer ="warehouse-UAT.selectquote.com";
$TargetDb="Commissions_SQS";
$TargetSchema ="dbo";
$TargetTable ="COMP_CoreCompReportingMonths"
$TargetAction = "Trunc" #Drop, Trunc

function ReadWrite {
    ,(Read-SqlTableData -ServerInstance $SourceServer -Database $SourceDb -SchemaName $SourceSchema -TableName $SourceTable) |
    Write-SqlTableData -ServerInstance $TargetServer -Database $TargetDb -SchemaName $TargetSchema -TableName $TargetTable;  
}

Switch($TargetAction)
{
    Drop {
        $Query = 'Drop Table If Exists '+$TargetSchema+'.'+$TargetTable
        Invoke-Sqlcmd -query $Query -ServerInstance $TargetServer -database $TargetDb;
        ReadWrite;
    }

    Trunc{
        $Query = 'Truncate Table '+$TargetSchema+'.'+$TargetTable
        Invoke-Sqlcmd -query $Query -ServerInstance $TargetServer -database $TargetDb;
        ReadWrite;
    }
    
};



