
#Trunc source
Invoke-Sqlcmd -query "Truncate Table [SelectCARE-SQAH].dbo.awf_session_timers" `
-ServerInstance "warehouse-uat.selectquote.com" `
-database "SelectCARE-SQAH";

#read from source and write to target
Invoke-Sqlcmd -query "select Agent_id, [start_date], [end_date] from [SelectCARE_SQAH].dbo.awf_session_timers WITH(NOLOCK) order by 1 desc " `
-ServerInstance "AH14SC-SCDBDT1.sqis-corp.com\test" `
-database "SelectCARE_SQAH" `
-OutputAs DataSet |
Write-SqlTableData  -ServerInstance "warehouse-uat.selectquote.com" `
-Database "SelectCare-SQAH" `
-SchemaName "dbo" `
-TableName "awf_session_timers" `
-passthru;


