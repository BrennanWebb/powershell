Function ETL_to_Server {
    #Variables
    $FileAttr                     = Get-Item $IntakeFilePath
    $FileExt                      = $FileAttr.Extension; Write-Host 'FileExt: '$FileExt;
    $FileDir                      = $FileAttr.DirectoryName; Write-Host 'FileDir:'$FileDir;
    $FileDirName                  = [System.IO.Path]::GetFileName($FileDir); Write-Host 'FileDirName:'$FileDirName;
    $FileName                     = ($FileAttr.Name -replace $FileExt,''); Write-Host 'FileName:'$FileName;
    $TempIntakeFilePath           =$TempDirPath+$FileAttr.Name; Write-Host 'TempIntakeFilePath:'$TempIntakeFilePath;
    $FileModifyTimeStamp          = $FileAttr.LastWriteTime; Write-Host 'FileModifyTimeStamp:'$FileModifyTimeStamp;
    $FileModifyTimeStampFormatted = ($FileAttr.LastWriteTime -f 'YYYY-MM-DD HH:MM:SS'-replace '[^0-9]', ''); Write-Host 'FileModifyTimeStampFormatted:'$FileModifyTimeStampFormatted;
    #Get-ItemProperty $IntakeFilePath |  Format-list -Property * -Force

    if (Test-Path $TempIntakeFilePath) {
        switch -exact ($FileExt){
            '.csv'{
                #write-host ('.csv source file')
                $Table=($TablePrefix+'_'+$FileName) -replace ' ', ''; #write-host $Table;
                $SheetName='' #Blank Sheet name for CSV
                drop_table $Table;
                ,(import-csv $TempIntakeFilePath) | Write-SqlTableData -ServerInstance $server -DatabaseName $db -SchemaName $schema -TableName $Table -Force -Timeout 600;
                alter_table $Table;
                update_table $Table;
                $Table=$null;
            }
            {$_ -eq '.xls' -or $_ -eq '.xlsx'} {
                #write-host ('.xls, .xlsx source file')
                #check for .xls file type.  If true convert to .xlsx
                if ($_ -eq'.xls'){
                    ConvertTo-ExcelXlsx -Path $TempIntakeFilePath -Force
                    
                    #Set original .xls file back to original modify timestamp
                    $FileAttr.LastWriteTime = $FileModifyTimeStamp

                    #replace extension with .xlsx and set new file modify timestamp to original timestamp.
                    $TempIntakeFilePath=$TempIntakeFilePath -replace '.xls','.xlsx';
                    $FileAttr               = Get-Item $TempIntakeFilePath 
                    $FileAttr.LastWriteTime = $FileModifyTimeStamp
                    
                }

                
                $Table=($TablePrefix+'_'+$FileName+'_'+$SheetName) -replace ' ', ''; #write-host $Table;
                drop_table $Table;
                ,(Import-Excel $TempIntakeFilePath -WorksheetName $SheetName -AsText *) 
                | Write-SqlTableData -ServerInstance $server -DatabaseName $db -SchemaName $schema -TableName $Table -Force -Timeout 600;
                alter_table $Table;
                update_table $Table;
                #$Table=$null;
                
                Remove-Item $TempIntakeFilePath;
                update_intake_table; 
            };
        };
    };
};

function drop_table {
    param ($Table);$sql = $null
    $sql = 'Drop table [adhocdata].[pslao].['+ $Table +'];'; #write-host $sql;
    invoke-sqlcmd -query $sql -connectionstring $Conn;
    $sql = $null
};

function alter_table {
    param ($Table);$sql = $null
    $sql = 'Alter table [adhocdata].[pslao].['+$Table+']
            Add [IntakeFilePath] varchar(500)
               ,[SheetName]      varchar(25 )
               ,[SessionID]      varchar(25)
               ,[User]           varchar(25)
               ,[DPMID]          varchar(25)
               ,[Type]           varchar(25)
               ,[Segment]        varchar(25)
               ,[LoadTime]       datetime;'; #write-host $sql;
    invoke-sqlcmd -query $sql -connectionstring $Conn;
    $sql = $null
};

function update_table {
    param ($Table);$sql = $null
    $sql = 'Update [adhocdata].[pslao].['+$Table+']
            Set [IntakeFilePath] ='+$Tick+$IntakeFilePath+$Tick+'
               ,[SheetName]      ='+$Tick+$SheetName+$Tick+'
               ,[SessionID]      ='+$Tick+$SessionID+$Tick+'
               ,[User]           ='+$Tick+$User+$Tick+'
               ,[DPMID]          ='+$Tick+$DPMID+$Tick+'
               ,[Type]           ='+$Tick+$Type+$Tick+'
               ,[Segment]        ='+$Tick+$Segment+$Tick+'
               ,[LoadTime]       =GetDate();'; #write-host $sql;
    invoke-sqlcmd -query $sql -connectionstring $Conn;
    $sql = $null
};

function update_intake_table{
    $sql = 'Update [AdHocData].[PSLAO].[DPM_Roster_Intake]
            Set [ProcTime]        =GetDate()
               ,[StorageFilePath] ='+$Tick+$TempIntakeFilePath+$Tick+'
               ,[TableName]       ='+$Tick+$Table+$Tick+'
            Where [SessionID]     ='+$Tick+$SessionID+$Tick+'
            and [ProcSeq]         ='+$Tick+$ProcSeq+$Tick+';'; #write-host $sql;
    
    invoke-sqlcmd -query $sql -connectionstring $Conn;
    $sql = $null
}

function intake_multi_tab{
$WsIndex = $worksheet.Index; write-host 'WsIndex:'$WsIndex;
}

$Server              ='hulk';
$Db                  ='adhocdata';
$Schema              ='pslao';
$Conn                ='data source='+$server+';initial catalog='+$db+';integrated security=true;'; #write-host $Conn;
$sql                 =' SELECT max(SessionID)SessionID,[User],DPMID,[TYPE],Segment,IntakeFilePath,SheetName, ProcSeq 
                        FROM [AdHocData].[PSLAO].[DPM_Roster_Intake]
                        Where ProcTime is null 
                        Group by [User],DPMID,[TYPE],Segment,IntakeFilePath, SheetName, ProcSeq
                        Order by SessionID, ProcSeq;';
$TablePrefix         ='DPM_INTAKE';
$Tick                = "'";

$psTable=invoke-sqlcmd -query $sql -connectionstring $Conn

Foreach($row in $psTable){
    $SessionID      = $Row[0]; Write-Host 'SessionID:'$SessionID
    $User           = $Row[1]; Write-Host 'User:'$User
    $DPMID          = $Row[2]; Write-Host 'DPMID:'$DPMID
    $Type           = $Row[3]; Write-Host 'Type:'$Type
    $Segment        = $Row[4]; Write-Host 'Segment:'$Segment
    $IntakeFilePath = $Row[5]; Write-Host 'IntakeFilePath:'$IntakeFilePath
    $SheetName      = $Row[6]; Write-Host 'SheetName:'$SheetName
    $ProcSeq        = $Row[7]; Write-Host 'ProcSeq:'$ProcSeq
    $IsFile         = (Get-Item $IntakeFilePath) -is [System.IO.FileInfo]; Write-Host 'IsFile:'$IsFile
    $TempDirPath   ='Q:\D858\F65006\SHARED\CompanyRead\Provider Listings\DPM\Automate\'; Write-Host 'TempDirPath:'$TempDirPath

    copy-item $IntakeFilePath -Destination $TempDirPath

    Switch ($IsFile) {
        $True {
            ETL_to_Server;
        };
        $False{
        };
    };
}; 

#Fire off server sproc

