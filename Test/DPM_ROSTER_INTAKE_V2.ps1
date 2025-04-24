Function ETL_to_Server {
    #Variables
    $FileAttr                     = Get-Item $IntakeFilePath
    $FileExt                      = $FileAttr.Extension; Log -Message ('FileExt: '+$FileExt);
    $FileDir                      = $FileAttr.DirectoryName; Log -Message  ('FileDir:'+$FileDir);
    $FileDirName                  = [System.IO.Path]::GetFileName($FileDir); Log -Message  ('FileDirName:'+$FileDirName);
    $FileName                     = ($FileAttr.Name -replace $FileExt,''); Log -Message  ('FileName:'+$FileName);
    $TempIntakeFilePath           = $TempDirPath+$FileAttr.Name; Log -Message  ('TempIntakeFilePath:'+$TempIntakeFilePath);
    $FileModifyTimeStamp          = $FileAttr.LastWriteTime; Log -Message  ('FileModifyTimeStamp:'+$FileModifyTimeStamp);
    $FileModifyTimeStampFormatted = ($FileAttr.LastWriteTime -f 'YYYY-MM-DD HH:MM:SS'-replace '[^0-9]', ''); Log -Message ('FileModifyTimeStampFormatted:'+$FileModifyTimeStampFormatted);
    #Get-ItemProperty $IntakeFilePath |  Format-list -Property * -Force

    if (Test-Path $TempIntakeFilePath) {
        switch -exact ($FileExt){
            '.csv'{
                #write-host ('.csv source file')
                $Table=($TablePrefix+'_'+$FileName) -replace ' ', ''; #write-host $Table;
                $SheetName='' #Blank Sheet name for CSV
                drop_table $Table;
                $m = ,(import-csv $TempIntakeFilePath) | Write-SqlTableData -ServerInstance $server -DatabaseName $db -SchemaName $schema -TableName $Table -Force -Timeout 600;
                Log -Message ('SQL Response:'+$m)
                alter_table $Table;
                update_table $Table;
                $Table=$null;
            }
            {$_ -eq '.xls' -or $_ -eq '.xlsx'} {
                
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

                #Import Excel sheet data to server
                $Table=($TablePrefix+'_'+$FileName+'_'+$SheetName) -replace ' ', ''; #write-host $Table;
                drop_table $Table;
                $m = ,(Import-Excel $TempIntakeFilePath -WorksheetName $SheetName -AsText *)`
                | Write-SqlTableData -ServerInstance $server -DatabaseName $db -SchemaName $schema -TableName $Table -Force -Timeout 600;
                Log -Message ('SQL Response:'+$m)
                alter_table $Table;
                update_table $Table;
                Remove-Item $TempIntakeFilePath;
                update_intake_table; 
            };
        };
    };
};

function drop_table {
    param ($Table);$sql = $null
    $sql = 'Drop table [adhocdata].[pslao].['+ $Table +'];'; #write-host $sql;
    $m = invoke-sqlcmd -query $sql -connectionstring $Conn -Verbose 4>&1;
    Log -Message ('SQL Response:'+$m)
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
               ,[LoadTime]       datetime
               ,[Action]         varchar(25)
               ,[RID]            int identity(1,1);'; #write-host $sql;
    $m = invoke-sqlcmd -query $sql -connectionstring $Conn -Verbose 4>&1;
    Log -Message ('SQL Response:'+$m)
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
               ,[Action]         ='+$Tick+$Action+$Tick+'
               ,[LoadTime]       =GetDate();'; #write-host $sql;
    $m = invoke-sqlcmd -query $sql -connectionstring $Conn -Verbose 4>&1;
    Log -Message ('SQL Response:'+$m)
    $sql = $null
};

function update_intake_table{
    $sql = 'Update [AdHocData].[PSLAO].[DPM_Roster_Intake]
            Set [PsProcTime]      =GetDate()
               ,[StorageFilePath] ='+$Tick+$TempIntakeFilePath+$Tick+'
               ,[TableName]       ='+$Tick+$Table+$Tick+'	     
            Where [SessionID]     ='+$Tick+$SessionID+$Tick+'
            and [PsProcSeq]         ='+$Tick+$PsProcSeq+$Tick+';'; #write-host $sql;
    
    $m = invoke-sqlcmd -query $sql -connectionstring $Conn -Verbose 4>&1;
    Log -Message ('SQL Response:'+$m)
    $sql = $null
};

function trigger_server_proc{
    $sql ='Exec [ADHOCDATA].[PSLAO].[sp_DPM_Roster_Intake];'; #write-host $sql;
    $m=invoke-sqlcmd -query $sql -connectionstring $Conn -Verbose 4>&1;
    Log -Message ('SQL Response:'+$m)
    $sql = $null
};

function Log{
    param([string]$Message)
    $LogFile ='Q:\D858\F65006\SHARED\CompanyRead\Provider Listings\DPM\Automate\PowerShell\DPM_ROSTER_INTAKE_LOG.txt';
    if(Test-Path $LogFile) {} else {New-Item $LogFile -ItemType file};
    ((Get-Date).ToString() + " - " + $Message) >> $LogFile;

};

Try
{
    Log -Message '************************** PWSH Start **************************'
    $Server              ='hulk';
    $Db                  ='adhocdata';
    $Schema              ='pslao';
    $Conn                ='data source='+$server+';initial catalog='+$db+';integrated security=true;'; #write-host $Conn;
    $sql                 =' SELECT max(SessionID)SessionID,[User],DPMID,[TYPE],Segment,IntakeFilePath,SheetName, PsProcSeq,Action 
                            FROM [AdHocData].[PSLAO].[DPM_Roster_Intake]
                            Where PsProcTime is null 
                            Group by [User],DPMID,[TYPE],Segment,IntakeFilePath, SheetName, PsProcSeq,Action
                            Order by SessionID, PsProcSeq;';
    $TablePrefix         ='DPM_INTAKE';
    $Tick                = "'";

    $psTable=invoke-sqlcmd -query $sql -connectionstring $Conn

    Foreach($row in $psTable){
        $SessionID      = $Row[0]; Log -Message  ('SessionID:'+$SessionID)
        $User           = $Row[1]; Log -Message  ('User:'+$User)
        $DPMID          = $Row[2]; Log -Message  ('DPMID:'+$DPMID)
        $Type           = $Row[3]; Log -Message  ('Type:'+$Type)
        $Segment        = $Row[4]; Log -Message  ('Segment:'+$Segment)
        $IntakeFilePath = $Row[5]; Log -Message  ('IntakeFilePath:'+$IntakeFilePath)
        $SheetName      = $Row[6]; Log -Message  ('SheetName:'+$SheetName)
        $PsProcSeq      = $Row[7]; Log -Message  ('PsProcSeq:'+$PsProcSeq)
        $Action         = $Row[8]; Log -Message  ('Restate:'+$Action)
        $IsFile         = (Get-Item $IntakeFilePath) -is [System.IO.FileInfo]; Log -Message ('IsFile:'+$IsFile)
        $TempDirPath   ='Q:\D858\F65006\SHARED\CompanyRead\Provider Listings\DPM\Automate\'; Log -Message ('TempDirPath:'+$TempDirPath)

        copy-item $IntakeFilePath -Destination $TempDirPath

        Switch ($IsFile) {
            $True {
                Log -Message '_____________________ETLStart_____________________'
                ETL_to_Server; 
                Log -Message '_____________________ETL End______________________'
            };
            $False{
            };
        };
    };

    #Trigger SQL Stored Proc
    #Log -Message '_____________________SQL Exec_____________________'
    #trigger_server_proc
    #Log -Message '_____________________SQL End______________________'
}

Catch
{
    Log -Message $_.Exception.Message;
    Log -Message $_.Exception.ItemName;
}

Log -Message '_____________________PWSH End______________________'

