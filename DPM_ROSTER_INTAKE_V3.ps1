param(
    [string] $IntakeFilePath,
    [string] $SheetName,
    [string] $SessionID = 'Session',
    [string] $User = 'User',
    [string] $DPMID = 'DPMID',
    [string] $Type = 'Type',
    [string] $Segment = 'Segment',
    [string] $Action = 'Action'
);
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
                $m = ,(import-csv $TempIntakeFilePath)`
                | Write-SqlTableData -ServerInstance $server -DatabaseName $db -SchemaName $schema -TableName $Table -Force -Timeout 0;
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
                | Write-SqlTableData -ServerInstance $server -DatabaseName $db -SchemaName $schema -TableName $Table -Force -Timeout 0;
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
    Log -Message '************************** PS Start **************************'
    $Server             ='hulk';Log -Message ('Server: '+$Server);
    $Db                 ='adhocdata';Log -Message ('Db: '+$Db);
    $Schema             ='pslao';Log -Message ('Schema: '+$Schema);
    $Conn               ='data source='+$server+';initial catalog='+$db+';integrated security=true;';Log -Message ('Conn: '+$Conn);
    $TablePrefix        ='DPM_INTAKE';Log -Message ('TablePrefix: '+$TablePrefix);
    $TempDirPath        ='Q:\D858\F65006\SHARED\CompanyRead\Provider Listings\DPM\Automate\';Log -Message ('TempDirPath: '+$TempDirPath);
    $IsFile             = (Get-Item $IntakeFilePath) -is [System.IO.FileInfo];Log -Message ('IsFile: '+$IsFile);
    $Tick               = "'";

    copy-item $IntakeFilePath -Destination $TempDirPath

    #$psTable=invoke-sqlcmd -query $sql -connectionstring $Conn

    Switch ($IsFile) {
        $True {
            ETL_to_Server; 
        };
        $False{
            #possible server email notification?
            #write error back to access?
        };
    };
}
Catch
{
    Log -Message $_.Exception.Message;
    Log -Message $_.Exception.ItemName;
};

Log -Message '************************** PS END **************************'






