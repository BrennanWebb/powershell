param(
    [string] $SourceFilePath,
)

write-host $SourceFilePath

#Variables
$FileAttr                     = Get-Item $SourceFilePath;
$FileExt                      = $FileAttr.Extension;
$FileDir                      = $FileAttr.DirectoryName;
$FileDirName                  = [System.IO.Path]::GetFileName($FileDir);
$FileName                     = ($FileAttr.Name -replace $FileExt,'');
$FileModifyTimeStamp          = $FileAttr.LastWriteTime;
$FileModifyTimeStampFormatted = ($FileAttr.LastWriteTime -f 'YYYY-MM-DD HH:MM:SS'-replace '[^0-9]', '');
#Get-ItemProperty $SourceFilePath |  Format-list -Property * -Force

#Server Variables
$Server              ='hulk';
$Db                  ='adhocdata';
$Schema              ='pslao';
$Conn                ='data source='+$server+';initial catalog='+$db+';integrated security=true;'; #write-host $Conn;
$TablePrefix         ='DPM_TEMP_';
$Tick                = "'"; #write-host $Tick;

if (Test-Path $SourceFilePath) {
    switch -exact ($FileExt){
        '.csv'{
            #write-host ('.csv source file')
            $Table=($TablePrefix+$FileDirName+'_'+$FileName) -replace ' ', ''; #write-host $Table;
            drop_table $Table;
            ,(import-csv $SourceFilePath) | Write-SqlTableData -ServerInstance $server -DatabaseName $db -SchemaName $schema -TableName $Table -Force
            alter_table $Table;
            update_table $Table;
            $Table=$null;
        }
        {($_ -eq '.xls') -or ($_ -eq '.xlsx')} {
            #write-host ('.xls, .xlsx source file')
            #check for .xls file type.  If true convert to .xlsx
            if ($FileExt ='.xls'){
                ConvertTo-ExcelXlsx -Path $SourceFilePath -Force
                
                #Set original .xls file back to original modify timestamp
                $FileAttr.LastWriteTime = $FileModifyTimeStamp

                #replace extension with .xlsx and set new file modify timestamp to original timestamp.
                $SourceFilePath=$SourceFilePath -replace '.xls','.xlsx'
                $FileAttr               = Get-Item $SourceFilePath 
                $FileAttr.LastWriteTime = $FileModifyTimeStamp
            }

            #scope in vba for specific workbook manipulation
            $xl = New-Object -ComObject Excel.Application
            $xl.Visible = $false
            $xl.Application.DisplayAlerts= $false
            $Workbook = $xl.Workbooks.Open($SourceFilePath)
            foreach ($worksheet in $workbook.sheets) {
                $SheetName = $worksheet.Name; #write-host $SheetName; 
                $Table=($TablePrefix+$FileDirName+'_'+$FileName+'_'+$SheetName) -replace ' ', ''; #write-host $Table;
                drop_table $Table;
                ,(Import-Excel $SourceFilePath -WorksheetName $SheetName -AsText *) 
                | Write-SqlTableData -ServerInstance $server -DatabaseName $db -SchemaName $schema -TableName $Table -Force;
                alter_table $Table;
                update_table $Table;
                $Table=$null;
            } 
            $Workbook.Close();
            $xl.Quit();
            [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($xl);

            Remove-Item $SourceFilePath; 
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
            Add [SourceFilePath] varchar(500)
               ,[SessionID] varchar(25)
               ,[User]    varchar(25)
               ,[DPMID]   varchar(25)
               ,[Type]    varchar(25)
               ,[Segment] varchar(25)
               ,[LoadTime] datetime;'; write-host $sql;
    invoke-sqlcmd -query $sql -connectionstring $Conn;
    $sql = $null
};

function update_table {
    param ($Table);$sql = $null
    $sql = 'Update [adhocdata].[pslao].['+$Table+']
            Set [SourceFilePath] ='+$Tick+$SourceFilePath+$Tick+'
               ,[SessionID] ='+$Tick+$SessionID+$Tick+'
               ,[User]    ='+$Tick+$User+$Tick+'
               ,[DPMID]   ='+$Tick+$DPMID+$Tick+'
               ,[Type]    ='+$Tick+$Type+$Tick+'
               ,[Segment] ='+$Tick+$Segment+$Tick+'
               ,[LoadTime] =GetDate();'; write-host $sql;
    invoke-sqlcmd -query $sql -connectionstring $Conn;
    $sql = $null
};