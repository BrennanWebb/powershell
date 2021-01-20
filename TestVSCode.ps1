param(
    [string] $SourceFilePath='Q:\D858\F65006\SHARED\CompanyRead\Provider Listings\DPM\Automate\HPN\Copy of 05-16-20 thru 05-30-20 HPN Appr Chng Trm - HPN.xlsx'
)


function drop_table {
    param ($table)
    $sql = 'Drop table '+ $table + ';'
    $conn="data source=$server;initial catalog=$db;integrated security=true;"
    write-host $sql  
    Clear-Variable -name sql
}

#File Variables
$FileAttr            = Get-Item $SourceFilePath 
$FileExt             = $FileAttr.Extension
$FileDir             = $FileAttr.DirectoryName 
$FileDirName         = [System.IO.Path]::GetFileName($FileDir)
$FileModifyTimeStamp = $FileAttr.LastWriteTime -f 'YYYY-MM-DD HH:MM:SS'-replace '[^0-9]', '' #write-host $FileModifyTimeStamp
#Get-ItemProperty $SourceFilePath |  Format-list -Property * -Force

#Server Variables
$server='hulk'
$db='adhocdata'
$schema='pslao'


if (Test-Path $SourceFilePath) {
    switch -exact ($FileExt){
        '.csv'{
            #write-host ('.csv source file')
            $table='BW_'+$FileDirName+'_'+$FileModifyTimeStamp
            #write-host $table
            ,(import-csv $SourceFilePath) | Write-SqlTableData -ServerInstance $server -DatabaseName $db -SchemaName $schema -TableName $table -Force
        }
        {($_ -eq '.xls') -or ($_ -eq '.xlsx')} {
            #write-host ('.xls, .xlsx source file')
            $xl = New-Object -ComObject Excel.Application
            $xl.Visible = $false
            $xl.Application.DisplayAlerts= $false
            $Workbook = $xl.Workbooks.Open($SourceFilePath)
            foreach ($worksheet in $workbook.sheets) {
                Clear-Variable -name table
                $SheetName = $worksheet.Name 
                #write-host $SheetName 
                $table='BW_'+$FileDirName+'_'+$FileModifyTimeStamp+'_'+$SheetName -replace ' ', ''
                #write-host $table
                drop_table $table
                ,(Import-Excel $SourceFilePath -WorksheetName $SheetName -AsText *) | Write-SqlTableData -ServerInstance $server -DatabaseName $db -SchemaName $schema -TableName $table -Force
            } 
            $Workbook.Close()
            $xl.Quit()

        }
    }
}