
<#Overview:
The following code will take in the path of an excel workbook and loop through each worksheet.
Each worksheet will be written to table under the designated server.database.schema.worksheet name.
The -force tag will overwrite any existing tables.
Powershell 5.0, 7.1.4 tested.
Required Modules: Install-Module -Name ImportExcel, SqlServer;
#>

[string] $IntakeFilePath ="{File Path}";
$Worksheets=Get-ExcelSheetInfo -Path $IntakeFilePath #generates a table for
$Server ="{Server}";
$Db="{Database}";
$Schema ="{Schema}"
$IsFile = (Get-Item $IntakeFilePath) -is [System.IO.FileInfo];

#Setup a function which will loop and return what the affected record count.
function ETL_to_Server {
    foreach ($Worksheet in $Worksheets) {
        $TableName = $Worksheet.Name.Replace(".","_");
        $Worksheet = $Worksheet.Name
        ,(Import-Excel -Path $IntakeFilePath -WorksheetName $Worksheet -AsText *) |
        Write-SqlTableData -ServerInstance $Server -Database $Db -SchemaName $Schema -TableName $TableName -Force;
    }   
};

#try calling function
try {
    #check if path exists
    switch ($IsFile){
        $True {
            #since file exists, call function
            ETL_to_Server; 
        };
        $False{
            #file does not exist, raise error.
            Write-Host "Supplied file path does not exist."
        };
    }
    
}
catch {
    #If any other errors not related to $isfile is found, roll it here.
    Write-Host $_.Exception.Message;
};
