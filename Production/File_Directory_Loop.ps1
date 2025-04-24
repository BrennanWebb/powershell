param (
    [string] $FileDir = 'U:\Automate\CC\'
);
$FileDirCount      = (Get-ChildItem $FileDir | Measure-Object).Count ; write-host '$FileDirCount: '$FileDirCount;
#$IsFolder          = (Get-Item $FileDir) -is [System.IO.DirectoryInfo]; write-host '$IsFolder: '$IsFolder; 

Get-ChildItem $FileDir |  #-Filter *.xlsx
Foreach-Object {
    $File= (Get-Item $_.FullName);
    $IsFile   = (Get-Item $File) -is [System.IO.FileInfo];
    Switch ($IsFile) {
        $True {
            write-host $File;
            D:\Humana\PowerShell\Import_File_To_SQL_Table.ps1 -SourceFilePath $File
        };
        $False{

        };
    };  
};