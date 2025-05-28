param(
    [string]$filepath ='U:\Brennan Webb\Powershell\input\*.csv',
    [string]$savepath ='U:\Brennan Webb\Powershell\output\collate.xlsx'
)

function Add-Worksheet{
    Param(
        $filepath
    )
        Try{
            Write-Host $_.Fullname -ForegroundColor green  
            $ws=$wb.Sheets.Add($mv,$mv,$mv,$filepath)
            $ws.Cells.EntireColumn.AutoFit()
        }
        Catch{
            Write-Host "File cannot be loaded: $filepath" -ForegroundColor red
        }
}
$mv=[System.Reflection.Missing]::Value 

$xl = New-Object -Com Excel.Application
$xl.Application.DisplayAlerts = $false
$xl.visible=$true

$wb=$xl.Workbooks.Add()

Get-ChildItem $filepath -Include *.csv |
    Sort-Object Name -desc|
    ForEach-Object{
        Add-Worksheet $_.Fullname
    }
$wb.Worksheets.Item('Sheet1').Delete()
$wb.SaveAs($savepath)
$wb.Close()
$xl.Quit()
