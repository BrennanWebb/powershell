<#
.SYNOPSIS
    Imports all worksheets from a specified Excel file into a SQL Server database.
.DESCRIPTION
    This script prompts the user for an Excel file, SQL Server connection details, and a target schema.
    It then iterates through each worksheet in the Excel file, creating a corresponding table in the
    SQL database and importing the data. It can optionally drop existing tables before import.
.PARAMETER IntakeFilePath
    The full path to the source Excel workbook. If not provided, a file picker dialog will open.
.PARAMETER Server
    The name of the target SQL Server instance (e.g., 'localhost\SQLEXPRESS'). This script uses Integrated Security (Windows Authentication).
.PARAMETER Database
    The name of the target database on the SQL Server.
.PARAMETER Schema
    The name of the schema within the database where tables will be created (e.g., 'dbo'). Defaults to 'dbo'.
.PARAMETER DropExistingTables
    A switch parameter that, if present, will drop the target table if it already exists before importing data.
.PARAMETER DebugMode
    A switch parameter that enables verbose debugging output for troubleshooting.
.EXAMPLE
    .\Import-ExcelToSql.ps1 -IntakeFilePath "C:\Data\MonthlyReport.xlsx" -Server "SQL01" -Database "Staging"
    This command imports all worksheets from 'MonthlyReport.xlsx' into the 'Staging' database on 'SQL01' using Windows Authentication.
.NOTES
    Designer: Brennan Webb & Gemini
    Script Engine: Gemini
    Version: 2.0.1
    Created: 2025-07-18
    Modified: 2025-09-15
    Change Log:
    - 1.0.0: Initial script creation.
    - 1.0.1: Corrected logical flaw regarding an assignment operator.
    - 1.0.2: Fixed typo in 'Read-Host' command.
    - 2.0.0: Replaced Write-SqlTableData with .NET SqlBulkCopy for robust timeout control.
    - 2.0.1: Added -TrustServerCertificate to Invoke-Sqlcmd and connection string to handle modern SQL encryption trust requirements.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, HelpMessage = "Enter the full path to the source Excel workbook.")]
    [string]$IntakeFilePath,

    [Parameter(Mandatory = $false, HelpMessage = "Enter the name of the target SQL Server instance.")]
    [string]$Server,

    [Parameter(Mandatory = $false, HelpMessage = "Enter the name of the target database.")]
    [string]$Database,

    [Parameter(Mandatory = $false, HelpMessage = "Enter the name of the target schema (e.g., dbo).")]
    [string]$Schema = 'dbo',

    [Parameter(Mandatory = $false, HelpMessage = "If specified, existing tables will be dropped before data import.")]
    [switch]$DropExistingTables,

    [Parameter(Mandatory = $false, HelpMessage = "Enable verbose debugging output.")]
    [switch]$DebugMode
)

# --- Script Initialization ---

if ($DebugMode) { $DebugPreference = 'Continue' }
try {
    Write-Debug "Checking for required modules: ImportExcel, SqlServer"
    Import-Module ImportExcel -ErrorAction Stop
    Import-Module SqlServer -ErrorAction Stop
}
catch {
    Write-Host "A required module is not installed. Please run: Install-Module -Name ImportExcel, SqlServer -Scope CurrentUser" -ForegroundColor Red
    return
}

# --- Parameter Handling & Interactive Prompts ---

if ([string]::IsNullOrEmpty($IntakeFilePath)) {
    try {
        Write-Host "Please select the Excel file to import." -ForegroundColor White
        Add-Type -AssemblyName System.Windows.Forms
        $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $fileDialog.Title = "Select Excel File"
        $fileDialog.Filter = "Excel Workbooks (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
        if ($fileDialog.ShowDialog() -eq 'OK') {
            $IntakeFilePath = $fileDialog.FileName
            Write-Host "File selected: $IntakeFilePath" -ForegroundColor Cyan
        }
        else {
            Write-Host "File selection cancelled. Exiting script." -ForegroundColor Yellow
            return
        }
    }
    catch {
        Write-Host "Could not open file picker. Please run the script with the -IntakeFilePath parameter." -ForegroundColor Red
        return
    }
}

if (-not (Test-Path -Path $IntakeFilePath -PathType Leaf)) {
    Write-Host "Error: The specified file path does not exist: $IntakeFilePath" -ForegroundColor Red
    return
}

if ([string]::IsNullOrEmpty($Server)) { $Server = Read-Host -Prompt "Enter the SQL Server instance name" }
if ([string]::IsNullOrEmpty($Database)) { $Database = Read-Host -Prompt "Enter the target database name" }


# --- Main Processing Block ---

try {
    Write-Host "Getting worksheet information from '$IntakeFilePath'..." -ForegroundColor Cyan
    $worksheets = Get-ExcelSheetInfo -Path $IntakeFilePath
    Write-Host "Found $($worksheets.Count) worksheets to process." -ForegroundColor Cyan

    foreach ($worksheetInfo in $worksheets) {
        $tableName = $worksheetInfo.Name.Replace(" ", "_").Replace(".", "_")
        $worksheetName = $worksheetInfo.Name

        Write-Host "--------------------------------------------------" -ForegroundColor Cyan
        Write-Host "Processing worksheet: '$worksheetName' -> Table: '[$Schema].[$tableName]'" -ForegroundColor Cyan

        # Read all Excel data into memory
        $data = Import-Excel -Path $IntakeFilePath -WorksheetName $worksheetName -AsText *

        if ($null -eq $data) {
            Write-Host "Worksheet '$worksheetName' is empty. Skipping." -ForegroundColor Yellow
            continue
        }

        # Handle table creation/dropping
        if ($DropExistingTables.IsPresent) {
            Write-Host "Dropping existing table '[$Schema].[$tableName]'..." -ForegroundColor Yellow
            Invoke-Sqlcmd -Query "DROP TABLE IF EXISTS [$Schema].[$tableName];" -ServerInstance $Server -Database $Database -TrustServerCertificate
        }

        # Build a CREATE TABLE query based on the data columns
        $firstRow = $data | Select-Object -First 1
        $columns = $firstRow.PSObject.Properties | ForEach-Object { "[{0}] NVARCHAR(MAX)" -f $_.Name }
        $createQuery = "IF OBJECT_ID('[$Schema].[$tableName]', 'U') IS NULL BEGIN CREATE TABLE [$Schema].[$tableName] ($([string]::Join(',', $columns))); END"
        Write-Debug "Ensuring table exists with query: $createQuery"
        Invoke-Sqlcmd -Query $createQuery -ServerInstance $Server -Database $Database -TrustServerCertificate

        # Convert PowerShell objects to a DataTable for SqlBulkCopy
        $dataTable = New-Object System.Data.DataTable
        $firstRow.PSObject.Properties | ForEach-Object {
            $null = $dataTable.Columns.Add($_.Name)
        }
        foreach ($row in $data) {
            $dataRow = $dataTable.NewRow()
            $row.PSObject.Properties | ForEach-Object {
                $dataRow[$_.Name] = $_.Value
            }
            $dataTable.Rows.Add($dataRow)
        }

        # Use SqlBulkCopy for a robust, timeout-controlled import
        Write-Host "Importing $($dataTable.Rows.Count) rows into '[$Schema].[$tableName]'..." -ForegroundColor Cyan
        $connectionString = "Server=$Server;Database=$Database;Integrated Security=True;TrustServerCertificate=True;"
        $bulkCopy = New-Object System.Data.SqlClient.SqlBulkCopy($connectionString)
        $bulkCopy.DestinationTableName = "[$Schema].[$tableName]"
        $bulkCopy.BulkCopyTimeout = 0 # 0 means no timeout
        $bulkCopy.WriteToServer($dataTable)

        Write-Host "Successfully imported worksheet '$worksheetName'." -ForegroundColor Green
    }

    Write-Host "--------------------------------------------------" -ForegroundColor Cyan
    Write-Host "All worksheets have been processed successfully." -ForegroundColor Green
}
catch {
    Write-Host "An unexpected error occurred during the import process." -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Script execution halted." -ForegroundColor Red
}