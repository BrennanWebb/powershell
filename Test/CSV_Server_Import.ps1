<#
.SYNOPSIS
    Imports a large CSV file into a Microsoft SQL Server table with high performance.

.DESCRIPTION
    This script is designed to handle very large CSV files (multi-gigabyte) for import into a SQL Server database.
    It uses the .NET SqlBulkCopy class for maximum performance and efficiency.

    Features:
    - Parses a 4-part SQL naming convention ([Server].[Database].[Schema].[Table]).
    - Uses Windows Authentication (trusted connection).
    - Checks if the destination table exists.
    - If the table does not exist, it infers the schema from the CSV and creates the table dynamically.
    - Includes a -Force switch to drop and recreate the destination table if it exists.
    - Uses batch loading for minimal memory overhead and high speed.

.PARAMETER Path
    The full path to the source CSV file. This parameter is mandatory.

.PARAMETER SqlServerTarget
    The full 4-part name of the destination SQL table.
    Example: '[MyServer.Domain.com].[MyDatabase].[dbo].[MyNewTable]'

.PARAMETER BatchSize
    The number of records to write to the server in a single batch.
    A larger batch size can improve performance but uses more memory.

.PARAMETER DefaultVarcharLength
    When creating a new table, this specifies the length for columns inferred as text (string) data.
    For example, 255 would create VARCHAR(255). Use 'max' for VARCHAR(MAX).
    
.PARAMETER Timeout
    The wait time, in seconds, for the bulk copy command to execute. A value of 0 indicates no time limit.

.PARAMETER DebugMode
    If specified, the script will output additional verbose information for troubleshooting.

.PARAMETER Force
    If specified, the script will drop the destination table if it already exists and recreate it.
    Use with caution as this will delete all data in the existing table.

.EXAMPLE
    .\CSV_Importer.ps1 -Path "C:\Data\HugeFile.csv" -SqlServerTarget "[SQL01].[StagingDB].[dbo].[ImportData]" -Force
    
    This command imports the 'HugeFile.csv' into the 'ImportData' table. If the table already exists, it will be dropped and recreated before the import.

.EXAMPLE
    .\CSV_Importer.ps1 -Path "C:\Data\AnotherFile.csv" -SqlServerTarget "[SQL01].[StagingDB].[dbo].[NewTable]" -BatchSize 1000 -DefaultVarcharLength 500 -DebugMode
    
    This command imports data into a new table, setting the batch size to 1,000 to troubleshoot network issues, creating text columns as VARCHAR(500), and enabling debug output.
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$Path,

    [Parameter(Mandatory = $true)]
    [string]$SqlServerTarget,

    [Parameter(Mandatory = $false)]
    [int]$BatchSize = 5000,

    [Parameter(Mandatory = $false)]
    [string]$DefaultVarcharLength = '255',
    
    [Parameter(Mandatory = $false)]
    [int]$Timeout = 0, # 0 = Infinite Timeout

    [Parameter(Mandatory = $false)]
    [switch]$DebugMode,

    [Parameter(Mandatory = $false)]
    [switch]$Force
)

# ---
#   Designer: Brennan Webb & Gemini
#   Script Engine: Gemini
#   Version: 1.5.1
#   Created: 2025-09-10
#   Modified: 2025-09-10
#   Change Log:
#       1.5.1: Implemented a more robust, Unicode-based method for stripping the BOM from column headers to prevent potential provider instability.
#       1.5.0: Changed numeric type inference to default to DECIMAL for safety. Added logic to strip UTF-8 BOM from the first column header.
#       1.4.1: Suppressed unwanted console output from Add-Member and DataTable.Rows.Add methods during loops.
#       1.4.0: Refactored file access to exclusively use the OLE DB provider, minimizing file lock time. Suppressed noisy output from DROP TABLE command.
#       1.3.0: Added -Timeout parameter and set BulkCopyTimeout to prevent timeout errors on large files. Reworked import logic for manual batching and improved progress reporting.
#       1.2.0: Added -Force switch to drop/recreate existing tables. Added version output in DebugMode.
#       1.1.0: Overhauled schema inference logic to be more robust. Added stricter date validation.
#       1.0.0: Initial script creation for high-performance SQL bulk import.
# ---

# --- Helper Functions ---
function Write-Log {
    param(
        [string]$Message,
        [string]$ForegroundColor
    )
    Write-Host $Message -ForegroundColor $ForegroundColor
}

function Test-IsDateString {
    param($stringValue)
    # An array of common date/time formats to check against. Add more as needed.
    $formats = @(
        'yyyy-MM-dd',
        'MM/dd/yyyy',
        'M/d/yyyy',
        'yyyy-MM-dd HH:mm:ss',
        'MM/dd/yyyy HH:mm:ss',
        'M/d/yyyy h:mm:ss tt'
    )
    foreach ($format in $formats) {
        try {
            # Use [DateTime]::ParseExact for strict format checking.
            [DateTime]::ParseExact($stringValue, $format, [System.Globalization.CultureInfo]::InvariantCulture) | Out-Null
            return $true # Success, it matches this format.
        }
        catch {}
    }
    return $false # Failed to match any of the defined formats.
}


# --- Main Script Body ---
# Add .NET assembly for DataTable if not already loaded
Add-Type -AssemblyName System.Data

# Begin main execution block with error handling
$Global:Error.Clear()
try {
    if ($DebugMode) { Write-Log -Message "DEBUG: Script Version: 1.5.1" -ForegroundColor Yellow }

    #region Parameter Validation and Interactive Prompts
    if (-not $PSBoundParameters.ContainsKey('Path')) {
        $Path = Read-Host -Prompt "Please enter the full path to the CSV file"
    }
    if (-not $PSBoundParameters.ContainsKey('SqlServerTarget')) {
        $SqlServerTarget = Read-Host -Prompt "Please enter the 4-part SQL target (e.g., [Server].[DB].[Schema].[Table])"
    }

    if (-not (Test-Path -Path $Path -PathType Leaf)) {
        throw "The file specified by -Path does not exist: $Path"
    }
    if ($Path -notlike "*.csv") {
        throw "The file specified by -Path must be a .csv file."
    }
    #endregion

    #region Parse 4-Part SQL Target Name
    Write-Log -Message "Parsing SQL target name: $SqlServerTarget" -ForegroundColor Cyan
    $regex = '^\[([^\]]+)\]\.\[([^\]]+)\]\.\[([^\]]+)\]\.\[([^\]]+)\]$'
    if ($SqlServerTarget -match $regex) {
        $serverName = $matches[1]
        $databaseName = $matches[2]
        $schemaName = $matches[3]
        $tableName = $matches[4]
        Write-Log -Message "  - Server: $serverName" -ForegroundColor Cyan
        Write-Log -Message "  - Database: $databaseName" -ForegroundColor Cyan
        Write-Log -Message "  - Schema: $schemaName" -ForegroundColor Cyan
        Write-Log -Message "  - Table: $tableName" -ForegroundColor Cyan
    }
    else {
        throw "The -SqlServerTarget format is invalid. Expected format: [Server].[Database].[Schema].[Table]"
    }
    #endregion

    #region SQL Connection and Table Check
    $connectionString = "Server=$serverName;Database=$databaseName;Integrated Security=True;"
    if ($DebugMode) { Write-Log -Message "DEBUG: Connection String: $connectionString" -ForegroundColor Yellow }

    $sqlConnection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    Write-Log -Message "Connecting to SQL server..." -ForegroundColor Cyan
    $sqlConnection.Open()
    Write-Log -Message "Connection successful." -ForegroundColor Green

    $checkTableQuery = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = '$schemaName' AND TABLE_NAME = '$tableName';"
    $command = New-Object System.Data.SqlClient.SqlCommand($checkTableQuery, $sqlConnection)
    $tableExists = $command.ExecuteScalar()
    #endregion

    #region Handle -Force switch for Existing Table
    if ($Force -and $tableExists) {
        Write-Log -Message "WARNING: -Force switch specified. Dropping existing table '[$schemaName].[$tableName]'." -ForegroundColor Yellow
        $dropCommand = New-Object System.Data.SqlClient.SqlCommand("DROP TABLE [$schemaName].[$tableName];", $sqlConnection)
        $dropCommand.ExecuteNonQuery() | Out-Null
        Write-Log -Message "Table dropped successfully." -ForegroundColor Green
        $tableExists = $false # Update status so the creation block will run
    }
    #endregion

    #region Schema Inference and Table Creation
    if (-not $tableExists) {
        Write-Log -Message "Table '[$schemaName].[$tableName]' not found. Preparing to create it." -ForegroundColor Yellow
        
        Write-Log -Message "Sampling first 1000 rows of CSV to infer data types..." -ForegroundColor Cyan
        $provider = "Microsoft.ACE.OLEDB.12.0"
        $csvConnString = "Provider=$provider;Data Source=`"$(Split-Path $Path)`";Extended Properties=`"text;HDR=Yes;FMT=Delimited`";"
        $tempCsvConnection = $null
        $tempDataReader = $null
        $csvHeaders = @()
        $sampleData = @()

        try {
            $tempCsvConnection = New-Object System.Data.OleDb.OleDbConnection($csvConnString)
            $tempCsvConnection.Open()
            $tempCommand = New-Object System.Data.OleDb.OleDbCommand("SELECT * FROM [$(Split-Path $Path -Leaf)]", $tempCsvConnection)
            $tempDataReader = $tempCommand.ExecuteReader()

            for ($i = 0; $i -lt $tempDataReader.FieldCount; $i++) {
                $csvHeaders += $tempDataReader.GetName($i)
            }

            # Clean the BOM (Byte Order Mark) from the first header if it exists
            $csvHeaders[0] = $csvHeaders[0].TrimStart([char]0xFEFF) # U+FEFF is the Unicode BOM character
            if ($DebugMode) { Write-Log -Message "DEBUG: First column header cleaned to: $($csvHeaders[0])" -ForegroundColor Yellow }


            $rowCount = 0
            while ($tempDataReader.Read() -and $rowCount -lt 1000) {
                $obj = [PSCustomObject]::new()
                foreach($header in $csvHeaders){
                    # Use the original (potentially unclean) header to read from the datareader
                    $originalHeader = $tempDataReader.GetName($csvHeaders.IndexOf($header))
                    $obj | Add-Member -MemberType NoteProperty -Name $header -Value $tempDataReader[$originalHeader] | Out-Null
                }
                $sampleData += $obj
                $rowCount++
            }
        }
        finally {
            if ($tempDataReader) { $tempDataReader.Close() }
            if ($tempCsvConnection) { $tempCsvConnection.Close() }
        }

        if ($csvHeaders.Count -eq 0) { throw "CSV file appears to be empty or header is invalid." }
        
        $columnTypes = [System.Collections.Specialized.OrderedDictionary]::new()

        foreach ($header in $csvHeaders) {
            $isStillDecimal = $true; $isStillDateTime = $true
            foreach ($row in $sampleData) {
                $value = $row.$header
                if ([string]::IsNullOrWhiteSpace($value)) { continue }
                if ($isStillDecimal -and ($value -as [decimal]) -eq $null) { $isStillDecimal = $false }
                if ($isStillDateTime -and -not (Test-IsDateString -stringValue $value)) { $isStillDateTime = $false }
            }

            # SAFER LOGIC: Default all numbers to DECIMAL to handle currency and integers.
            if ($isStillDecimal) { $columnTypes[$header] = 'DECIMAL(18, 2)' }
            elseif ($isStillDateTime) { $columnTypes[$header] = 'DATETIME' }
            else { $columnTypes[$header] = "VARCHAR($DefaultVarcharLength)" }
        }
        
        $createTableQuery = "CREATE TABLE [$schemaName].[$tableName] ("
        $columnDefinitions = @()
        foreach ($key in $columnTypes.Keys) {
            $columnDefinitions += " `n  [$key] $($columnTypes[$key])"
        }
        $createTableQuery += $columnDefinitions -join ","
        $createTableQuery += "`n);"

        if ($DebugMode) { Write-Log -Message "DEBUG: Generated CREATE TABLE statement:`n$createTableQuery" -ForegroundColor Yellow }

        Write-Log -Message "Executing CREATE TABLE statement..." -ForegroundColor Cyan
        $command.CommandText = $createTableQuery
        $command.ExecuteNonQuery() | Out-Null
        Write-Log -Message "Table '[$schemaName].[$tableName]' created successfully." -ForegroundColor Green
    }
    else {
        Write-Log -Message "Target table '[$schemaName].[$tableName]' already exists. Appending data." -ForegroundColor Cyan
    }
    #endregion

    #region Bulk Import Data
    Write-Log -Message "Starting bulk import process..." -ForegroundColor Cyan

    $bulkCopy = New-Object System.Data.SqlClient.SqlBulkCopy($sqlConnection)
    $bulkCopy.DestinationTableName = "[$schemaName].[$tableName]"
    $bulkCopy.BatchSize = $BatchSize
    $bulkCopy.BulkCopyTimeout = $Timeout # Set timeout using the new parameter
    
    # Note: Requires Microsoft Access Database Engine 2010/2016 Redistributable for the OLE DB provider.
    $provider = "Microsoft.ACE.OLEDB.12.0"
    $csvConnString = "Provider=$provider;Data Source=`"$(Split-Path $Path)`";Extended Properties=`"text;HDR=Yes;FMT=Delimited`";"
    if ($DebugMode) { Write-Log -Message "DEBUG: Using OleDb connection for robust CSV parsing." -ForegroundColor Yellow }
    $csvConnection = New-Object System.Data.OleDb.OleDbConnection($csvConnString)
    $csvConnection.Open()
    
    $csvCommand = New-Object System.Data.OleDb.OleDbCommand("SELECT * FROM [$(Split-Path $Path -Leaf)]", $csvConnection)
    $dataReader = $csvCommand.ExecuteReader()

    # Create an in-memory DataTable to hold batches
    $dataTable = New-Object System.Data.DataTable

    # Add columns to the DataTable that match the CSV reader, cleaning the first column name
    $firstColumnName = $dataReader.GetName(0).TrimStart([char]0xFEFF)
    $dataTable.Columns.Add($firstColumnName) | Out-Null

    for ($i = 1; $i -lt $dataReader.FieldCount; $i++) {
        $dataTable.Columns.Add($dataReader.GetName($i)) | Out-Null
    }

    $batchCounter = 1; $totalRowsCopied = 0

    # Loop through the CSV data reader
    while ($dataReader.Read()) {
        $rowValues = New-Object object[]($dataReader.FieldCount)
        $dataReader.GetValues($rowValues) | Out-Null
        $dataTable.Rows.Add($rowValues) | Out-Null

        if ($dataTable.Rows.Count -ge $BatchSize) {
            $bulkCopy.WriteToServer($dataTable)
            $totalRowsCopied += $dataTable.Rows.Count
            Write-Log -Message "  Wrote batch $batchCounter ($totalRowsCopied total rows)..." -ForegroundColor Cyan
            $dataTable.Rows.Clear()
            $batchCounter++
        }
    }

    # Write any remaining rows in the final, partial batch
    if ($dataTable.Rows.Count -gt 0) {
        $bulkCopy.WriteToServer($dataTable)
        $totalRowsCopied += $dataTable.Rows.Count
        Write-Log -Message "  Wrote final batch $batchCounter ($totalRowsCopied total rows)..." -ForegroundColor Cyan
    }

    $dataReader.Close()
    $csvConnection.Close()

    Write-Log -Message "Bulk import completed successfully. Total rows copied: $totalRowsCopied" -ForegroundColor Green
    #endregion
}
catch {
    $errorMessage = $_.Exception.Message
    if ($errorMessage -like "*The '$provider'* provider is not registered on the local machine*") {
        $errorMessage = "The OLE DB provider '$provider' is not registered. Please install the Microsoft Access Database Engine Redistributable (64-bit if using 64-bit PowerShell)."
    }
    Write-Log -Message "ERROR: $errorMessage" -ForegroundColor Red
    if ($DebugMode) {
        Write-Log -Message "DEBUG: Full exception details:" -ForegroundColor Yellow
        $_ | Format-List -Force | Out-String | Write-Host -ForegroundColor Yellow
    }
}
finally {
    if ($sqlConnection -and $sqlConnection.State -eq 'Open') {
        Write-Log -Message "Closing SQL connection." -ForegroundColor Cyan
        $sqlConnection.Close()
    }
}