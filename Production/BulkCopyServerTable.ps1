<#
.SYNOPSIS
    Copies a SQL Server table's schema and data from a source to a target server using four-part naming.

.DESCRIPTION
    This script facilitates the migration of a SQL Server table, including its schema and data, from a source to a target destination. It uses SQL Server Management Objects (SMO) for schema replication and a memory-efficient streaming method for the data transfer. A progress bar is displayed during the data transfer.

    This script is self-contained and will attempt to install its own dependencies (the 'SqlServer' PowerShell module) if they are not found. All status messages are timestamped.

.PARAMETER Source
    The full four-part name of the source table in the format [Server].[Database].[Schema].[Table]. Brackets are optional.

.PARAMETER Target
    The full four-part name of the target table in the format [Server].[Database].[Schema].[Table]. Brackets are optional.

.PARAMETER SampleSize
    Specifies the number of rows to copy (TOP N). A value of 0 (the default) copies all rows.

.PARAMETER BatchSize
    Specifies the number of rows in each batch sent to the target server during the bulk copy operation. The default is 50,000.

.PARAMETER SchemaOption
    Determines how the target table's schema is created if it does not exist.
    - '1' (Default): Full schema replication via SMO.
    - '2': Simplified schema where all columns are created as NVARCHAR(MAX).

.PARAMETER DebugMode
    A switch that, when present, enables verbose debugging output. This output includes generated SQL, variable states, and function traces.

.EXAMPLE
    .\BulkCopyServerTable.ps1 -Source "SQL01.SalesDB.dbo.Orders" -Target "SQL02.SalesArchive.dbo.Orders_2024"
    This command uses SMO to replicate the 'Orders' table schema and then streams the data, showing a progress bar.

.NOTES
    Version: 30.2
    Author:  Gemini
    
    --- Change Log ---
    - V30.2: Replaced faulty OBJECT_ID check with a more reliable query against INFORMATION_SCHEMA.TABLES.
    - V30.1: Suppressed unwanted output from the SqlBulkCopy.ColumnMappings.Add method.
    - V30.0: Replaced unreliable SMO IsComputed property check with a direct query to sys.columns for definitive computed column detection.
    - V29.0: Implemented pre-loading of column properties with SMO to address computed column filtering.
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$Source,

    [Parameter(Mandatory = $false)]
    [string]$Target,

    [Parameter(Mandatory = $false)]
    [int]$SampleSize = 0,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 2147483647)]
    [int]$BatchSize = 50000,

    [Parameter(Mandatory = $false)]
    [ValidateSet('1', '2')]
    [string]$SchemaOption = '1',

    [Parameter(Mandatory = $false)]
    [switch]$DebugMode
)

# --- Logging and Dependency Functions ---

function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,

        [Parameter(Mandatory=$false)]
        [System.ConsoleColor]$ForegroundColor = 'Cyan'
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $stampedMessage = "[$timestamp] $Message"
    Write-Host -Object $stampedMessage -ForegroundColor $ForegroundColor
}

function Write-DebugMessage {
    param([string]$Message)
    if ($DebugMode) {
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $stampedMessage = "[$timestamp][DEBUG] $Message"
        Write-Host -Object $stampedMessage -ForegroundColor Magenta
    }
}

function Confirm-SqlServerModule {
    if (Get-Module -ListAvailable -Name SqlServer) {
        Write-DebugMessage "The 'SqlServer' PowerShell module is already available."
        return
    }

    Write-Log -Message "INFO: The required 'SqlServer' PowerShell module is not found."
    Write-Log -Message "INFO: Attempting to install it from the PSGallery (this may take a moment)..."
    
    try {
        Install-Module -Name SqlServer -Scope CurrentUser -Repository PSGallery -Force -AcceptLicense -ErrorAction Stop
        Write-Log -Message "SUCCESS: The 'SqlServer' module has been installed." -ForegroundColor Green
    }
    catch {
        Write-Log -Message "ERROR: Failed to automatically install the 'SqlServer' module." -ForegroundColor Red
        Write-Log -Message "Please check your internet connection and PowerShell execution policy, then try running this command manually:" -ForegroundColor Red
        Write-Host "Install-Module -Name SqlServer -Scope CurrentUser" -ForegroundColor White
        exit 1
    }
}

function Show-VersionBanner {
    try {
        $scriptPath = $MyInvocation.MyCommand.Path
        if ($scriptPath -and (Test-Path $scriptPath)) {
            $scriptContent = Get-Content -Path $scriptPath -ErrorAction Stop -Raw
            if ($scriptContent -match '(?m)\.NOTES\s*Version:\s*([0-9]+\.[0-9]+)') {
                $version = $Matches[1]
                $debugString = if ($DebugMode) { " | Debug Mode Enabled" } else { "" }
                Write-Log -Message "--- BulkCopyServerTable | Version: V$($version)$debugString ---"
            }
        }
    } catch {
        Write-Log -Message "WARNING: Could not determine script version." -ForegroundColor Yellow
    }
}


# --- Core Logic Functions ---

function Split-FourPartName {
    param ([ValidateNotNullOrEmpty()][string]$InputName)
    Write-DebugMessage "Parsing four-part name: $InputName"
    $pattern = '(?<!\[[^\]]*)\.(?![^\[]*\])'; $parts = [regex]::Split($InputName, $pattern)
    if ($parts.Count -ne 4) { throw "Input must be in format Server.Database.Schema.Table. Got: $InputName" }
    $cleanParts = $parts | ForEach-Object { $_ -replace '^\[|\]$', '' }
    if ($cleanParts -contains '') { throw "Empty parts are not allowed in input: $InputName" }
    return @{ Server = $cleanParts[0]; Database = $cleanParts[1]; Schema = $cleanParts[2]; Table = $cleanParts[3] }
}

function Get-ConnectionString {
    param ([ValidateNotNullOrEmpty()][string]$FourPartName)
    $parsed = Split-FourPartName $FourPartName
    return "Server=$($parsed.Server);Database=$($parsed.Database);Integrated Security=True;Connect Timeout=15"
}

function Start-DataTransfer {
    param (
        [ValidateNotNullOrEmpty()][string]$SourceFourPartName,
        [ValidateNotNullOrEmpty()][string]$TargetFourPartName,
        [string]$SchemaOptionToUse,
        [int]$SampleSize,
        [int]$BatchSizeToUse
    )
    $src = Split-FourPartName $SourceFourPartName
    $tgt = Split-FourPartName $TargetFourPartName
    $srcConnStr = Get-ConnectionString $SourceFourPartName
    $tgtConnStr = Get-ConnectionString $TargetFourPartName

    # SMO is still useful for schema scripting and getting a base list of columns
    Write-DebugMessage "Using SMO to get source table metadata."
    $sourceServer = New-Object Microsoft.SqlServer.Management.Smo.Server($src.Server)
    $sourceDb = $sourceServer.Databases[$src.Database]
    if (-not $sourceDb) { throw "SMO could not find source database '$($src.Database)'." }
    $sourceTable = $sourceDb.Tables[$src.Table, $src.Schema]
    if (-not $sourceTable) { throw "SMO could not find source table '[$($src.Schema)].[$($src.Table)]'." }
    
    # --- !! RELIABLE COMPUTED COLUMN DETECTION !! ---
    # The SMO IsComputed property can be unreliable. Query sys.columns directly for the truth.
    Write-DebugMessage "Directly querying sys.columns on $($src.Server) for computed column names."
    $computedColumnNames = [System.Collections.Generic.List[string]]::new()
    $computedColQuery = "SELECT name FROM sys.columns WHERE object_id = OBJECT_ID(@TableName) AND is_computed = 1"
    $sourceQueryConn = $null
    try {
        $sourceQueryConn = New-Object System.Data.SqlClient.SqlConnection($srcConnStr)
        $sourceQueryConn.Open()
        $cmd = New-Object System.Data.SqlClient.SqlCommand($computedColQuery, $sourceQueryConn)
        $cmd.Parameters.AddWithValue("@TableName", "[$($src.Schema)].[$($src.Table)]") | Out-Null
        $reader = $cmd.ExecuteReader()
        while ($reader.Read()) {
            $computedColumnNames.Add($reader.GetString(0))
        }
    }
    catch {
        throw "Failed to query source for computed column list: $_"
    }
    finally {
        if ($sourceQueryConn) { $sourceQueryConn.Dispose() }
    }
    
    if ($computedColumnNames.Count -gt 0) {
        Write-DebugMessage "Found $($computedColumnNames.Count) computed column(s) via direct query: $($computedColumnNames -join ', ')"
    } else {
        Write-DebugMessage "No computed columns were found via direct query."
    }

    # Get all column names from SMO, then filter using our reliable list.
    $allColumnNames = $sourceTable.Columns | ForEach-Object { $_.Name }
    $transferableColumns = $allColumnNames | Where-Object { $_ -notin $computedColumnNames }
    Write-DebugMessage "Found $($transferableColumns.Count) transferable (non-computed) columns: $($transferableColumns -join ', ')"

    $tgtConn = New-Object System.Data.SqlClient.SqlConnection $tgtConnStr
    $bulkCopy = $null
    $srcReaderConn = $null
    $reader = $null

    try {
        $tgtConn.Open()
        Write-DebugMessage "Successfully connected to target server: $($tgt.Server)"

        # --- !! RELIABLE TABLE EXISTENCE CHECK !! ---
        # Query INFORMATION_SCHEMA for a definitive check.
        $checkCmd = $tgtConn.CreateCommand()
        $checkCmd.CommandText = "SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = @Schema AND TABLE_NAME = @Table"
        $checkCmd.Parameters.AddWithValue("@Schema", $tgt.Schema) | Out-Null
        $checkCmd.Parameters.AddWithValue("@Table", $tgt.Table) | Out-Null
        $tableExists = $checkCmd.ExecuteScalar()

        if ($tableExists) {
            Write-Log -Message "INFO: Target table [$($tgt.Schema)].[$($tgt.Table)] exists."
            $dataCmd = $tgtConn.CreateCommand(); $dataCmd.CommandText = "SELECT COUNT(*) FROM [$($tgt.Schema)].[$($tgt.Table)]"; $rowCount = $dataCmd.ExecuteScalar()
            if ($rowCount -gt 0) {
                Write-Host "[$([DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))] Target table contains data. Enter 1 to truncate, 2 to insert additional records: " -ForegroundColor White -NoNewline
                if ((Read-Host) -eq '1') {
                    $truncateCmd = $tgtConn.CreateCommand(); $truncateCmd.CommandText = "TRUNCATE TABLE [$($tgt.Schema)].[$($tgt.Table)]"; [void]$truncateCmd.ExecuteNonQuery()
                    Write-Log -Message "SUCCESS: Table [$($tgt.Schema)].[$($tgt.Table)] truncated." -ForegroundColor Green
                }
            }
        } else {
            Write-Log -Message "INFO: Target table [$($tgt.Schema)].[$($tgt.Table)] does not exist. Creating schema..."
            
            if ($SchemaOptionToUse -eq '1') {
                $scripter = New-Object Microsoft.SqlServer.Management.Smo.Scripter($sourceServer)
                $scripter.Options.ScriptDrops = $false; $scripter.Options.WithDependencies = $false; $scripter.Options.Indexes = $true
                $scripter.Options.DriAll = $true; $scripter.Options.NoCollation = $false; $scripter.Options.SchemaQualify = $true
                
                $scriptCollection = $scripter.EnumScript(@($sourceTable.Urn))
                $createScript = ($scriptCollection -join "`nGO`n").Trim()

                $createScript = $createScript.Replace("[$($src.Table)]", "[$($tgt.Table)]")
                $createScript = $createScript.Replace("CONSTRAINT [PK_$($src.Table)]", "CONSTRAINT [PK_$($tgt.Table)]")

                Write-DebugMessage "Executing generated SMO script..."
                foreach ($batch in ($createScript -split '\nGO\n')) {
                    if (-not [string]::IsNullOrWhiteSpace($batch)) {
                        Write-DebugMessage "Executing batch:`n$batch"
                        $createCmd = $tgtConn.CreateCommand(); $createCmd.CommandText = $batch; [void]$createCmd.ExecuteNonQuery()
                    }
                }
            } else {
                $cols = $sourceTable.Columns | ForEach-Object { "[$($_.Name)] NVARCHAR(MAX) NULL" }
                $createSql = "CREATE TABLE [$($tgt.Schema)].[$($tgt.Table)] (`n" + ($cols -join ",`n") + "`n)"
                Write-DebugMessage "Executing CREATE TABLE statement (NVARCHAR MAX):`n$createSql"
                $createCmd = $tgtConn.CreateCommand(); $createCmd.CommandText = $createSql; [void]$createCmd.ExecuteNonQuery()
            }
            Write-Log -Message "SUCCESS: Table [$($tgt.Schema)].[$($tgt.Table)] created." -ForegroundColor Green
        }

        Write-Log -Message "INFO: Preparing data stream..."

        $totalRows = 0
        if ($SampleSize -eq 0) {
            Write-DebugMessage "Executing COUNT_BIG(*) on source to get total rows for progress bar."
            $countConn = New-Object System.Data.SqlClient.SqlConnection($srcConnStr)
            try {
                $countConn.Open()
                $countCmd = $countConn.CreateCommand(); $countCmd.CommandText = "SELECT COUNT_BIG(*) FROM [$($src.Schema)].[$($src.Table)] WITH (NOLOCK)"; $countCmd.CommandTimeout = 300
                $totalRows = $countCmd.ExecuteScalar()
                Write-DebugMessage "Source table has $totalRows total rows."
            } catch { Write-Log -Message "WARNING: Could not retrieve total row count for progress bar." -ForegroundColor Yellow }
            finally { if ($countConn) { $countConn.Dispose() } }
        } else {
            $totalRows = $SampleSize
        }

        # Check for identity columns on the source using SMO
        $hasIdentityColumn = $false
        if ($SchemaOptionToUse -eq '1' -and ($sourceTable.Columns | Where-Object { $_.Identity -eq $true })) {
            $hasIdentityColumn = $true
        }

        $bulkCopyOptions = [System.Data.SqlClient.SqlBulkCopyOptions]::Default
        if ($hasIdentityColumn) { $bulkCopyOptions = [System.Data.SqlClient.SqlBulkCopyOptions]::KeepIdentity }
        
        $bulkCopy = New-Object System.Data.SqlClient.SqlBulkCopy($tgtConn, $bulkCopyOptions, $null)
        $bulkCopy.DestinationTableName = "[$($tgt.Schema)].[$($tgt.Table)]"
        $bulkCopy.BatchSize = $BatchSizeToUse; $bulkCopy.BulkCopyTimeout = 0; $bulkCopy.NotifyAfter = $BatchSizeToUse

        $progressEventHandler = [System.Data.SqlClient.SqlRowsCopiedEventHandler]{
            param($sourceObject, $e)
            $rows = $e.RowsCopied
            $activity = "Streaming data to [$($tgt.Table)]"
            $status = "Copied {0:N0} rows..." -f $rows
            if ($totalRows -gt 0) {
                $percent = [math]::Round(($rows / $totalRows) * 100)
                Write-Progress -Activity $activity -Status $status -PercentComplete $percent
            } else { Write-Progress -Activity $activity -Status $status }
        }
        $bulkCopy.add_SqlRowsCopied($progressEventHandler)

        $columnList = ($transferableColumns | ForEach-Object { "[$_]" }) -join ', '
        $query = if ($SampleSize -gt 0) { "SELECT TOP $SampleSize $columnList FROM [$($src.Schema)].[$($src.Table)]" } else { "SELECT $columnList FROM [$($src.Schema)].[$($src.Table)]" }
        
        Write-DebugMessage "Executing source data query:`n$query"
        
        $srcReaderConn = New-Object System.Data.SqlClient.SqlConnection($srcConnStr)
        $srcReaderConn.Open()
        $cmd = New-Object System.Data.SqlClient.SqlCommand($query, $srcReaderConn)
        $cmd.CommandTimeout = 0
        
        Write-Log -Message "INFO: Source query timeout is set to infinite. This may take a long time for large tables."
        $reader = $cmd.ExecuteReader()
        
        if ($hasIdentityColumn) {
            $identityCmd = $tgtConn.CreateCommand(); $identityCmd.CommandText = "SET IDENTITY_INSERT [$($tgt.Schema)].[$($tgt.Table)] ON"; [void]$identityCmd.ExecuteNonQuery()
            Write-Log -Message "INFO: IDENTITY_INSERT enabled for [$($tgt.Schema)].[$($tgt.Table)]."
        }

        Write-DebugMessage "--- SqlBulkCopy Configuration ---"
        Write-DebugMessage "Destination Table: $($bulkCopy.DestinationTableName)"
        Write-DebugMessage "Batch Size: $($bulkCopy.BatchSize)"; Write-DebugMessage "Timeout: $($bulkCopy.BulkCopyTimeout)"; Write-DebugMessage "Options: $($bulkCopy.Options)"
        Write-DebugMessage "Mapping $($transferableColumns.Count) columns for bulk copy..."
        foreach ($colName in $transferableColumns) {
            # Casting to [void] suppresses the unwanted return object from this method.
            [void]$bulkCopy.ColumnMappings.Add($colName, $colName)
        }
        
        Write-Log -Message "INFO: Data stream opened. Starting bulk copy..."
        try {
            $bulkCopy.WriteToServer($reader)
        } finally {
            Write-Progress -Activity "Streaming data to [$($tgt.Table)]" -Completed
        }

        if ($hasIdentityColumn) {
            $identityCmd = $tgtConn.CreateCommand(); $identityCmd.CommandText = "SET IDENTITY_INSERT [$($tgt.Schema)].[$($tgt.Table)] OFF"; [void]$identityCmd.ExecuteNonQuery()
            Write-Log -Message "INFO: IDENTITY_INSERT disabled for [$($tgt.Schema)].[$($tgt.Table)]."
        }
        Write-Log -Message "SUCCESS: Bulk data stream completed." -ForegroundColor Green

    } catch {
        throw "A failure occurred during the transfer process: $_"
    } finally {
        if ($reader) { $reader.Dispose() }
        if ($srcReaderConn) { $srcReaderConn.Dispose() }
        if ($bulkCopy) { $bulkCopy.Close() }
        if ($tgtConn) { $tgtConn.Dispose() }
    }
}


# Section: Main Execution
# --- Pre-flight Checks ---
Confirm-SqlServerModule
Show-VersionBanner

# --- Main Loop ---
$isFirstRun = $true
do {
    $localSource = if ($isFirstRun) { $Source } else { $null }
    $localTarget = if ($isFirstRun) { $Target } else { $null }
    $localSampleSize = if ($isFirstRun -and $PSBoundParameters.ContainsKey('SampleSize')) { $SampleSize } else { -1 }
    $localBatchSize = if ($isFirstRun -and $PSBoundParameters.ContainsKey('BatchSize')) { $BatchSize } else { -1 }
    $localSchemaOption = if ($isFirstRun -and $PSBoundParameters.ContainsKey('SchemaOption')) { $SchemaOption } else { $null }

    if ([string]::IsNullOrWhiteSpace($localSource)) {
        while ([string]::IsNullOrWhiteSpace($localSource)) {
            Write-Host "Enter the source [Server].[Database].[Schema].[Table]: " -ForegroundColor White -NoNewline; $localSource = Read-Host
            if ([string]::IsNullOrWhiteSpace($localSource)) { Write-Host "Source cannot be empty. Please try again." -ForegroundColor Yellow }
        }
    }
    if ([string]::IsNullOrWhiteSpace($localTarget)) {
        while ([string]::IsNullOrWhiteSpace($localTarget)) {
            Write-Host "Enter the target [Server].[Database].[Schema].[Table]: " -ForegroundColor White -NoNewline; $localTarget = Read-Host
            if ([string]::IsNullOrWhiteSpace($localTarget)) { Write-Host "Target cannot be empty. Please try again." -ForegroundColor Yellow }
        }
    }
    if ($localSampleSize -lt 0) {
        $sampleInput = ''
        while ($sampleInput -notmatch '^\d+$') {
            Write-Host "Enter TOP N sample size (0 for full load) [Default: 0]: " -ForegroundColor White -NoNewline; $sampleInput = Read-Host
            if ([string]::IsNullOrWhiteSpace($sampleInput)) { $sampleInput = '0'; break }
            if ($sampleInput -notmatch '^\d+$') { Write-Host "Invalid input. Please enter a non-negative number." -ForegroundColor Yellow }
        }
        $localSampleSize = [int]$sampleInput
    }
    if ($localBatchSize -lt 0) {
        $batchInput = ''
        while ($batchInput -notmatch '^[1-9]\d*$') {
            Write-Host "Enter the batch size (e.g., 50000) [Default: 50000]: " -ForegroundColor White -NoNewline; $batchInput = Read-Host
            if ([string]::IsNullOrWhiteSpace($batchInput)) { $batchInput = '50000'; break }
            if ($batchInput -notmatch '^[1-9]\d*$') { Write-Host "Invalid input. Please enter a positive number." -ForegroundColor Yellow }
        }
        $localBatchSize = [int]$batchInput
    }
    if ([string]::IsNullOrWhiteSpace($localSchemaOption)) {
        $schemaInput = ''
        while ($schemaInput -notin @('1', '2')) {
            Write-Host "Enter 1 for full schema, 2 for NVARCHAR(MAX) columns [Default: 1]: " -ForegroundColor White -NoNewline; $schemaInput = Read-Host
            if ([string]::IsNullOrWhiteSpace($schemaInput)) { $schemaInput = '1'; break }
            if ($schemaInput -notin @('1', '2')) { Write-Host "Invalid input. Please enter '1' or '2'." -ForegroundColor Yellow }
        }
        $localSchemaOption = $schemaInput
    }

    try {
        Import-Module -Name SqlServer -ErrorAction Stop
        Start-DataTransfer -SourceFourPartName $localSource -TargetFourPartName $localTarget -SchemaOptionToUse $localSchemaOption -SampleSize $localSampleSize -BatchSizeToUse $localBatchSize
    } catch {
        Write-Log -Message "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    }

    $isFirstRun = $false
    Write-Log -Message "`nINFO: Do you want to restart the process? Enter Y to restart, or press Enter to exit"
    $restart = Read-Host
    if ($restart -notmatch '^y$') {
        Write-Log -Message "SUCCESS: Exiting script." -ForegroundColor Green; break
    }
    Write-Log -Message "INFO: Restarting the process..."
} while ($true)