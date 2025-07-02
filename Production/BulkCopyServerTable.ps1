<#
.SYNOPSIS
    Copies a SQL Server table's schema and data from a source to a target server using four-part naming.

.DESCRIPTION
    This script facilitates the migration of a SQL Server table, including its schema and data, from a source to a target destination. It uses four-part naming convention ([Server].[Database].[Schema].[Table]) to identify the source and target.

    The script offers two modes for schema creation on the target:
    1.  Full Schema: Replicates the source table's structure precisely, including data types, primary keys, indexes, IDENTITY properties, and computed columns.
    2.  NVARCHAR(MAX) Schema: Creates a simplified version of the table where all columns are of the NVARCHAR(MAX) data type. This is useful for quick data staging or when the exact schema is not required.

    Key Features:
    - Supports both full table and partial (TOP N rows) data transfer.
    - Handles existing target tables by prompting the user to truncate or append data.
    - Correctly manages IDENTITY_INSERT when replicating tables with IDENTITY columns in full schema mode.
    - Provides robust error handling and color-coded console output for clarity.
    - Includes a -DebugMode switch for detailed execution tracing.

.PARAMETER Source
    The full four-part name of the source table in the format [Server].[Database].[Schema].[Table]. Brackets are optional.

.PARAMETER Target
    The full four-part name of the target table in the format [Server].[Database].[Schema].[Table]. Brackets are optional.

.PARAMETER SampleSize
    Specifies the number of rows to copy (TOP N). A value of 0 (the default) copies all rows.

.PARAMETER SchemaOption
    Determines how the target table's schema is created if it does not exist.
    - '1' (Default): Full schema replication (data types, PK, indexes, IDENTITY, computed columns).
    - '2': Simplified schema where all columns are created as NVARCHAR(MAX).

.PARAMETER DebugMode
    A switch that, when present, enables verbose debugging output. This output includes generated SQL, variable states, and function traces.

.EXAMPLE
    .\BulkCopyServerTable.ps1 -Source "SQL01.SalesDB.dbo.Orders" -Target "SQL02.SalesArchive.dbo.Orders_2024" -SchemaOption 1
    This command copies the entire 'Orders' table from SQL01 to SQL02, replicating the full schema.

.EXAMPLE
    .\BulkCopyServerTable.ps1 -Source "[SQL01].[SalesDB].[dbo].[Customers]" -Target "[SQL_STG].[Staging].[dbo].[Customers]" -SampleSize 1000 -SchemaOption 2
    This command copies the top 1000 rows from the 'Customers' table into a staging table where all columns will be NVARCHAR(MAX).

.EXAMPLE
    .\BulkCopyServerTable.ps1 -DebugMode
    This command runs the script in interactive mode, prompting the user for all inputs, and enables detailed debug logging.
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
    [ValidateSet('1', '2')]
    [string]$SchemaOption = '1',

    [Parameter(Mandatory = $false)]
    [switch]$DebugMode
)

# V19.0
# Purpose:
#   Copies a SQL Server table's schema and data from a source to a target using four-part naming
#   (Server.Database.Schema.Table). Replicates table structure or creates a simplified schema
#   based on user choice, and transfers data efficiently using SqlBulkCopy.
#
# Approach:
#   - Parses four-part names and builds connection strings.
#   - Retrieves source data into a DataTable, with optional TOP N row limits.
#   - Prompts user to choose between full schema (data types, PK, indexes, IDENTITY, computed columns)
#     or simplified schema (all columns, including computed, as NVARCHAR(MAX)).
#   - Checks if target table exists, compares schema, checks for data, and prompts for truncate/insert.
#   - Replicates schema (full or simplified) and transfers data with SqlBulkCopy.
#
# Key Features:
#   - Fully parameterized for automation with interactive fallback.
#   - Get-Help compatible documentation.
#   - -DebugMode switch for verbose execution tracing.
#   - Handles IDENTITY columns with explicit insertion (SET IDENTITY_INSERT, KeepIdentity) for full schema.
#   - Supports computed columns (persisted and non-persisted) with accurate definitions for full schema.
#   - Replicates clustered and nonclustered rowstore indexes for full schema.
#   - Supports a wide range of SQL Server data types, with warnings for deprecated types (e.g., text).
#   - Allows target table in NVARCHAR(MAX) mode to have a subset of source columns, mapping only
#     matching columns for data transfer.
#   - Uses SqlBulkCopy with connection string constructor for reliable data transfer.
#   - Supports recursive execution with user prompt to restart (Y) or exit (Enter) after each run.
#
# Limitations:
#   - Does not support columnstore indexes, foreign keys, or full user-defined types (UDTs)/CLR types.
#   - Requires SQL Server 2017+ due to STRING_AGG usage.
#   - IDENTITY_INSERT requires non-null IDENTITY values in source data (full schema only).
#   - In NVARCHAR(MAX) mode, data for source columns not present in the target table is skipped.
#   - Computed columns are materialized as NVARCHAR(MAX) in simplified schema mode, losing their computed logic.

# Function: Write-DebugMessage
# Purpose: Writes a magenta-colored debug message to the console if DebugMode is enabled.
function Write-DebugMessage {
    param(
        [string]$Message
    )
    if ($DebugMode) {
        Write-Host "[DEBUG] $Message" -ForegroundColor Magenta
    }
}

# Function: Parse-FourPartName
# Purpose: Parses a four-part table name (Server.Database.Schema.Table) into components,
#          handling bracketed or unbracketed names and validating syntax.
function Parse-FourPartName {
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$InputName
    )
    Write-DebugMessage "Parsing four-part name: $InputName"

    # Validate bracket balance to ensure proper syntax
    $openBrackets = ($InputName.ToCharArray() | Where-Object { $_ -eq '[' }).Count
    $closeBrackets = ($InputName.ToCharArray() | Where-Object { $_ -eq ']' }).Count
    if ($openBrackets -ne $closeBrackets) {
        throw "Unbalanced brackets in input: $InputName"
    }

    # Split on dots outside brackets using regex
    $pattern = '(?<!\[[^\]]*)\.(?![^\[]*\])'
    $parts = [regex]::Split($InputName, $pattern)
    if ($parts.Count -ne 4) {
        throw "Input must be in format Server.Database.Schema.Table — bracketed or unbracketed. Got: $InputName"
    }

    # Remove brackets from each part and validate no empty parts
    $cleanParts = $parts | ForEach-Object { $_ -replace '^\[|\]$', '' }
    if ($cleanParts -contains '') {
        throw "Empty parts are not allowed in input: $InputName"
    }

    # Return a hashtable with the parsed components
    $result = @{
        Server   = $cleanParts[0]
        Database = $cleanParts[1]
        Schema   = $cleanParts[2]
        Table    = $cleanParts[3]
    }
    Write-DebugMessage "Parsed result: Server=$($result.Server), Database=$($result.Database), Schema=$($result.Schema), Table=$($result.Table)"
    return $result
}

# Function: Get-ConnectionString
# Purpose: Builds a SQL Server connection string from a four-part name,
#          using Integrated Security and a 15-second connection timeout.
function Get-ConnectionString {
    param (
        [ValidateNotNullOrEmpty()]
        [string]$FourPartName
    )
    $parsed = Parse-FourPartName $FourPartName
    $connStr = "Server=$($parsed.Server);Database=$($parsed.Database);Integrated Security=True;Connect Timeout=15"
    Write-DebugMessage "Generated Connection String for $($parsed.Server): $connStr"
    return $connStr
}

# Function: Get-TableSchema
# Purpose: Retrieves schema metadata for a table, including columns, data types, constraints,
#          computed columns, IDENTITY properties, and indexes.
function Get-TableSchema {
    param (
        [ValidateNotNullOrEmpty()]
        [string]$FourPartName,
        [System.Data.SqlClient.SqlConnection]$Connection
    )
    $parsed = Parse-FourPartName $FourPartName
    Write-DebugMessage "Getting table schema for [$($parsed.Schema)].[$($parsed.Table)]"

    # Query column metadata
    $colCmd = $Connection.CreateCommand()
    $colCmd.CommandText = @"
SELECT
    c.COLUMN_NAME,
    c.DATA_TYPE,
    c.CHARACTER_MAXIMUM_LENGTH,
    c.NUMERIC_PRECISION,
    c.NUMERIC_SCALE,
    c.DATETIME_PRECISION,
    c.IS_NULLABLE,
    CASE WHEN kcu.COLUMN_NAME IS NOT NULL THEN 1 ELSE 0 END AS IS_PRIMARY_KEY,
    CASE WHEN cc.definition IS NOT NULL THEN 1 ELSE 0 END AS IS_COMPUTED,
    cc.definition AS COMPUTED_DEFINITION,
    cc.is_persisted AS IS_PERSISTED,
    CASE WHEN sc.is_identity = 1 THEN 1 ELSE 0 END AS IS_IDENTITY,
    ic.seed_value AS IDENTITY_SEED,
    ic.increment_value AS IDENTITY_INCREMENT
FROM [$($parsed.Database)].INFORMATION_SCHEMA.COLUMNS c
JOIN [$($parsed.Database)].sys.tables t ON c.TABLE_NAME = t.name
JOIN [$($parsed.Database)].sys.schemas s ON t.schema_id = s.schema_id
JOIN [$($parsed.Database)].sys.columns sc ON t.object_id = sc.object_id AND c.COLUMN_NAME = sc.name
LEFT JOIN [$($parsed.Database)].sys.identity_columns ic ON sc.object_id = ic.object_id AND sc.name = ic.name
LEFT JOIN [$($parsed.Database)].INFORMATION_SCHEMA.KEY_COLUMN_USAGE kcu
    ON c.TABLE_SCHEMA = kcu.TABLE_SCHEMA
    AND c.TABLE_NAME = kcu.TABLE_NAME
    AND c.COLUMN_NAME = kcu.COLUMN_NAME
    AND kcu.CONSTRAINT_NAME LIKE 'PK_%'
LEFT JOIN [$($parsed.Database)].sys.computed_columns cc
    ON c.TABLE_SCHEMA = OBJECT_SCHEMA_NAME(cc.object_id)
    AND c.TABLE_NAME = OBJECT_NAME(cc.object_id)
    AND c.COLUMN_NAME = cc.name
WHERE c.TABLE_SCHEMA = @Schema AND c.TABLE_NAME = @Table
ORDER BY c.ORDINAL_POSITION
"@
    $colCmd.Parameters.AddWithValue("@Schema", $parsed.Schema) | Out-Null
    $colCmd.Parameters.AddWithValue("@Table", $parsed.Table) | Out-Null
    Write-DebugMessage "Executing schema query for table: $($parsed.Table)"
    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $colCmd
    $schemaTable = New-Object System.Data.DataTable
    $adapter.Fill($schemaTable) | Out-Null

    # Query index metadata
    $idxCmd = $Connection.CreateCommand()
    $idxCmd.CommandText = @"
SELECT
    i.name AS IndexName,
    i.is_unique AS IsUnique,
    i.type AS IndexType,
    CASE WHEN i.type = 1 THEN 'CLUSTERED' ELSE 'NONCLUSTERED' END AS IndexKind,
    STRING_AGG(CASE WHEN ic.is_included_column = 0 THEN c.name ELSE NULL END, ',') AS KeyColumns,
    STRING_AGG(CASE WHEN ic.is_included_column = 1 THEN c.name ELSE NULL END, ',') AS IncludedColumns
FROM [$($parsed.Database)].sys.indexes i
JOIN [$($parsed.Database)].sys.index_columns ic ON i.object_id = ic.object_id AND i.index_id = ic.index_id
JOIN [$($parsed.Database)].sys.columns c ON ic.object_id = c.object_id AND ic.column_id = c.column_id
JOIN [$($parsed.Database)].sys.tables t ON i.object_id = t.object_id
JOIN [$($parsed.Database)].sys.schemas s ON t.schema_id = s.schema_id
WHERE s.name = @Schema AND t.name = @Table
    AND i.is_primary_key = 0 AND i.type IN (1, 2)
GROUP BY i.name, i.is_unique, i.type
HAVING STRING_AGG(CASE WHEN ic.is_included_column = 0 THEN c.name ELSE NULL END, ',') IS NOT NULL
"@
    $idxCmd.Parameters.AddWithValue("@Schema", $parsed.Schema) | Out-Null
    $idxCmd.Parameters.AddWithValue("@Table", $parsed.Table) | Out-Null
    Write-DebugMessage "Executing index query for table: $($parsed.Table)"
    $idxAdapter = New-Object System.Data.SqlClient.SqlDataAdapter $idxCmd
    $indexTable = New-Object System.Data.DataTable
    $idxAdapter.Fill($indexTable) | Out-Null

    # Check if primary key is clustered
    $pkCmd = $Connection.CreateCommand()
    $pkCmd.CommandText = @"
SELECT
    CASE WHEN i.type = 1 THEN 1 ELSE 0 END AS IsClustered
FROM [$($parsed.Database)].sys.indexes i
JOIN [$($parsed.Database)].sys.tables t ON i.object_id = t.object_id
JOIN [$($parsed.Database)].sys.schemas s ON t.schema_id = s.schema_id
WHERE s.name = @Schema AND t.name = @Table AND i.is_primary_key = 1
"@
    $pkCmd.Parameters.AddWithValue("@Schema", $parsed.Schema) | Out-Null
    $pkCmd.Parameters.AddWithValue("@Table", $parsed.Table) | Out-Null
    $pkIsClustered = $pkCmd.ExecuteScalar()
    if ($null -eq $pkIsClustered) { $pkIsClustered = 0 }

    Write-DebugMessage "Found $($schemaTable.Rows.Count) columns, $($indexTable.Rows.Count) indexes. PK Clustered: $pkIsClustered"

    return @{
        SchemaTable = $schemaTable
        IndexTable = $indexTable
        PkIsClustered = $pkIsClustered
    }
}

# Function: Compare-TableSchemas
# Purpose: Compares source and target table schemas for compatibility based on schema option.
function Compare-TableSchemas {
    param (
        [System.Data.DataTable]$SourceSchema,
        [System.Data.DataTable]$TargetSchema,
        [System.Data.DataTable]$SourceIndexes,
        [System.Data.DataTable]$TargetIndexes,
        [int]$SourcePkIsClustered,
        [int]$TargetPkIsClustered,
        [string]$SchemaOptionToCompare
    )
    Write-DebugMessage "Comparing schemas with SchemaOption: $SchemaOptionToCompare"

    if ($SchemaOptionToCompare -eq '2') {
        # NVARCHAR(MAX) mode: Check target columns are a subset of source and NVARCHAR(MAX)
        Write-DebugMessage "Comparing in NVARCHAR(MAX) mode."
        $sourceCols = $SourceSchema.Rows | Select-Object -ExpandProperty COLUMN_NAME | Sort-Object
        $targetCols = $TargetSchema.Rows | Select-Object -ExpandProperty COLUMN_NAME | Sort-Object
        $missingCols = $sourceCols | Where-Object { $targetCols -notcontains $_ }
        
        foreach ($tgtCol in $targetCols) {
            if ($sourceCols -notcontains $tgtCol) {
                return "Target column [$tgtCol] does not exist in source table."
            }
        }
        
        foreach ($row in $TargetSchema.Rows) {
            if ($row.DATA_TYPE -ne 'nvarchar' -or $row.CHARACTER_MAXIMUM_LENGTH -ne -1) {
                return "Target column [$($row.COLUMN_NAME)] is not NVARCHAR(MAX)."
            }
            if ($row.IS_IDENTITY -eq 1) {
                return "Target column [$($row.COLUMN_NAME)] has IDENTITY, which is not allowed in NVARCHAR(MAX) mode."
            }
            if ($row.IS_COMPUTED -eq 1) {
                return "Target column [$($row.COLUMN_NAME)] is computed, which is not allowed in NVARCHAR(MAX) mode."
            }
            if ($row.IS_PRIMARY_KEY -eq 1) {
                return "Target column [$($row.COLUMN_NAME)] is part of a primary key, which is not allowed in NVARCHAR(MAX) mode."
            }
        }
        if ($TargetIndexes.Rows.Count -gt 0) {
            return "Target table has indexes, which are not allowed in NVARCHAR(MAX) mode."
        }
        Write-DebugMessage "NVARCHAR(MAX) schema is compatible. Missing source columns in target: $($missingCols -join ', ')"
        return @{
            IsCompatible = $true
            MissingColumns = $missingCols
        } # Schemas compatible, with possible missing columns
    } else {
        # Full schema mode: Exact match
        Write-DebugMessage "Comparing in Full Schema mode."
        if ($SourceSchema.Rows.Count -ne $TargetSchema.Rows.Count) {
            return "Column count mismatch: Source ($($SourceSchema.Rows.Count)) vs Target ($($TargetSchema.Rows.Count))."
        }
        $sourceCols = $SourceSchema.Rows | Sort-Object COLUMN_NAME
        $targetCols = $TargetSchema.Rows | Sort-Object COLUMN_NAME
        for ($i = 0; $i -lt $sourceCols.Count; $i++) {
            $src = $sourceCols[$i]
            $tgt = $targetCols[$i]
            if ($src.COLUMN_NAME -ne $tgt.COLUMN_NAME) { return "Column name mismatch: Source [$($src.COLUMN_NAME)] vs Target [$($tgt.COLUMN_NAME)]." }
            if ($src.DATA_TYPE -ne $tgt.DATA_TYPE) { return "Data type mismatch for [$($src.COLUMN_NAME)]: Source [$($src.DATA_TYPE)] vs Target [$($tgt.DATA_TYPE)]." }
            if ($src.CHARACTER_MAXIMUM_LENGTH -ne $tgt.CHARACTER_MAXIMUM_LENGTH) { return "Length mismatch for [$($src.COLUMN_NAME)]: Source [$($src.CHARACTER_MAXIMUM_LENGTH)] vs Target [$($tgt.CHARACTER_MAXIMUM_LENGTH)]." }
            if ($src.NUMERIC_PRECISION -ne $tgt.NUMERIC_PRECISION) { return "Precision mismatch for [$($src.COLUMN_NAME)]: Source [$($src.NUMERIC_PRECISION)] vs Target [$($tgt.NUMERIC_PRECISION)]." }
            if ($src.NUMERIC_SCALE -ne $tgt.NUMERIC_SCALE) { return "Scale mismatch for [$($src.COLUMN_NAME)]: Source [$($src.NUMERIC_SCALE)] vs Target [$($tgt.NUMERIC_SCALE)]." }
            if ($src.DATETIME_PRECISION -ne $tgt.DATETIME_PRECISION) { return "Datetime precision mismatch for [$($src.COLUMN_NAME)]: Source [$($src.DATETIME_PRECISION)] vs Target [$($tgt.DATETIME_PRECISION)]." }
            if ($src.IS_NULLABLE -ne $tgt.IS_NULLABLE) { return "Nullability mismatch for [$($src.COLUMN_NAME)]: Source [$($src.IS_NULLABLE)] vs Target [$($tgt.IS_NULLABLE)]." }
            if ($src.IS_PRIMARY_KEY -ne $tgt.IS_PRIMARY_KEY) { return "Primary key status mismatch for [$($src.COLUMN_NAME)]: Source [$($src.IS_PRIMARY_KEY)] vs Target [$($tgt.IS_PRIMARY_KEY)]." }
            if ($src.IS_COMPUTED -ne $tgt.IS_COMPUTED) { return "Computed column status mismatch for [$($src.COLUMN_NAME)]: Source [$($src.IS_COMPUTED)] vs Target [$($tgt.IS_COMPUTED)]." }
            if ($src.IS_COMPUTED -eq 1) {
                if ($src.COMPUTED_DEFINITION -ne $tgt.COMPUTED_DEFINITION) { return "Computed column definition mismatch for [$($src.COLUMN_NAME)]." }
                if ($src.IS_PERSISTED -ne $tgt.IS_PERSISTED) { return "Computed column persisted status mismatch for [$($src.COLUMN_NAME)]: Source [$($src.IS_PERSISTED)] vs Target [$($src.IS_PERSISTED)]." }
            }
            if ($src.IS_IDENTITY -ne $tgt.IS_IDENTITY) { return "IDENTITY status mismatch for [$($src.COLUMN_NAME)]: Source [$($src.IS_IDENTITY)] vs Target [$($tgt.IS_IDENTITY)]." }
            if ($src.IS_IDENTITY -eq 1) {
                if ($src.IDENTITY_SEED -ne $tgt.IDENTITY_SEED -or $src.IDENTITY_INCREMENT -ne $tgt.IDENTITY_INCREMENT) { return "IDENTITY properties mismatch for [$($src.COLUMN_NAME)]: Source [IDENTITY($($src.IDENTITY_SEED),$($src.IDENTITY_INCREMENT))] vs Target [IDENTITY($($tgt.IDENTITY_SEED),$($tgt.IDENTITY_INCREMENT))]." }
            }
        }
        if ($SourcePkIsClustered -ne $TargetPkIsClustered) { return "Primary key clustering mismatch: Source [IsClustered=$SourcePkIsClustered] vs Target [IsClustered=$TargetPkIsClustered]." }
        if ($SourceIndexes.Rows.Count -ne $TargetIndexes.Rows.Count) { return "Index count mismatch: Source ($($SourceIndexes.Rows.Count)) vs Target ($($TargetIndexes.Rows.Count))." }
        
        $sourceIdx = $SourceIndexes.Rows | Sort-Object KeyColumns, IncludedColumns
        $targetIdx = $TargetIndexes.Rows | Sort-Object KeyColumns, IncludedColumns
        for ($i = 0; $i -lt $sourceIdx.Count; $i++) {
            $src = $sourceIdx[$i]
            $tgt = $targetIdx[$i]
            if ($src.IndexType -ne $tgt.IndexType) { return "Index type mismatch for index: Source [Type=$($src.IndexType)] vs Target [Type=$($tgt.IndexType)]." }
            if ($src.IsUnique -ne $tgt.IsUnique) { return "Index uniqueness mismatch for index: Source [IsUnique=$($src.IsUnique)] vs Target [IsUnique=$($tgt.IsUnique)]." }
            if ($src.KeyColumns -ne $tgt.KeyColumns) { return "Index key columns mismatch: Source [$($src.KeyColumns)] vs Target [$($tgt.KeyColumns)]." }
            if ($src.IncludedColumns -ne $tgt.IncludedColumns) { return "Index included columns mismatch: Source [$($src.IncludedColumns)] vs Target [$($tgt.IncludedColumns)]." }
        }
        Write-DebugMessage "Full schema comparison passed. Schemas match."
        return $null # Schemas match
    }
}

# Function: Get-DataTable
# Purpose: Retrieves data from the source table into a DataTable, optionally limiting rows
#          with TOP N, for use in bulk data insertion.
function Get-DataTable {
    param (
        [ValidateNotNullOrEmpty()]
        [string]$FourPartName,
        [int]$Top = 0
    )
    # Parse the input name and get the connection string
    $parsed = Parse-FourPartName $FourPartName
    $connStr = Get-ConnectionString $FourPartName

    # Build the SELECT query, adding TOP N if specified
    $conn = New-Object System.Data.SqlClient.SqlConnection $connStr
    $query = if ($Top -gt 0) {
        "SELECT TOP $Top * FROM [$($parsed.Schema)].[$($parsed.Table)]"
    } else {
        "SELECT * FROM [$($parsed.Schema)].[$($parsed.Table)]"
    }
    Write-DebugMessage "Executing data retrieval query: $query"

    # Execute the query and fill a DataTable
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $query
    $cmd.CommandTimeout = 30
    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
    $table = New-Object System.Data.DataTable
    try {
        $conn.Open()
        [void]$adapter.Fill($table)
    } catch {
        throw "Failed to fetch data from [$($parsed.Schema)].[$($parsed.Table)]: $_"
    } finally {
        $conn.Close()
        $conn.Dispose()
    }
    Write-DebugMessage "Retrieved $($table.Rows.Count) rows from source."
    return $table
}

# Function: BulkInsert-DataTable
# Purpose: Replicates the source table's schema (full or NVARCHAR(MAX)) and performs bulk data
#          insertion, handling existing tables, schema comparison, and truncate/insert options.
function BulkInsert-DataTable {
    param (
        [ValidateNotNullOrEmpty()]
        [string]$SourceFourPartName,
        [ValidateNotNullOrEmpty()]
        [string]$TargetFourPartName,
        [System.Data.DataTable]$DataTable,
        [string]$SchemaOptionToUse
    )
    # Parse source and target four-part names and get connection strings
    $src = Parse-FourPartName $SourceFourPartName
    $tgt = Parse-FourPartName $TargetFourPartName
    $srcConnStr = Get-ConnectionString $SourceFourPartName
    $tgtConnStr = Get-ConnectionString $TargetFourPartName

    # Fetch source schema
    $srcConn = New-Object System.Data.SqlClient.SqlConnection $srcConnStr
    try {
        $srcConn.Open()
        $srcSchema = Get-TableSchema -FourPartName $SourceFourPartName -Connection $srcConn
        $schemaTable = $srcSchema.SchemaTable
        $indexTable = $srcSchema.IndexTable
        $pkIsClustered = $srcSchema.PkIsClustered
        if ($schemaTable.Rows.Count -eq 0) {
            throw "No columns found for source table [$($src.Schema)].[$($src.Table)]"
        }
    } catch {
        throw "Failed to fetch schema metadata for [$($src.Schema)].[$($src.Table)]: $_"
    } finally {
        $srcConn.Close()
        $srcConn.Dispose()
    }

    # Adjust DataTable: Keep only columns from source schema
    $validCols = $schemaTable.Rows | Select-Object -ExpandProperty COLUMN_NAME
    $colsToRemove = $DataTable.Columns | Select-Object -ExpandProperty ColumnName | Where-Object { $validCols -notcontains $_ }
    if ($colsToRemove) {
        Write-DebugMessage "Removing columns from in-memory DataTable that are not in source schema: $($colsToRemove -join ', ')"
        foreach ($col in $colsToRemove) {
            $DataTable.Columns.Remove($col)
        }
    }

    # Check for IDENTITY and computed columns
    $identityCols = $schemaTable.Rows | Where-Object { $_.IS_IDENTITY -eq 1 } | Select-Object -ExpandProperty COLUMN_NAME
    $computedCols = $schemaTable.Rows | Where-Object { $_.IS_COMPUTED -eq 1 } | Select-Object -ExpandProperty COLUMN_NAME
    $useIdentityInsert = $false
    if ($SchemaOptionToUse -eq '1' -and $identityCols.Count -gt 0) {
        foreach ($identityCol in $identityCols) {
            if (-not $DataTable.Columns.Contains($identityCol)) {
                Write-Host "WARNING: IDENTITY column [$identityCol] missing in source data. Proceeding without IDENTITY_INSERT." -ForegroundColor Yellow
            } else {
                $hasValidValues = $true
                foreach ($row in $DataTable.Rows) {
                    if ($row[$identityCol] -eq [DBNull]::Value -or $null -eq $row[$identityCol]) {
                        $hasValidValues = $false
                        break
                    }
                }
                if (-not $hasValidValues) {
                    Write-Host "WARNING: IDENTITY column [$identityCol] contains null or invalid values. Proceeding without IDENTITY_INSERT." -ForegroundColor Yellow
                } else {
                    $useIdentityInsert = $true
                    Write-DebugMessage "IDENTITY column [$identityCol] has valid data. Will use IDENTITY_INSERT."
                }
            }
        }
        if ($computedCols.Count -gt 0) {
            Write-Host "WARNING: Computed columns [$($computedCols -join ', ')] will have their values computed in the target and are excluded from data insertion." -ForegroundColor Yellow
        }
    } elseif ($SchemaOptionToUse -eq '2') {
        if ($identityCols.Count -gt 0) {
            Write-Host "WARNING: IDENTITY columns [$($identityCols -join ', ')] will be NVARCHAR(MAX) in target." -ForegroundColor Yellow
        }
        if ($computedCols.Count -gt 0) {
            Write-Host "WARNING: Computed columns [$($computedCols -join ', ')] will be materialized as NVARCHAR(MAX) in target." -ForegroundColor Yellow
        }
    }

    # Connect to target database
    $tgtConn = New-Object System.Data.SqlClient.SqlConnection $tgtConnStr
    try {
        $tgtConn.Open()
        Write-DebugMessage "Successfully connected to target server: $($tgt.Server)"

        # Check if target table exists
        $checkCmd = $tgtConn.CreateCommand()
        $checkCmd.CommandText = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = @Schema AND TABLE_NAME = @Table"
        $checkCmd.Parameters.AddWithValue("@Schema", $tgt.Schema) | Out-Null
        $checkCmd.Parameters.AddWithValue("@Table", $tgt.Table) | Out-Null
        $exists = $checkCmd.ExecuteScalar()

        if ($exists -gt 0) {
            Write-Host "INFO: Target table [$($tgt.Schema)].[$($tgt.Table)] exists. Checking schema compatibility..." -ForegroundColor Cyan

            # Fetch target schema
            $tgtSchema = Get-TableSchema -FourPartName $TargetFourPartName -Connection $tgtConn
            $tgtSchemaTable = $tgtSchema.SchemaTable
            $tgtIndexTable = $tgtSchema.IndexTable
            $tgtPkIsClustered = $tgtSchema.PkIsClustered

            # Compare schemas
            $schemaCheck = Compare-TableSchemas -SourceSchema $schemaTable -TargetSchema $tgtSchemaTable `
                -SourceIndexes $indexTable -TargetIndexes $tgtIndexTable `
                -SourcePkIsClustered $pkIsClustered -TargetPkIsClustered $tgtPkIsClustered `
                -SchemaOptionToCompare $SchemaOptionToUse
            
            if ($SchemaOptionToUse -eq '2') {
                if (-not $schemaCheck.IsCompatible) {
                    Write-Host "WARNING: Schema mismatch between source and target: $($schemaCheck)" -ForegroundColor Yellow
                    throw "Cannot proceed due to incompatible target table schema."
                }
                if ($schemaCheck.MissingColumns) {
                    Write-Host "WARNING: Target table [$($tgt.Schema)].[$($tgt.Table)] is missing columns [$($schemaCheck.MissingColumns -join ', ')]. Data for these columns will be skipped." -ForegroundColor Yellow
                }
                Write-Host "INFO: Target table schema is compatible for option 2 (NVARCHAR(MAX) mode)." -ForegroundColor Cyan
            } else {
                if ($schemaCheck) {
                    Write-Host "WARNING: Schema mismatch between source and target: $schemaCheck" -ForegroundColor Yellow
                    throw "Cannot proceed due to incompatible target table schema."
                }
                Write-Host "INFO: Target table schema matches expected schema for option 1." -ForegroundColor Cyan
            }

            # Check for IDENTITY columns in target (full schema only)
            if ($SchemaOptionToUse -eq '1') {
                $tgtIdentityCols = $tgtSchemaTable.Rows | Where-Object { $_.IS_IDENTITY -eq 1 } | Select-Object -ExpandProperty COLUMN_NAME
                if ($tgtIdentityCols.Count -eq 0 -and $useIdentityInsert) {
                    Write-Host "WARNING: Target table [$($tgt.Schema)].[$($tgt.Table)] has no IDENTITY columns. Proceeding without IDENTITY_INSERT." -ForegroundColor Yellow
                    $useIdentityInsert = $false
                }
            }

            # Check for data in target table
            $dataCmd = $tgtConn.CreateCommand()
            $dataCmd.CommandText = "SELECT COUNT(*) FROM [$($tgt.Schema)].[$($tgt.Table)]"
            $rowCount = $dataCmd.ExecuteScalar()
            if ($rowCount -gt 0) {
                Write-Host "INFO: Target table contains $rowCount row(s). Prompting for action..." -ForegroundColor Cyan
                Write-Host "Target table contains data. Enter 1 to truncate, 2 to insert additional records: " -ForegroundColor White -NoNewline
                $choice = Read-Host
                if ($choice -eq '1') {
                    Write-Host "INFO: User chose to truncate table [$($tgt.Schema)].[$($tgt.Table)]." -ForegroundColor Cyan
                    $truncateCmd = $tgtConn.CreateCommand()
                    $truncateCmd.CommandText = "TRUNCATE TABLE [$($tgt.Schema)].[$($tgt.Table)]"
                    Write-DebugMessage "Executing TRUNCATE TABLE statement."
                    [void]$truncateCmd.ExecuteNonQuery()
                    Write-Host "SUCCESS: Table [$($tgt.Schema)].[$($tgt.Table)] truncated." -ForegroundColor Green
                } else {
                    Write-Host "INFO: User chose to insert additional records into [$($tgt.Schema)].[$($tgt.Table)]." -ForegroundColor Cyan
                }
            } else {
                Write-Host "INFO: Target table is empty. Proceeding with data insertion." -ForegroundColor Cyan
            }
        } else {
            Write-Host "INFO: Target table [$($tgt.Schema)].[$($tgt.Table)] does not exist. Creating new table with schema option $SchemaOptionToUse..." -ForegroundColor Cyan

            if ($SchemaOptionToUse -eq '2') {
                # NVARCHAR(MAX) mode: Create table with all columns as NVARCHAR(MAX)
                $cols = @()
                foreach ($row in $schemaTable.Rows) {
                    $name = $row.COLUMN_NAME
                    $cols += "[$name] NVARCHAR(MAX) NULL"
                }
                if ($cols.Count -eq 0) { throw "No valid column definitions found for [$($tgt.Schema)].[$($tgt.Table)]." }
                $createSql = "CREATE TABLE [$($tgt.Schema)].[$($tgt.Table)] (`n" + ($cols -join ",`n") + "`n)"
                Write-DebugMessage "Executing CREATE TABLE statement (NVARCHAR MAX):`n$createSql"
                $createCmd = $tgtConn.CreateCommand()
                $createCmd.CommandText = $createSql
                [void]$createCmd.ExecuteNonQuery()
                Write-Host "SUCCESS: Table [$($tgt.Schema)].[$($tgt.Table)] created with all columns as NVARCHAR(MAX)." -ForegroundColor Green
            } else {
                # Full schema mode
                $hasClusteredIndex = ($indexTable.Rows | Where-Object { $_.IndexType -eq 1 } | Select-Object -First 1) -or ($pkIsClustered -eq 1)
                $primaryKeyType = if ($pkIsClustered -eq 1) { 'CLUSTERED' } elseif ($hasClusteredIndex) { 'NONCLUSTERED' } else { 'CLUSTERED' }

                $cols = @()
                $pkCols = @()
                foreach ($row in $schemaTable.Rows) {
                    $name = $row.COLUMN_NAME
                    if ($row.IS_COMPUTED -eq 1) {
                        if ([string]::IsNullOrWhiteSpace($row.COMPUTED_DEFINITION)) { throw "Invalid computed column definition for [$name] in [$($src.Schema)].[$($src.Table)]." }
                        $persisted = if ($row.IS_PERSISTED -eq $true) { ' PERSISTED' } else { '' }
                        $cols += "[$name] AS $($row.COMPUTED_DEFINITION)$persisted"
                    } else {
                        $type = switch ($row.DATA_TYPE) {
                            'nvarchar'      { if ($row.CHARACTER_MAXIMUM_LENGTH -gt 0) { "NVARCHAR($($row.CHARACTER_MAXIMUM_LENGTH))" } else { "NVARCHAR(MAX)" } }
                            'varchar'       { if ($row.CHARACTER_MAXIMUM_LENGTH -gt 0) { "VARCHAR($($row.CHARACTER_MAXIMUM_LENGTH))" } else { "VARCHAR(MAX)" } }
                            'char'          { "CHAR($($row.CHARACTER_MAXIMUM_LENGTH))" }
                            'nchar'         { "NCHAR($($row.CHARACTER_MAXIMUM_LENGTH))" }
                            'decimal'       { "DECIMAL($($row.NUMERIC_PRECISION),$($row.NUMERIC_SCALE))" }
                            'numeric'       { "NUMERIC($($row.NUMERIC_PRECISION),$($row.NUMERIC_SCALE))" }
                            'varbinary'     { if ($row.CHARACTER_MAXIMUM_LENGTH -gt 0) { "VARBINARY($($row.CHARACTER_MAXIMUM_LENGTH))" } else { "VARBINARY(MAX)" } }
                            'binary'        { if ($row.CHARACTER_MAXIMUM_LENGTH -gt 0) { "BINARY($($row.CHARACTER_MAXIMUM_LENGTH))" } else { "BINARY(MAX)" } }
                            'datetime2'     { if ($row.DATETIME_PRECISION -gt 0) { "DATETIME2($($row.DATETIME_PRECISION))" } else { "DATETIME2" } }
                            'time'          { if ($row.DATETIME_PRECISION -gt 0) { "TIME($($row.DATETIME_PRECISION))" } else { "TIME" } }
                            'datetimeoffset'{ if ($row.DATETIME_PRECISION -gt 0) { "DATETIMEOFFSET($($row.DATETIME_PRECISION))" } else { "DATETIMEOFFSET" } }
                            'float'         { if ($row.NUMERIC_PRECISION -gt 0) { "FLOAT($($row.NUMERIC_PRECISION))" } else { "FLOAT" } }
                            'text'          { Write-Host "WARNING: Deprecated data type 'text' used in [$($src.Schema)].[$($src.Table)]." -ForegroundColor Yellow; 'TEXT' }
                            'ntext'         { Write-Host "WARNING: Deprecated data type 'ntext' used in [$($src.Schema)].[$($src.Table)]." -ForegroundColor Yellow; 'NTEXT' }
                            'image'         { Write-Host "WARNING: Deprecated data type 'image' used in [$($src.Schema)].[$($src.Table)]." -ForegroundColor Yellow; 'IMAGE' }
                            default         {
                                if (($row.DATA_TYPE -in ('int', 'bigint', 'smallint', 'tinyint', 'real', 'money', 'smallmoney', 'date', 'datetime', 'smalldatetime', 'uniqueidentifier', 'xml', 'geometry', 'geography', 'hierarchyid', 'timestamp', 'rowversion', 'sql_variant'))) {
                                    $row.DATA_TYPE.ToUpper()
                                } else {
                                    Write-Host "WARNING: Potentially unsupported data type '$($row.DATA_TYPE)'. Using as-is." -ForegroundColor Yellow
                                    $row.DATA_TYPE.ToUpper()
                                }
                            }
                        }
                        $identity = ''
                        if ($row.IS_IDENTITY -eq 1) {
                            if ($null -ne $row.IDENTITY_SEED -and $null -ne $row.IDENTITY_INCREMENT) {
                                $identity = " IDENTITY($($row.IDENTITY_SEED),$($row.IDENTITY_INCREMENT))"
                            } else {
                                Write-Host "WARNING: IDENTITY column [$name] has missing seed/increment values. Defaulting to IDENTITY(1,1)." -ForegroundColor Yellow
                                $identity = " IDENTITY(1,1)"
                            }
                        }
                        $nullability = if ($row.IS_NULLABLE -eq 'YES') { 'NULL' } else { 'NOT NULL' }
                        $cols += "[$name] $type$identity $nullability"
                    }
                    if ($row.IS_PRIMARY_KEY -eq 1) {
                        $pkCols += "[$name]"
                    }
                }
                if ($cols.Count -eq 0) { throw "No valid column definitions found for [$($tgt.Schema)].[$($tgt.Table)]." }
                $createSql = "CREATE TABLE [$($tgt.Schema)].[$($tgt.Table)] (`n" + ($cols -join ",`n")
                if ($pkCols.Count -gt 0) {
                    $createSql += ",`nCONSTRAINT [PK_$($tgt.Table)] PRIMARY KEY $primaryKeyType (" + ($pkCols -join ', ') + ")"
                }
                $createSql += "`n)"
                Write-DebugMessage "Executing CREATE TABLE statement (Full Schema):`n$createSql"
                $createCmd = $tgtConn.CreateCommand()
                $createCmd.CommandText = $createSql
                [void]$createCmd.ExecuteNonQuery()
                Write-Host "SUCCESS: Table [$($tgt.Schema)].[$($tgt.Table)] created with full schema." -ForegroundColor Green

                # Create indexes
                foreach ($idx in $indexTable.Rows) {
                    $idxName = $idx.IndexName; $isUnique = $idx.IsUnique; $indexKind = $idx.IndexKind; $keyColumns = $idx.KeyColumns; $includedColumns = $idx.IncludedColumns
                    if ($indexKind -eq 'CLUSTERED' -and $hasClusteredIndex) {
                        Write-Host "INFO: Skipping clustered index [$idxName] as one already exists." -ForegroundColor Cyan
                        continue
                    }
                    $idxSql = "CREATE "; if ($isUnique -eq $true) { $idxSql += "UNIQUE " }; $idxSql += "$indexKind INDEX [$idxName] ON [$($tgt.Schema)].[$($tgt.Table)] ($keyColumns)"
                    if (-not [string]::IsNullOrWhiteSpace($includedColumns)) { $idxSql += " INCLUDE ($includedColumns)" }
                    Write-DebugMessage "Executing CREATE INDEX statement:`n$idxSql"
                    $idxCmd = $tgtConn.CreateCommand()
                    $idxCmd.CommandText = $idxSql
                    [void]$idxCmd.ExecuteNonQuery()
                    Write-Host "SUCCESS: Index [$idxName] ($indexKind) created on [$($tgt.Schema)].[$($tgt.Table)]." -ForegroundColor Green
                }
            }
        }

        # Perform bulk data insertion
        $bulkCopyOptions = if ($SchemaOptionToUse -eq '1' -and $useIdentityInsert) { [System.Data.SqlClient.SqlBulkCopyOptions]::KeepIdentity } else { [System.Data.SqlClient.SqlBulkCopyOptions]::Default }
        $bulkCopy = New-Object System.Data.SqlClient.SqlBulkCopy($tgtConnStr, $bulkCopyOptions)
        $bulkCopy.DestinationTableName = "[$($tgt.Schema)].[$($tgt.Table)]"
        $bulkCopy.BatchSize = 10000
        $bulkCopy.BulkCopyTimeout = 600
        Write-DebugMessage "SqlBulkCopy options: $bulkCopyOptions"

        # Map columns: Use non-computed columns for Option 1, target table columns for Option 2
        $columnsToMap = if ($SchemaOptionToUse -eq '2') {
            # In NVARCHAR mode, only map columns that actually exist in the target
            $tgtSchemaForMapping = Get-TableSchema -FourPartName $TargetFourPartName -Connection $tgtConn
            $tgtSchemaForMapping.SchemaTable.Rows | Select-Object -ExpandProperty COLUMN_NAME
        } else {
            $schemaTable.Rows | Where-Object { $_.IS_COMPUTED -eq 0 } | Select-Object -ExpandProperty COLUMN_NAME
        }
        Write-DebugMessage "Mapping columns for bulk copy: $($columnsToMap -join ', ')"
        foreach ($col in $columnsToMap) {
            if ($DataTable.Columns.Contains($col)) {
                [void]$bulkCopy.ColumnMappings.Add($col, $col)
            }
        }

        try {
            if ($SchemaOptionToUse -eq '1' -and $useIdentityInsert) {
                $identityCmd = $tgtConn.CreateCommand()
                $identityCmd.CommandText = "SET IDENTITY_INSERT [$($tgt.Schema)].[$($tgt.Table)] ON"
                [void]$identityCmd.ExecuteNonQuery()
                Write-Host "INFO: IDENTITY_INSERT enabled for [$($tgt.Schema)].[$($tgt.Table)]." -ForegroundColor Cyan
                Write-DebugMessage "Executed: SET IDENTITY_INSERT ON"
            }
            $bulkCopy.WriteToServer($DataTable)
            Write-Host "SUCCESS: Bulk insert completed successfully." -ForegroundColor Green
        } catch {
            throw "Bulk insert failed: $_"
        } finally {
            if ($SchemaOptionToUse -eq '1' -and $useIdentityInsert) {
                $identityCmd = $tgtConn.CreateCommand()
                $identityCmd.CommandText = "SET IDENTITY_INSERT [$($tgt.Schema)].[$($tgt.Table)] OFF"
                [void]$identityCmd.ExecuteNonQuery()
                Write-Host "INFO: IDENTITY_INSERT disabled for [$($tgt.Schema)].[$($tgt.Table)]." -ForegroundColor Cyan
                Write-DebugMessage "Executed: SET IDENTITY_INSERT OFF"
            }
        }
    } catch {
        throw "Failed to process [$($tgt.Schema)].[$($tgt.Table)]: $_"
    } finally {
        $tgtConn.Close()
        $tgtConn.Dispose()
        if ($bulkCopy) {
            $bulkCopy.Close()
        }
    }
}

# Section: Main Execution
# Purpose: Orchestrates the data transfer by collecting input, fetching source data,
#          and calling BulkInsert-DataTable for schema replication and data insertion.
do {
    # If parameters are not provided, fall back to interactive prompts
    if ([string]::IsNullOrWhiteSpace($Source)) {
        Write-Host "Enter the source [Server].[Database].[Schema].[Table]: " -ForegroundColor White -NoNewline
        $Source = Read-Host
    }
    if ([string]::IsNullOrWhiteSpace($Target)) {
        Write-Host "Enter the target [Server].[Database].[Schema].[Table]: " -ForegroundColor White -NoNewline
        $Target = Read-Host
    }
    # These are not mandatory, so we only ask if they were not provided via parameters and their default is still set.
    # For SampleSize, the default is 0.
    if ($PSBoundParameters.ContainsKey('SampleSize') -eq $false) {
        Write-Host "Enter TOP N sample size (0 for full load) [Default: 0]: " -ForegroundColor White -NoNewline
        $sampleInput = Read-Host
        if ($sampleInput -match '^\d+$') {
            $SampleSize = [int]$sampleInput
        } else {
            $SampleSize = 0 # Default if input is invalid
        }
    }
    # For SchemaOption, the default is '1'.
    if ($PSBoundParameters.ContainsKey('SchemaOption') -eq $false) {
        Write-Host "Enter 1 for full schema, 2 for NVARCHAR(MAX) columns [Default: 1]: " -ForegroundColor White -NoNewline
        $schemaInput = Read-Host
        if ($schemaInput -eq '2') {
            $SchemaOption = '2'
        } else {
            $SchemaOption = '1' # Default if input is invalid or 1
        }
    }

    if ($SchemaOption -eq '2') {
        Write-Host "INFO: User chose NVARCHAR(MAX) columns for target table." -ForegroundColor Cyan
    } else {
        Write-Host "INFO: User chose full schema replication for target table." -ForegroundColor Cyan
    }

    try {
        Write-Host "INFO: Loading data from $Source..." -ForegroundColor Cyan
        $data = Get-DataTable -FourPartName $Source -Top $SampleSize

        # Execute the bulk insert operation
        BulkInsert-DataTable -SourceFourPartName $Source -TargetFourPartName $Target -DataTable $data -SchemaOptionToUse $SchemaOption

        Write-Host "SUCCESS: Done. Rows transferred: $($data.Rows.Count)" -ForegroundColor Green
    } catch {
        Write-Host "ERROR: Error occurred during data transfer: $($_.Exception.Message)" -ForegroundColor Red
    }

    Write-Host "INFO: Do you want to restart the process? Enter Y to restart, or press Enter to exit" -ForegroundColor Cyan
    $restart = Read-Host
    if ($restart -eq 'Y' -or $restart -eq 'y') {
        Write-Host "INFO: Restarting the process..." -ForegroundColor Cyan
        # Reset parameters for next loop if running interactively
        $Source = $null
        $Target = $null
        $PSBoundParameters.Remove('SampleSize') | Out-Null
        $PSBoundParameters.Remove('SchemaOption') | Out-Null
        $SampleSize = 0
        $SchemaOption = '1'
    } else {
        Write-Host "SUCCESS: Exiting script." -ForegroundColor Green
    }
} while ($restart -eq 'Y' -or $restart -eq 'y')