#V18.11
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
#   - Handles IDENTITY columns with explicit insertion (SET IDENTITY_INSERT, KeepIdentity) for full schema,
#     including seed/increment detection, data validation, and existing table checks.
#   - Supports computed columns (persisted and non-persisted) with accurate definitions for full schema,
#     or materializes their values as NVARCHAR(MAX) in simplified schema.
#   - Excludes computed columns from data insertion in full schema mode, with their values computed in the target.
#   - Replicates clustered and nonclustered rowstore indexes for full schema.
#   - Supports a wide range of SQL Server data types, with warnings for deprecated types (e.g., text).
#   - Prompts user to choose between full schema replication or creating target table with all columns
#     (including computed) as NVARCHAR(MAX), excluding IDENTITY, PKs, and indexes in simplified mode.
#   - Allows target table in NVARCHAR(MAX) mode to have a subset of source columns, mapping only
#     matching columns for data transfer.
#   - Checks if target table exists, verifies schema compatibility, checks for data, and prompts
#     user to truncate (TRUNCATE TABLE) or insert additional records if data exists.
#   - Uses SqlBulkCopy with connection string constructor for reliable data transfer.
#   - Provides user prompts for source, target, sample size, schema option, and truncate/insert,
#     with robust error handling.
#   - Supports recursive execution with user prompt to restart (Y) or exit (Enter) after each run.
#   - Streamlined output with color-coded messages: errors (red), info (cyan), warnings (yellow),
#     successes (green), covering main steps and outcomes only.
#
# Limitations:
#   - Does not support columnstore indexes, foreign keys, or full user-defined types (UDTs)/CLR types.
#   - Requires SQL Server 2017+ due to STRING_AGG usage.
#   - May fail for UDTs if not defined in the target database.
#   - IDENTITY_INSERT requires non-null IDENTITY values in source data (full schema only).
#   - Existing target tables with incompatible schemas will cause errors in full schema mode.
#   - In NVARCHAR(MAX) mode, data for source columns not present in the target table is skipped.
#   - NVARCHAR(MAX) mode may produce unreadable data for binary or spatial types.
#   - Computed columns are materialized as NVARCHAR(MAX) in simplified schema mode, losing their computed logic.

# Function: Parse-FourPartName
# Purpose: Parses a four-part table name (Server.Database.Schema.Table) into components,
#          handling bracketed or unbracketed names and validating syntax.
function Parse-FourPartName {
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$InputName
    )

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
    return @{
        Server   = $cleanParts[0]
        Database = $cleanParts[1]
        Schema   = $cleanParts[2]
        Table    = $cleanParts[3]
    }
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
    return "Server=$($parsed.Server);Database=$($parsed.Database);Integrated Security=True;Connect Timeout=15"
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
        [int]$SchemaOption
    )

    if ($SchemaOption -eq 2) {
        # NVARCHAR(MAX) mode: Check target columns are a subset of source and NVARCHAR(MAX)
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
        return @{
            IsCompatible = $true
            MissingColumns = $missingCols
        } # Schemas compatible, with possible missing columns
    } else {
        # Full schema mode: Exact match
        if ($SourceSchema.Rows.Count -ne $TargetSchema.Rows.Count) {
            return "Column count mismatch: Source ($($SourceSchema.Rows.Count)) vs Target ($($TargetSchema.Rows.Count))."
        }
        $sourceCols = $SourceSchema.Rows | Sort-Object COLUMN_NAME
        $targetCols = $TargetSchema.Rows | Sort-Object COLUMN_NAME
        for ($i = 0; $i -lt $sourceCols.Count; $i++) {
            $src = $sourceCols[$i]
            $tgt = $targetCols[$i]
            if ($src.COLUMN_NAME -ne $tgt.COLUMN_NAME) {
                return "Column name mismatch: Source [$($src.COLUMN_NAME)] vs Target [$($tgt.COLUMN_NAME)]."
            }
            if ($src.DATA_TYPE -ne $tgt.DATA_TYPE) {
                return "Data type mismatch for [$($src.COLUMN_NAME)]: Source [$($src.DATA_TYPE)] vs Target [$($tgt.DATA_TYPE)]."
            }
            if ($src.CHARACTER_MAXIMUM_LENGTH -ne $tgt.CHARACTER_MAXIMUM_LENGTH) {
                return "Length mismatch for [$($src.COLUMN_NAME)]: Source [$($src.CHARACTER_MAXIMUM_LENGTH)] vs Target [$($tgt.CHARACTER_MAXIMUM_LENGTH)]."
            }
            if ($src.NUMERIC_PRECISION -ne $tgt.NUMERIC_PRECISION) {
                return "Precision mismatch for [$($src.COLUMN_NAME)]: Source [$($src.NUMERIC_PRECISION)] vs Target [$($tgt.NUMERIC_PRECISION)]."
            }
            if ($src.NUMERIC_SCALE -ne $tgt.NUMERIC_SCALE) {
                return "Scale mismatch for [$($src.COLUMN_NAME)]: Source [$($src.NUMERIC_SCALE)] vs Target [$($tgt.NUMERIC_SCALE)]."
            }
            if ($src.DATETIME_PRECISION -ne $tgt.DATETIME_PRECISION) {
                return "Datetime precision mismatch for [$($src.COLUMN_NAME)]: Source [$($src.DATETIME_PRECISION)] vs Target [$($tgt.DATETIME_PRECISION)]."
            }
            if ($src.IS_NULLABLE -ne $tgt.IS_NULLABLE) {
                return "Nullability mismatch for [$($src.COLUMN_NAME)]: Source [$($src.IS_NULLABLE)] vs Target [$($tgt.IS_NULLABLE)]."
            }
            if ($src.IS_PRIMARY_KEY -ne $tgt.IS_PRIMARY_KEY) {
                return "Primary key status mismatch for [$($src.COLUMN_NAME)]: Source [$($src.IS_PRIMARY_KEY)] vs Target [$($tgt.IS_PRIMARY_KEY)]."
            }
            if ($src.IS_COMPUTED -ne $tgt.IS_COMPUTED) {
                return "Computed column status mismatch for [$($src.COLUMN_NAME)]: Source [$($src.IS_COMPUTED)] vs Target [$($tgt.IS_COMPUTED)]."
            }
            if ($src.IS_COMPUTED -eq 1) {
                if ($src.COMPUTED_DEFINITION -ne $tgt.COMPUTED_DEFINITION) {
                    return "Computed column definition mismatch for [$($src.COLUMN_NAME)]."
                }
                if ($src.IS_PERSISTED -ne $tgt.IS_PERSISTED) {
                    return "Computed column persisted status mismatch for [$($src.COLUMN_NAME)]: Source [$($src.IS_PERSISTED)] vs Target [$($src.IS_PERSISTED)]."
                }
            }
            if ($src.IS_IDENTITY -ne $tgt.IS_IDENTITY) {
                return "IDENTITY status mismatch for [$($src.COLUMN_NAME)]: Source [$($src.IS_IDENTITY)] vs Target [$($tgt.IS_IDENTITY)]."
            }
            if ($src.IS_IDENTITY -eq 1) {
                if ($src.IDENTITY_SEED -ne $tgt.IDENTITY_SEED -or $src.IDENTITY_INCREMENT -ne $tgt.IDENTITY_INCREMENT) {
                    return "IDENTITY properties mismatch for [$($src.COLUMN_NAME)]: Source [IDENTITY($($src.IDENTITY_SEED),$($src.IDENTITY_INCREMENT))] vs Target [IDENTITY($($tgt.IDENTITY_SEED),$($tgt.IDENTITY_INCREMENT))]."
                }
            }
        }
        if ($SourcePkIsClustered -ne $TargetPkIsClustered) {
            return "Primary key clustering mismatch: Source [IsClustered=$SourcePkIsClustered] vs Target [IsClustered=$TargetPkIsClustered]."
        }
        if ($SourceIndexes.Rows.Count -ne $TargetIndexes.Rows.Count) {
            return "Index count mismatch: Source ($($SourceIndexes.Rows.Count)) vs Target ($($TargetIndexes.Rows.Count))."
        }
        $sourceIdx = $SourceIndexes.Rows | Sort-Object KeyColumns, IncludedColumns
        $targetIdx = $TargetIndexes.Rows | Sort-Object KeyColumns, IncludedColumns
        for ($i = 0; $i -lt $sourceIdx.Count; $i++) {
            $src = $sourceIdx[$i]
            $tgt = $targetIdx[$i]
            if ($src.IndexType -ne $tgt.IndexType) {
                return "Index type mismatch for index: Source [Type=$($src.IndexType)] vs Target [Type=$($tgt.IndexType)]."
            }
            if ($src.IsUnique -ne $tgt.IsUnique) {
                return "Index uniqueness mismatch for index: Source [IsUnique=$($src.IsUnique)] vs Target [IsUnique=$($tgt.IsUnique)]."
            }
            if ($src.KeyColumns -ne $tgt.KeyColumns) {
                return "Index key columns mismatch: Source [$($src.KeyColumns)] vs Target [$($tgt.KeyColumns)]."
            }
            if ($src.IncludedColumns -ne $tgt.IncludedColumns) {
                return "Index included columns mismatch: Source [$($src.IncludedColumns)] vs Target [$($tgt.IncludedColumns)]."
            }
        }
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

    # Execute the query and fill a DataTable
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $query
    $cmd.CommandTimeout = 30
    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
    $table = New-Object System.Data.DataTable
    try {
        $conn.Open()
        $adapter.Fill($table) | Out-Null
    } catch {
        throw "Failed to fetch data from [$($parsed.Schema)].[$($parsed.Table)]: $_"
    } finally {
        $conn.Close()
        $conn.Dispose()
    }
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
        [int]$SchemaOption
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
    foreach ($col in $colsToRemove) {
        $DataTable.Columns.Remove($col)
    }

    # Check for IDENTITY and computed columns
    $identityCols = $schemaTable.Rows | Where-Object { $_.IS_IDENTITY -eq 1 } | Select-Object -ExpandProperty COLUMN_NAME
    $computedCols = $schemaTable.Rows | Where-Object { $_.IS_COMPUTED -eq 1 } | Select-Object -ExpandProperty COLUMN_NAME
    $useIdentityInsert = $false
    if ($SchemaOption -eq 1 -and $identityCols.Count -gt 0) {
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
                }
            }
        }
        if ($computedCols.Count -gt 0) {
            Write-Host "WARNING: Computed columns [$($computedCols -join ', ')] in source table [$($src.Schema)].[$($src.Table)] will have their values computed in the target and are excluded from data insertion." -ForegroundColor Yellow
        }
    } elseif ($SchemaOption -eq 2) {
        if ($identityCols.Count -gt 0) {
            Write-Host "WARNING: IDENTITY columns [$($identityCols -join ', ')] in source table [$($src.Schema)].[$($src.Table)] will be NVARCHAR(MAX) in target." -ForegroundColor Yellow
        }
        if ($computedCols.Count -gt 0) {
            Write-Host "WARNING: Computed columns [$($computedCols -join ', ')] in source table [$($src.Schema)].[$($src.Table)] will be materialized as NVARCHAR(MAX) in target." -ForegroundColor Yellow
        }
    }

    # Connect to target database
    $tgtConn = New-Object System.Data.SqlClient.SqlConnection $tgtConnStr
    try {
        $tgtConn.Open()

        # Check if target table exists
        $checkCmd = $tgtConn.CreateCommand()
        $checkCmd.CommandText = @"
SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES
WHERE TABLE_SCHEMA = @Schema AND TABLE_NAME = @Table
"@
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
                -SchemaOption $SchemaOption
            if ($SchemaOption -eq 2) {
                if (-not $schemaCheck.IsCompatible) {
                    Write-Host "WARNING: Schema mismatch between source and target: $($schemaCheck)" -ForegroundColor Yellow
                    throw "Cannot proceed due to incompatible target table schema."
                }
                if ($schemaCheck.MissingColumns) {
                    Write-Host "WARNING: Target table [$($tgt.Schema)].[$($tgt.Table)] is missing columns [$($schemaCheck.MissingColumns -join ', ')]. Data for these columns will be skipped." -ForegroundColor Yellow
                }
                Write-Host "INFO: Target table schema is compatible for option $SchemaOption (NVARCHAR(MAX) mode)." -ForegroundColor Cyan
            } else {
                if ($schemaCheck) {
                    Write-Host "WARNING: Schema mismatch between source and target: $schemaCheck" -ForegroundColor Yellow
                    throw "Cannot proceed due to incompatible target table schema."
                }
                Write-Host "INFO: Target table schema matches expected schema for option $SchemaOption." -ForegroundColor Cyan
            }

            # Check for IDENTITY columns in target (full schema only)
            if ($SchemaOption -eq 1) {
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
                $choice = Read-Host "Target table contains data. Enter 1 to truncate, 2 to insert additional records"
                if ($choice -eq '1') {
                    Write-Host "INFO: User chose to truncate table [$($tgt.Schema)].[$($tgt.Table)]." -ForegroundColor Cyan
                    $truncateCmd = $tgtConn.CreateCommand()
                    $truncateCmd.CommandText = "TRUNCATE TABLE [$($tgt.Schema)].[$($tgt.Table)]"
                    [void]$truncateCmd.ExecuteNonQuery()
                    Write-Host "SUCCESS: Table [$($tgt.Schema)].[$($tgt.Table)] truncated." -ForegroundColor Green
                } else {
                    Write-Host "INFO: User chose to insert additional records into [$($tgt.Schema)].[$($tgt.Table)]." -ForegroundColor Cyan
                }
            } else {
                Write-Host "INFO: Target table is empty. Proceeding with data insertion." -ForegroundColor Cyan
            }
        } else {
            Write-Host "INFO: Target table [$($tgt.Schema)].[$($tgt.Table)] does not exist. Creating new table with schema option $SchemaOption..." -ForegroundColor Cyan

            if ($SchemaOption -eq 2) {
                # NVARCHAR(MAX) mode: Create table with all columns as NVARCHAR(MAX)
                $cols = @()
                foreach ($row in $schemaTable.Rows) {
                    $name = $row.COLUMN_NAME
                    $cols += "[$name] NVARCHAR(MAX) NULL"
                }
                if ($cols.Count -eq 0) {
                    throw "No valid column definitions found for [$($tgt.Schema)].[$($tgt.Table)]."
                }
                $createSql = "CREATE TABLE [$($tgt.Schema)].[$($tgt.Table)] (`n" + ($cols -join ",`n") + "`n)"
                $createCmd = $tgtConn.CreateCommand()
                $createCmd.CommandText = $createSql
                [void]$createCmd.ExecuteNonQuery()
                Write-Host "SUCCESS: Table [$($tgt.Schema)].[$($tgt.Table)] created with all columns (including computed) as NVARCHAR(MAX)." -ForegroundColor Green
            } else {
                # Full schema mode
                $hasClusteredIndex = ($indexTable.Rows | Where-Object { $_.IndexType -eq 1 } | Select-Object -First 1) -or ($pkIsClustered -eq 1)
                $primaryKeyType = if ($pkIsClustered -eq 1) { 'CLUSTERED' } elseif ($hasClusteredIndex) { 'NONCLUSTERED' } else { 'CLUSTERED' }

                $cols = @()
                $pkCols = @()
                foreach ($row in $schemaTable.Rows) {
                    $name = $row.COLUMN_NAME
                    if ($row.IS_COMPUTED -eq 1) {
                        if ([string]::IsNullOrWhiteSpace($row.COMPUTED_DEFINITION)) {
                            throw "Invalid computed column definition for [$name] in [$($src.Schema)].[$($src.Table)]."
                        }
                        $persisted = if ($row.IS_PERSISTED -eq $true) { ' PERSISTED' } else { '' }
                        $cols += "[$name] AS $($row.COMPUTED_DEFINITION)$persisted"
                    } else {
                        $type = switch ($row.DATA_TYPE) {
                            'nvarchar' { 
                                if ($row.CHARACTER_MAXIMUM_LENGTH -gt 0) { "NVARCHAR($($row.CHARACTER_MAXIMUM_LENGTH))" } else { "NVARCHAR(MAX)" }
                            }
                            'varchar' { 
                                if ($row.CHARACTER_MAXIMUM_LENGTH -gt 0) { "VARCHAR($($row.CHARACTER_MAXIMUM_LENGTH))" } else { "VARCHAR(MAX)" }
                            }
                            'char' { "CHAR($($row.CHARACTER_MAXIMUM_LENGTH))" }
                            'nchar' { "NCHAR($($row.CHARACTER_MAXIMUM_LENGTH))" }
                            'decimal' { "DECIMAL($($row.NUMERIC_PRECISION),$($row.NUMERIC_SCALE))" }
                            'numeric' { "NUMERIC($($row.NUMERIC_PRECISION),$($row.NUMERIC_SCALE))" }
                            'varbinary' { 
                                if ($row.CHARACTER_MAXIMUM_LENGTH -gt 0) { "VARBINARY($($row.CHARACTER_MAXIMUM_LENGTH))" } else { "VARBINARY(MAX)" }
                            }
                            'binary' { 
                                if ($row.CHARACTER_MAXIMUM_LENGTH -gt 0) { "BINARY($($row.CHARACTER_MAXIMUM_LENGTH))" } else { "BINARY(MAX)" }
                            }
                            'datetime2' { 
                                if ($row.DATETIME_PRECISION -gt 0) { "DATETIME2($($row.DATETIME_PRECISION))" } else { "DATETIME2" }
                            }
                            'time' { 
                                if ($row.DATETIME_PRECISION -gt 0) { "TIME($($row.DATETIME_PRECISION))" } else { "TIME" }
                            }
                            'int' { 'INT' }
                            'bigint' { 'BIGINT' }
                            'smallint' { 'SMALLINT' }
                            'tinyint' { 'TINYINT' }
                            'float' { 
                                if ($row.NUMERIC_PRECISION -gt 0) { "FLOAT($($row.NUMERIC_PRECISION))" } else { "FLOAT" }
                            }
                            'real' { 'REAL' }
                            'money' { 'MONEY' }
                            'smallmoney' { 'SMALLMONEY' }
                            'date' { 'DATE' }
                            'datetime' { 'DATETIME' }
                            'smalldatetime' { 'SMALLDATETIME' }
                            'datetimeoffset' { 
                                if ($row.DATETIME_PRECISION -gt 0) { "DATETIMEOFFSET($($row.DATETIME_PRECISION))" } else { "DATETIMEOFFSET" }
                            }
                            'uniqueidentifier' { 'UNIQUEIDENTIFIER' }
                            'xml' { 'XML' }
                            'geometry' { 'GEOMETRY' }
                            'geography' { 'GEOGRAPHY' }
                            'hierarchyid' { 'HIERARCHYID' }
                            'timestamp' { 'TIMESTAMP' }
                            'rowversion' { 'ROWVERSION' }
                            'sql_variant' { 'SQL_VARIANT' }
                            'text' { 
                                Write-Host "WARNING: Deprecated data type 'text' used in [$($src.Schema)].[$($src.Table)]. Consider upgrading." -ForegroundColor Yellow
                                'TEXT'
                            }
                            'ntext' { 
                                Write-Host "WARNING: Deprecated data type 'ntext' used in [$($src.Schema)].[$($src.Table)]. Consider upgrading." -ForegroundColor Yellow
                                'NTEXT'
                            }
                            'image' { 
                                Write-Host "WARNING: Deprecated data type 'image' used in [$($src.Schema)].[$($src.Table)]. Consider upgrading." -ForegroundColor Yellow
                                'IMAGE'
                            }
                            default { 
                                Write-Host "WARNING: Unsupported or user-defined data type '$($row.DATA_TYPE)' in [$($src.Schema)].[$($src.Table)]. Using as-is, which may cause errors." -ForegroundColor Yellow
                                $row.DATA_TYPE.ToUpper()
                            }
                        }
                        $identity = ''
                        if ($row.IS_IDENTITY -eq 1) {
                            if ($null -ne $row.IDENTITY_SEED -and $null -ne $row.IDENTITY_INCREMENT) {
                                $identity = " IDENTITY($($row.IDENTITY_SEED),$($row.IDENTITY_INCREMENT))"
                            } else {
                                Write-Host "WARNING: IDENTITY column [$name] in [$($src.Schema)].[$($src.Table)] has missing seed/increment values. Defaulting to IDENTITY(1,1)." -ForegroundColor Yellow
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
                if ($cols.Count -eq 0) {
                    throw "No valid column definitions found for [$($tgt.Schema)].[$($tgt.Table)]."
                }
                $createSql = "CREATE TABLE [$($tgt.Schema)].[$($tgt.Table)] (`n" + ($cols -join ",`n")
                if ($pkCols.Count -gt 0) {
                    $createSql += ",`nCONSTRAINT [PK_$($tgt.Table)] PRIMARY KEY $primaryKeyType (" + ($pkCols -join ', ') + ")"
                }
                $createSql += "`n)"
                $createCmd = $tgtConn.CreateCommand()
                $createCmd.CommandText = $createSql
                [void]$createCmd.ExecuteNonQuery()
                Write-Host "SUCCESS: Table [$($tgt.Schema)].[$($tgt.Table)] created with full schema." -ForegroundColor Green

                # Create indexes
                foreach ($idx in $indexTable.Rows) {
                    $idxName = $idx.IndexName
                    $isUnique = $idx.IsUnique
                    $indexKind = $idx.IndexKind
                    $keyColumns = $idx.KeyColumns
                    $includedColumns = $idx.IncludedColumns
                    if ($indexKind -eq 'CLUSTERED' -and $hasClusteredIndex) {
                        Write-Host "INFO: Skipping clustered index [$idxName] on [$($tgt.Schema)].[$($tgt.Table)] as a clustered index already exists." -ForegroundColor Cyan
                        continue
                    }
                    $idxSql = "CREATE "
                    if ($isUnique -eq $true) { $idxSql += "UNIQUE " }
                    $idxSql += "$indexKind INDEX [$idxName] ON [$($tgt.Schema)].[$($tgt.Table)] ($keyColumns)"
                    if (-not [string]::IsNullOrWhiteSpace($includedColumns)) {
                        $idxSql += " INCLUDE ($includedColumns)"
                    }
                    $idxCmd = $tgtConn.CreateCommand()
                    $idxCmd.CommandText = $idxSql
                    [void]$idxCmd.ExecuteNonQuery()
                    Write-Host "SUCCESS: Index [$idxName] ($indexKind) created on [$($tgt.Schema)].[$($tgt.Table)]." -ForegroundColor Green
                }
            }
        }

        # Perform bulk data insertion
        $bulkCopyOptions = if ($SchemaOption -eq 1 -and $useIdentityInsert) { [System.Data.SqlClient.SqlBulkCopyOptions]::KeepIdentity } else { [System.Data.SqlClient.SqlBulkCopyOptions]::Default }
        $bulkCopy = New-Object System.Data.SqlClient.SqlBulkCopy($tgtConnStr, $bulkCopyOptions)
        $bulkCopy.DestinationTableName = "[$($tgt.Schema)].[$($tgt.Table)]"
        $bulkCopy.BatchSize = 10000
        $bulkCopy.BulkCopyTimeout = 600

        # Map columns: Use non-computed columns for Option 1, target table columns for Option 2
        $columnsToMap = if ($SchemaOption -eq 2) {
            $tgtSchemaTable.Rows | Select-Object -ExpandProperty COLUMN_NAME
        } else {
            $schemaTable.Rows | Where-Object { $_.IS_COMPUTED -eq 0 } | Select-Object -ExpandProperty COLUMN_NAME
        }
        foreach ($col in $columnsToMap) {
            if ($DataTable.Columns.Contains($col)) {
                $bulkCopy.ColumnMappings.Add($col, $col) | Out-Null
            }
        }

        try {
            if ($SchemaOption -eq 1 -and $useIdentityInsert) {
                $identityCmd = $tgtConn.CreateCommand()
                $identityCmd.CommandText = "SET IDENTITY_INSERT [$($tgt.Schema)].[$($tgt.Table)] ON"
                [void]$identityCmd.ExecuteNonQuery()
                Write-Host "INFO: IDENTITY_INSERT enabled for [$($tgt.Schema)].[$($tgt.Table)]." -ForegroundColor Cyan
            }
            $bulkCopy.WriteToServer($DataTable)
            Write-Host "SUCCESS: Bulk insert completed successfully." -ForegroundColor Green
        } catch {
            throw "Bulk insert failed: $_"
        } finally {
            if ($SchemaOption -eq 1 -and $useIdentityInsert) {
                $identityCmd = $tgtConn.CreateCommand()
                $identityCmd.CommandText = "SET IDENTITY_INSERT [$($tgt.Schema)].[$($tgt.Table)] OFF"
                [void]$identityCmd.ExecuteNonQuery()
                Write-Host "INFO: IDENTITY_INSERT disabled for [$($tgt.Schema)].[$($tgt.Table)]." -ForegroundColor Cyan
            }
        }
    } catch {
        throw "Failed to process [$($tgt.Schema)].[$($tgt.Table)]: $_"
    } finally {
        $tgtConn.Close()
        $tgtConn.Dispose()
        if ($bulkCopy) {
            $bulkCopy.Close()
            $bulkCopy.Dispose()
        }
    }
}

# Section: Main Execution
# Purpose: Orchestrates the data transfer by fetching source data into a DataTable
#          and calling BulkInsert-DataTable for schema replication and data insertion,
#          with recursive option to restart or exit.
do {
    # Section: Runtime Prompts
    # Purpose: Collects user input for source table, target table, sample size, and schema option,
    #          validates the input, and initiates the data transfer process.
    $sourceInput = Read-Host "Enter the source [Server].[Database].[Schema].[Table]"
    $targetInput = Read-Host "Enter the target [Server].[Database].[Schema].[Table]"
    $sampleInput = Read-Host "Enter TOP N sample size (0 for full load)"
    $schemaInput = Read-Host "Enter 1 for full schema (data types, PK, indexes, IDENTITY, computed columns), 2 for NVARCHAR(MAX) columns only"
    [int]$sampleSize = 0
    if ($sampleInput -match '^\d+$') {
        $sampleSize = [int]$sampleInput
        if ($sampleSize -lt 0) {
            throw "Sample size cannot be negative: $sampleInput"
        }
    }
    [int]$schemaOption = 1 # Default to full schema
    if ($schemaInput -eq '2') {
        $schemaOption = 2
        Write-Host "INFO: User chose NVARCHAR(MAX) columns for target table." -ForegroundColor Cyan
    } else {
        Write-Host "INFO: User chose full schema replication for target table." -ForegroundColor Cyan
    }

    try {
        Write-Host "INFO: Loading data from $sourceInput..." -ForegroundColor Cyan
        $data = Get-DataTable -FourPartName $sourceInput -Top $sampleSize

        # Convert array of objects to DataTable if necessary
        if ($data -is [System.Data.DataTable]) {
            # Already a DataTable, no action needed
        } elseif ($data -is [System.Object[]]) {
            $dt = New-Object System.Data.DataTable
            if ($data.Length -gt 0) {
                foreach ($prop in $data[0].PSObject.Properties.Name) {
                    $dt.Columns.Add($prop) | Out-Null
                }
                foreach ($row in $data) {
                    $newRow = $dt.NewRow()
                    foreach ($col in $dt.Columns) {
                        $newRow[$col.ColumnName] = $row.$($col.ColumnName)
                    }
                    $dt.Rows.Add($newRow)
                }
            }
            $data = $dt
        } else {
            throw "Unexpected type: $($data.GetType().FullName)"
        }

        # Execute the bulk insert operation
        BulkInsert-DataTable -SourceFourPartName $sourceInput -TargetFourPartName $targetInput -DataTable $data -SchemaOption $schemaOption

        Write-Host "SUCCESS: Done. Rows transferred: $($data.Rows.Count)" -ForegroundColor Green
    } catch {
        Write-Host "ERROR: Error occurred during data transfer: $_" -ForegroundColor Red
    }

    Write-Host "INFO: Do you want to restart the process? Enter Y to restart, or press Enter to exit" -ForegroundColor Cyan
    $restart = Read-Host
    if ($restart -eq 'Y' -or $restart -eq 'y') {
        Write-Host "INFO: Restarting the process..." -ForegroundColor Cyan
    } else {
        Write-Host "SUCCESS: Exiting script." -ForegroundColor Green
    }
} while ($restart -eq 'Y' -or $restart -eq 'y')