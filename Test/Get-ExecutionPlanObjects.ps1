<#
.SYNOPSIS
    Intelligently retrieves execution plans for SQL queries, parsing all referenced database objects into a destination table and logging any errors.

.DESCRIPTION
    This script implements a hybrid approach to retrieve execution plans. For each query, it first attempts to get the fast "Estimated Plan". If that fails (e.g., due to temporary tables), it automatically falls back to getting the "Actual Plan" by fully executing the query. If both attempts fail, the SQL Server error message is logged to the destination table.
    It reads the source queries provided by a user-defined query, processes them in batches, parses all unique, non-temporary objects from the plans, and bulk inserts the results.

    CRITICAL NOTE: For queries that require the "Actual Plan" fallback, this script will perform a full execution. This may still result in significant server load for complex queries.

.PARAMETER SourceServer
    The FQDN or instance name of the SQL Server where the source query will be executed.

.PARAMETER SourceQuery
    A full T-SQL query that returns columns named [correlation_id], [database_name], and [sql_text].

.PARAMETER DestinationServer
    The FQDN or instance name of the SQL Server where the destination table is located.

.PARAMETER DestinationDatabase
    The name of the database containing the destination table.

.PARAMETER DestinationTable
    The name of the table where the results will be stored. This table must have an 'ErrorMsg' column to support error logging.

.PARAMETER BatchSize
    The number of records to process from the source query's results in each iteration.

.PARAMETER ShowProgress
    If specified, the script will first run a COUNT(*) on the source query to get a total and display a progress bar. This may add to the script's total runtime.

.PARAMETER DebugMode
    A switch that enables verbose logging for troubleshooting.

.EXAMPLE
    $MyQuery = "SELECT a.correlation_id, a.database_name, a.sql_text FROM ..."
    .\Get-ExecutionPlanObjects.ps1 -SourceServer "warehouse.selectquote.com" -SourceQuery $MyQuery -DestinationServer "warehouse.selectquote.com" -DestinationDatabase "Bi_monitoring" -DestinationTable "AuditSelectsForQuickSightReporting_ObjectList" -ShowProgress

.NOTES
    Version: 5.4
    Author: Gemini, Powershell Developer
    Compatibility: PowerShell 5.1 or higher.
    
    V5.2: Added error logging to the destination table for failed queries.
    V5.3: Re-introduced the optional -ShowProgress switch.
    V5.4: Fixed a bug in the Log-ProcessingError function where SQL parameters were not being passed correctly. The function now uses a robust .NET SqlClient method.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceServer,

    [Parameter(Mandatory = $true)]
    [string]$SourceQuery,

    [Parameter(Mandatory = $true)]
    [string]$DestinationServer,

    [Parameter(Mandatory = $true)]
    [string]$DestinationDatabase,

    [Parameter(Mandatory = $true)]
    [string]$DestinationTable,

    [Parameter(Mandatory = $false)]
    [int]$BatchSize = 1000,

    [Parameter(Mandatory = $false)]
    [switch]$ShowProgress,

    [Parameter(Mandatory = $false)]
    [switch]$DebugMode
)

#region Functions
function Log-ProcessingError {
    param(
        [string]$ServerInstance,
        [string]$Database,
        [string]$Table,
        [string]$CorrelationId,
        [string]$ErrorMessage
    )

    if ($DebugMode) { Write-Host -Object "DEBUG ($CorrelationId): Logging failure to destination table..." -ForegroundColor Yellow }
    
    $maxErrorLength = 2000
    if ($ErrorMessage.Length -gt $maxErrorLength) {
        $truncatedErrorMessage = $ErrorMessage.Substring(0, $maxErrorLength)
    } else {
        $truncatedErrorMessage = $ErrorMessage
    }

    $connection = $null
    $command = $null
    try {
        $connectionString = "Server=$($ServerInstance);Database=$($Database);Integrated Security=True;TrustServerCertificate=True;"
        $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
        $connection.Open()

        $insertQuery = "INSERT INTO dbo.$($Table) (correlation_id, ErrorMsg) VALUES (@correlationId, @errorMsg);"
        $command = New-Object System.Data.SqlClient.SqlCommand($insertQuery, $connection)
        
        # Add parameters safely to prevent SQL injection
        $command.Parameters.AddWithValue("@correlationId", $CorrelationId) | Out-Null
        $command.Parameters.AddWithValue("@errorMsg", $truncatedErrorMessage) | Out-Null
        
        $command.ExecuteNonQuery() | Out-Null
        
        if ($DebugMode) { Write-Host -Object "DEBUG ($CorrelationId): Successfully logged error." -ForegroundColor Green }
    }
    catch {
        Write-Host -Object "FATAL: Could not log error for correlation_id '$($CorrelationId)' to the database. Error: $($_.Exception.Message)" -ForegroundColor Red
    }
    finally {
        if ($null -ne $command) { $command.Dispose() }
        if ($null -ne $connection) { $connection.Close() }
    }
}

function Get-ExecutionPlan {
    param(
        [string]$ServerInstance,
        [string]$Database,
        [string]$SqlText,
        [string]$CorrelationId,
        [string]$DestinationServer,
        [string]$DestinationDatabase,
        [string]$DestinationTable
    )
    
    # --- Attempt 1: Fast Path (Estimated Plan) ---
    if ($DebugMode) { Write-Host -Object "DEBUG ($CorrelationId): Attempting FAST path (Estimated Plan)..." -ForegroundColor Cyan }
    $estimatedPlanCommand = "SET SHOWPLAN_XML ON;"
    $estimatedPlanQuery = "$($estimatedPlanCommand)`nGO`n$($SqlText)"

    try {
        $planResult = Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -TrustServerCertificate -Query $estimatedPlanQuery -MaxCharLength ([int]::MaxValue) -ErrorAction Stop
        
        $planFragments = @()
        foreach ($resultSet in $planResult) {
            if ($resultSet) {
                $potentialPlan = $resultSet.Item(0)
                if ($potentialPlan -is [string] -and $potentialPlan -like '<*showplan*>') { $planFragments += $potentialPlan }
            }
        }
        
        if ($planFragments.Count -eq 0) {
            if ($DebugMode) { Write-Host -Object "DEBUG ($CorrelationId): FAST path executed but returned no plan. No fallback." -ForegroundColor Yellow }
            return $null
        }

        $masterPlanXmlString = "<MasterShowPlan>" + ($planFragments -join '') + "</MasterShowPlan>"
        [xml]$masterPlanXml = $masterPlanXmlString
        if ($DebugMode) { Write-Host -Object "DEBUG ($CorrelationId): FAST path successful." -ForegroundColor Green }
        return $masterPlanXml
    }
    catch {
        # --- Attempt 2: Slow Path Fallback (Actual Plan) ---
        if ($DebugMode) {
            $errorMessage = $_.Exception.Message.Split([System.Environment]::NewLine)[0]
            Write-Host -Object "DEBUG ($CorrelationId): FAST path failed ('$errorMessage'). Falling back to SLOW path (Actual Plan)..." -ForegroundColor Yellow
        }

        try {
            $actualPlanCommand = "SET STATISTICS XML ON;"
            $actualPlanQuery = "$($actualPlanCommand)`nGO`n$($SqlText)"
            $planResult = Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $Database -TrustServerCertificate -Query $actualPlanQuery -MaxCharLength ([int]::MaxValue) -ErrorAction Stop
        
            $planFragments = @()
            foreach ($resultSet in $planResult) {
                if ($resultSet) {
                    $potentialPlan = $resultSet.Item(0)
                    if ($potentialPlan -is [string] -and $potentialPlan -like '<*showplan*>') { $planFragments += $potentialPlan }
                }
            }
        
            if ($planFragments.Count -eq 0) {
                if ($DebugMode) { Write-Host -Object "DEBUG ($CorrelationId): SLOW path executed but returned no plan." -ForegroundColor Yellow }
                return $null
            }

            $masterPlanXmlString = "<MasterShowPlan>" + ($planFragments -join '') + "</MasterShowPlan>"
            [xml]$masterPlanXml = $masterPlanXmlString
            if ($DebugMode) { Write-Host -Object "DEBUG ($CorrelationId): SLOW path successful." -ForegroundColor Green }
            return $masterPlanXml
        }
        catch {
             $finalErrorMessage = $_.Exception.Message
             Write-Host -Object "Warning: ($CorrelationId) script failed on both Estimated and Actual plan attempts. Logging error..." -ForegroundColor Red
             Log-ProcessingError -ServerInstance $DestinationServer -Database $DestinationDatabase -Table $DestinationTable -CorrelationId $CorrelationId -ErrorMessage $finalErrorMessage
             return $null
        }
    }
}

function Parse-PlanObjects {
    param(
        [xml]$ExecutionPlanXml,
        [string]$CorrelationId
    )
    if ($DebugMode) { Write-Host -Object "DEBUG ($CorrelationId): Parsing execution plan to find objects..." -ForegroundColor Cyan }

    try {
        $ns = New-Object System.Xml.XmlNamespaceManager($ExecutionPlanXml.NameTable)
        $ns.AddNamespace("shp", "http://schemas.microsoft.com/sqlserver/2004/07/showplan")
        
        $objectNodes = $ExecutionPlanXml.SelectNodes("//shp:Object", $ns)
        $uniqueObjectNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        foreach($node in $objectNodes) {
            $db = $node.GetAttribute("Database")
            $schema = $node.GetAttribute("Schema")
            $table = $node.GetAttribute("Table")
            
            if ($table -notlike "#*" -and -not ([string]::IsNullOrWhiteSpace($db)) -and -not ([string]::IsNullOrWhiteSpace($schema))) {
                $fullName = "$($db).$($schema).$($table)".Replace('[','').Replace(']','')
                $uniqueObjectNames.Add($fullName) | Out-Null
            }
        }

        if ($DebugMode -and $uniqueObjectNames.Count -gt 0) { Write-Host -Object "DEBUG ($CorrelationId): Found $($uniqueObjectNames.Count) unique objects in the plan." -ForegroundColor Green }
        
        return $uniqueObjectNames
    }
    catch {
        Write-Host -Object "Warning ($CorrelationId): Failed to parse objects from the provided execution plan. Error: $($_.Exception.Message)" -ForegroundColor Yellow
        return @()
    }
}
#endregion Functions

#region Main Script Body
function Main {
    # --- Version Header ---
    Write-Host -Object "Running Get-ExecutionPlanObjects.ps1 - Version 5.4" -ForegroundColor White
    
    try {
        # --- Pre-flight Checks ---
        Write-Host -Object "Starting pre-flight checks..." -ForegroundColor Cyan
        $destTableExists = Invoke-Sqlcmd -Query "SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = '$($DestinationTable)'" -ServerInstance $DestinationServer -Database $DestinationDatabase -TrustServerCertificate
        if(-not $destTableExists) {
             Write-Host -Object "ERROR: Destination table '[$($DestinationDatabase)].[dbo].[$($DestinationTable)]' not found." -ForegroundColor Red
             return
        }
        Write-Host -Object "Destination table check passed." -ForegroundColor Green
        
        # --- Get Total Record Count for Progress Bar ---
        $totalRecordCount = 0
        if ($ShowProgress) {
            Write-Host -Object "ShowProgress enabled. Performing initial count of source records..." -ForegroundColor Cyan
            try {
                $countQuery = ";WITH SourceData AS ($($SourceQuery)) SELECT COUNT_BIG(*) FROM SourceData;"
                $countResult = Invoke-Sqlcmd -ServerInstance $SourceServer -Database "master" -Query $countQuery -MaxCharLength ([int]::MaxValue) -TrustServerCertificate
                $totalRecordCount = $countResult.Item(0)
                Write-Host -Object "Found a total of $totalRecordCount records to process." -ForegroundColor Green
            } catch {
                Write-Host -Object "Warning: Could not perform initial count. Progress bar will be disabled. Error: $($_.Exception.Message)" -ForegroundColor Yellow
                $ShowProgress = $false
            }
        }

        # --- Initialization ---
        $offset = 0
        $totalRowsProcessed = 0
        $totalObjectsFound = 0
        $batchNum = 1
        $recordsProcessedCounter = 0

        # --- Main Processing Loop ---
        while ($true) {
            Write-Host -Object "Processing batch #$($batchNum) (Rows $($offset) to $($offset + $BatchSize -1))..." -ForegroundColor Cyan
            
            $batchQuery = ";WITH SourceData AS ($($SourceQuery)) SELECT correlation_id, database_name, sql_text FROM SourceData ORDER BY correlation_id OFFSET $offset ROWS FETCH NEXT $BatchSize ROWS ONLY;"
            $sourceDataBatch = Invoke-Sqlcmd -ServerInstance $SourceServer -Database "master" -Query $batchQuery -MaxCharLength ([int]::MaxValue) -TrustServerCertificate -ErrorAction SilentlyContinue

            if ($null -eq $sourceDataBatch) {
                if ($totalRowsProcessed -eq 0) { Write-Host -Object "Warning: The source query returned an error or no data. Please check your -SourceQuery for errors." -ForegroundColor Red }
                break 
            }

            if ($sourceDataBatch.Count -eq 0) {
                Write-Host -Object "No more rows to process. Exiting loop." -ForegroundColor Green
                break
            }

            $bulkData = New-Object System.Data.DataTable
            $bulkData.Columns.Add("correlation_id", [string]) | Out-Null
            $bulkData.Columns.Add("object_name", [string]) | Out-Null
            $bulkData.Columns.Add("ErrorMsg", [string]) | Out-Null

            foreach ($row in $sourceDataBatch) {
                $recordsProcessedCounter++
                if ($ShowProgress -and $totalRecordCount -gt 0) {
                    $percentComplete = ($recordsProcessedCounter / $totalRecordCount) * 100
                    $status = "Processing record $recordsProcessedCounter of $totalRecordCount"
                    Write-Progress -Activity "Parsing SQL Execution Plans" -Status $status -PercentComplete $percentComplete
                }

                $correlationId = $row.correlation_id
                $dbName = $row.database_name
                $sqlText = $row.sql_text
                
                if ([string]::IsNullOrWhiteSpace($sqlText)) {
                    Write-Host -Object "Warning: SQL text is empty for correlation_id '$($correlationId)'. Skipping." -ForegroundColor Yellow
                    continue
                }
                
                [xml]$plan = Get-ExecutionPlan -ServerInstance $SourceServer -Database $dbName -SqlText $sqlText -CorrelationId $correlationId -DestinationServer $DestinationServer -DestinationDatabase $DestinationDatabase -DestinationTable $DestinationTable
                
                if ($plan) {
                    $foundObjects = Parse-PlanObjects -ExecutionPlanXml $plan -CorrelationId $correlationId
                    
                    foreach ($objectName in $foundObjects) {
                        $newRow = $bulkData.NewRow()
                        $newRow.correlation_id = $correlationId
                        $newRow.object_name = $objectName
                        $bulkData.Rows.Add($newRow)
                    }
                    $totalObjectsFound += $foundObjects.Count
                }
            } 

            if ($bulkData.Rows.Count -gt 0) {
                Write-Host -Object "Bulk inserting $($bulkData.Rows.Count) successful objects found in this batch..." -ForegroundColor Cyan
                try {
                    $bulkCopyConnectionString = "Server=$($DestinationServer);Database=$($DestinationDatabase);Integrated Security=True;TrustServerCertificate=True;"
                    $bulkCopy = New-Object System.Data.SqlClient.SqlBulkCopy($bulkCopyConnectionString)
                    $bulkCopy.DestinationTableName = $DestinationTable
                    $bulkCopy.ColumnMappings.Add("correlation_id", "correlation_id") | Out-Null
                    $bulkCopy.ColumnMappings.Add("object_name", "object_name") | Out-Null
                    $bulkCopy.ColumnMappings.Add("ErrorMsg", "ErrorMsg") | Out-Null
                    $bulkCopy.WriteToServer($bulkData)
                    Write-Host -Object "Successfully inserted $($bulkData.Rows.Count) rows into '$($DestinationTable)'." -ForegroundColor Green
                } catch {
                    Write-Host -Object "ERROR: Failed to bulk insert data. Error: $($_.Exception.Message)" -ForegroundColor Red
                } finally {
                    if ($bulkCopy) { $bulkCopy.Close() }
                }
            } else {
                Write-Host -Object "No new objects to insert for this batch." -ForegroundColor Cyan
            }

            $totalRowsProcessed += $sourceDataBatch.Count
            $offset += $BatchSize
            $batchNum++
        } 

        if ($ShowProgress) {
            Write-Progress -Activity "Parsing SQL Execution Plans" -Completed
        }

        Write-Host -Object "=================================================" -ForegroundColor Green
        Write-Host -Object "Script finished." -ForegroundColor Green
        Write-Host -Object "Total rows processed from source query: $totalRowsProcessed" -ForegroundColor Green
        Write-Host -Object "Total objects found and inserted: $totalObjectsFound" -ForegroundColor Green
        Write-Host -Object "=================================================" -ForegroundColor Green
    } catch {
        Write-Host -Object "A critical error occurred in the Main block: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host -Object "Script execution has been halted." -ForegroundColor Red
    }
}

# --- Execute the Main Function ---
Main
#endregion Main Script Body