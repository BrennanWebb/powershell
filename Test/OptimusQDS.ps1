<#
.SYNOPSIS
    A unified application for T-SQL tuning that performs a one-time setup on its first run and executes performance analysis on all subsequent runs.

.DESCRIPTION
    OptimusQDS is an automated T-SQL tuning advisor that leverages the Gemini AI.

    FIRST RUN: If the script detects that the Optimus database does not exist, it will run a one-time setup wizard. This creates the database, all required tables with cascading deletes, views, and securely prompts for the API key.

    SUBSEQUENT RUNS: The script generates a new RunID to batch all of its work. It scans QDS-enabled databases and identifies the worst-performing non-adhoc queries, logs analysis metadata, and then logs each individual tuning recommendation from Gemini.

.PARAMETER ServerName
    The FQDN or IP address of the target SQL Server instance. Will be prompted if not provided.

.PARAMETER DatabaseName
    The name of the Optimus application database. Defaults to "Optimus".

.PARAMETER MasterKeyPassword
    The password for the database master key, required to decrypt the API key during analysis runs. Will be prompted if not provided.

.PARAMETER MetricName
    The name of the performance metric from dbo.Config_Metric to use for identifying "worst offenders". Defaults to 'Memory Cost Score'.

.PARAMETER TopN
    The number of top queries to analyze for each database. Defaults to 5.

.PARAMETER DebugMode
    Enables verbose console output and writes detailed, relational logs to the dbo.Analysis_Debug table.

.NOTES
    Designer: Brennan Webb & Gemini
    Script Engine: Gemini
    Version: 3.5.5
    Created: 2025-07-23
    Modified: 2025-09-27
    Change Log:
    - v3.5.5: Updated default model names to stable, modern identifiers (gemini-2.5-flash).
              Reverted API URL construction in Invoke-GeminiAnalysis to use the colon-syntax (:) which is supported for most Gemini model names, resolving the 404 error caused by an unstable model name.
    - v3.5.4: Fixed the Gemini API URL to use '/' instead of ':' to resolve 404 errors. 
              Corrected the SQL UPDATE statement for Analysis_Meta to handle null/failed API responses, resolving the 'Incorrect syntax near WHERE' error.
              Updated default model names to stable identifiers (e.g., 'gemini-1.5-flash').
    - v3.5.3: Added ON DELETE CASCADE to all relevant foreign keys to simplify data cleanup.
    - v3.5.2: Removed the 24-hour time filter from Get-WorstOffenders to analyze the entire QDS retention period.
    - v3.5.1: Standardized the T-SQL formulas for all default metrics for consistency and accuracy.
    - v3.5.0: Removed redundant 'OriginatingDatabase' column from dbo.Analysis_Meta.
             Updated Get-WorstOffenders to only analyze queries with a valid object_id (non-adhoc queries).
    - v3.4.1: Fixed a SQL error in Get-WorstOffenders by correcting the function used to get the object's database name.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ServerName,

    [Parameter(Mandatory = $false)]
    [string]$DatabaseName = "Optimus",

    [Parameter(Mandatory = $false)]
    [System.Security.SecureString]$MasterKeyPassword,

    [Parameter(Mandatory = $false)]
    [string]$MetricName = 'Memory Cost Score',

    [Parameter(Mandatory = $false)]
    [int]$TopN = 5,

    [Parameter(Mandatory = $false)]
    [switch]$DebugMode
)

# --- Script-level Variables ---
$script:ScriptVersion = "3.5.5"

#region Helper Functions
function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $true)]
        [ValidateSet('INFO', 'SUCCESS', 'WARN', 'ERROR', 'PROMPT')]
        [string]$Level
    )
    $color = 'Cyan'
    switch ($Level) {
        'SUCCESS' { $color = 'Green' }
        'WARN'    { $color = 'Yellow' }
        'ERROR'   { $color = 'Red' }
        'PROMPT'  { $color = 'White' }
    }
    Write-Host "[$Level] $Message" -ForegroundColor $color
}

function Test-Prerequisites {
    Write-Log -Level 'INFO' -Message "Checking PowerShell version..."
    if ($PSVersionTable.PSVersion.Major -lt 5 -or ($PSVersionTable.PSVersion.Major -eq 5 -and $PSVersionTable.PSVersion.Minor -lt 1)) {
        Write-Log -Level 'ERROR' -Message "PowerShell 5.1 or higher is required. This script is running on version $($PSVersionTable.PSVersion)."
        return $false
    }
    Write-Log -Level 'SUCCESS' -Message "PowerShell version check passed ($($PSVersionTable.PSVersion))."
    return $true
}
#endregion

#region Setup and Initialization Functions
function Initialize-Database {
    param(
        [Parameter(Mandatory = $true)][string]$TargetServer,
        [Parameter(Mandatory = $true)][string]$DbName
    )
    
    $newPromptText = @"
### 1. Persona and Goal
You are a world-class database performance tuning expert for Microsoft SQL Server. Your sole task is to analyze the provided T-SQL execution plan and database object schemas to identify performance bottlenecks. You will then generate actionable, precise tuning recommendations.

### 2. Critical Output Rules
- Your ENTIRE response MUST be a single T-SQL block comment (`/* ... */`).
- DO NOT use markdown, conversational text, backticks, or any language outside of the T-SQL comment block.
- Each recommendation must start on a new numbered line (e.g., `1. `, `2. `).
- Each recommendation's title MUST follow the format: `[Category] - [Title]`.
- Use one of the following categories: Index, Syntax, Statistics, Architecture, Configuration.
- The body of each recommendation must strictly follow the 'a. Problem / b. Recommendation / c. Expected Result' format.

### 3. Example of Perfect Output (One-Shot Example)
This is the exact format you must follow.

/*
1. [Index] - [Costly Key Lookup on Sales.SalesOrderDetail]
    a. Problem: The query performs a Key Lookup to retrieve the 'OrderQty' and 'UnitPrice' columns. This happens because the existing non-clustered index `IX_SalesOrderDetail_ProductID` does not include these columns, forcing a second, expensive lookup into the clustered index for each qualifying row.
    b. Recommendation: Create a new non-clustered index that 'covers' the query.

       CREATE NONCLUSTERED INDEX [IX_SalesOrderDetail_ProductID_Covering]
       ON [Sales].[SalesOrderDetail] ([ProductID])
       INCLUDE ([OrderQty],[UnitPrice]);

    c. Expected Result: This new covering index will allow the server to satisfy the entire query with a single Index Seek operation, eliminating the costly Key Lookup and significantly reducing logical I/O and query duration.
*/

### 4. Instruction for No Findings
If you analyze the plan and find no significant performance improvements, your entire response MUST be the following T-SQL block comment:
/*
1. [General] - [No Tuning Recommendations Found]
    a. Problem: N/A
    b. Recommendation: No tuning recommendations found. The query plan appears optimal given the provided schemas.
    c. Expected Result: N/A
*/
"@.Replace("'", "''")

    $createTablesQuery = @"
USE [$DbName];
CREATE TABLE dbo.Config_Model (ModelID INT PRIMARY KEY IDENTITY, ModelName VARCHAR(100) UNIQUE NOT NULL, IsDefault BIT NOT NULL DEFAULT 0);
CREATE TABLE dbo.Config_Prompt (PromptID INT PRIMARY KEY IDENTITY, PromptName VARCHAR(100) UNIQUE NOT NULL, PromptText NVARCHAR(MAX) NOT NULL, IsDefault BIT NOT NULL DEFAULT 0);
CREATE TABLE dbo.Config_Metric (MetricID INT PRIMARY KEY IDENTITY, MetricName VARCHAR(100) UNIQUE NOT NULL, MetricDescription VARCHAR(500) NOT NULL, MetricFormula NVARCHAR(1000) NOT NULL, IsDefault BIT NOT NULL DEFAULT 0);
CREATE TABLE dbo.Config_ApiKey (ApiKeyID INT PRIMARY KEY IDENTITY, ApiKeyName VARCHAR(100) NOT NULL, ApiKeyValueEncrypted VARBINARY(256) NOT NULL);

CREATE TABLE dbo.Analysis_Run (RunID INT PRIMARY KEY IDENTITY, RunTimestamp DATETIME DEFAULT GETDATE());
CREATE TABLE dbo.Analysis_Meta (AnalysisMetaID INT PRIMARY KEY IDENTITY, RunID INT NOT NULL, MetricID INT NOT NULL, ModelID INT NOT NULL, PromptID INT NOT NULL, SourceQueryID BIGINT NOT NULL, ObjectDatabase sysname NULL, ObjectSchema sysname NULL, ObjectName sysname NULL, MetricValue BIGINT NULL, ObjectsParsed INT NULL, AnalysisDurationMS INT NULL, DatabaseRank INT NULL, ServerRank INT NULL, CONSTRAINT FK_Analysis_Meta_Run FOREIGN KEY (RunID) REFERENCES dbo.Analysis_Run(RunID) ON DELETE CASCADE, CONSTRAINT FK_Analysis_Meta_Metric FOREIGN KEY (MetricID) REFERENCES dbo.Config_Metric(MetricID), CONSTRAINT FK_Analysis_Meta_Model FOREIGN KEY (ModelID) REFERENCES dbo.Config_Model(ModelID), CONSTRAINT FK_Analysis_Meta_Prompt FOREIGN KEY (PromptID) REFERENCES dbo.Config_Prompt(PromptID));
CREATE TABLE dbo.Analysis_Result (AnalysisResultID INT PRIMARY KEY IDENTITY, AnalysisMetaID INT NOT NULL, TuningResponseType VARCHAR(50) NULL, TuningResponse NVARCHAR(MAX) NULL, CONSTRAINT FK_Analysis_Result_Meta FOREIGN KEY (AnalysisMetaID) REFERENCES dbo.Analysis_Meta(AnalysisMetaID) ON DELETE CASCADE);
CREATE TABLE dbo.Analysis_Debug (DebugID INT PRIMARY KEY IDENTITY, AnalysisMetaID INT NULL, DebugFunctionName VARCHAR(255) NOT NULL, Message VARCHAR(MAX) NOT NULL, LogTimestamp DATETIME DEFAULT GETDATE(), CONSTRAINT FK_Analysis_Debug_Meta FOREIGN KEY (AnalysisMetaID) REFERENCES dbo.Analysis_Meta(AnalysisMetaID) ON DELETE CASCADE);
"@

    $insertDataQuery = @"
USE [$DbName];
INSERT INTO dbo.Config_Model (ModelName, IsDefault) VALUES
('gemini-2.5-flash', 1),
('gemini-2.5-pro', 0),
('gemini-1.5-pro', 0);
INSERT INTO dbo.Config_Metric (MetricName, MetricDescription, MetricFormula, IsDefault) VALUES
('Memory Cost Score', 'A weighted score based on total duration, execution count, and total memory grant size.', '(SUM(rs.avg_duration) * SUM(rs.count_executions) * SUM(rs.avg_query_max_used_memory))', 1),
('Total CPU', 'Ranks queries by the total CPU time multiplied by execution count.', '(SUM(rs.avg_cpu_time) * SUM(rs.count_executions))', 0),
('Total Duration', 'Ranks queries by the total duration multiplied by execution count.', '(SUM(rs.avg_duration) * SUM(rs.count_executions))', 0),
('Total Logical Reads', 'Ranks queries by the total logical I/O multiplied by execution count.', '(SUM(rs.avg_logical_io_reads) * SUM(rs.count_executions))', 0),
('Max Memory Grant KB', 'Ranks queries by the largest single memory grant requested.', '(MAX(rs.max_query_max_used_memory))', 0),
('Total TempDb Spills KB', 'Ranks queries by the total amount of data spilled to TempDb.', '(SUM(rs.total_spills_kb))', 0);
INSERT INTO dbo.Config_Prompt (PromptName, PromptText, IsDefault) VALUES ('Standard Query Tuning', '$newPromptText', 1);
"@
    
    $createViewQuery = @"
CREATE VIEW dbo.vw_Analysis_Summary
AS
SELECT
    run.RunID,
    run.RunTimestamp,
    r.AnalysisResultID,
    m.AnalysisMetaID,
    m.SourceQueryID,
    m.ObjectDatabase,
    m.ObjectSchema,
    m.ObjectName,
    m.ServerRank,
    m.DatabaseRank,
    cmet.MetricName,
    m.MetricValue,
    m.ObjectsParsed,
    m.AnalysisDurationMS,
    cmod.ModelName,
    cp.PromptName,
    r.TuningResponseType,
    r.TuningResponse
FROM
    dbo.Analysis_Result AS r
    JOIN dbo.Analysis_Meta AS m ON r.AnalysisMetaID = m.AnalysisMetaID
    JOIN dbo.Analysis_Run AS run ON m.RunID = run.RunID
    JOIN dbo.Config_Model AS cmod ON m.ModelID = cmod.ModelID
    JOIN dbo.Config_Prompt AS cp ON m.PromptID = cp.PromptID
    JOIN dbo.Config_Metric AS cmet ON m.MetricID = cmet.MetricID;
"@

    try {
        Write-Log -Level 'INFO' -Message "Creating database '[$DbName]'..."
        Invoke-Sqlcmd -ServerInstance $TargetServer -Database "master" -Query "CREATE DATABASE [$DbName];" -TrustServerCertificate -ErrorAction Stop
        Write-Log -Level 'SUCCESS' -Message "Database '[$DbName]' created successfully."

        Write-Log -Level 'INFO' -Message "Creating tables in '[$DbName]'..."
        Invoke-Sqlcmd -ServerInstance $TargetServer -Database $DbName -Query $createTablesQuery -TrustServerCertificate -ErrorAction Stop
        
        Write-Log -Level 'INFO' -Message "Populating default data in '[$DbName]'..."
        Invoke-Sqlcmd -ServerInstance $TargetServer -Database $DbName -Query $insertDataQuery -TrustServerCertificate -ErrorAction Stop
        
        Write-Log -Level 'INFO' -Message "Creating summary view in '[$DbName]'..."
        Invoke-Sqlcmd -ServerInstance $TargetServer -Database $DbName -Query $createViewQuery -TrustServerCertificate -ErrorAction Stop
        
        Write-Log -Level 'SUCCESS' -Message "Schema, data, and view created successfully."
        return $true
    }
    catch {
        Write-Log -Level 'ERROR' -Message "Failed during database initialization. Error: $($_.Exception.Message)"
        return $false
    }
}

function Set-InitialApiKey {
    param(
        [Parameter(Mandatory = $true)][string]$TargetServer,
        [Parameter(Mandatory = $true)][string]$DbName
    )
    # This function remains unchanged from the previous version
    Write-Log -Level 'INFO' -Message "--- Initial API Key Setup ---"
    Write-Log -Level 'PROMPT' -Message "To secure the API key, a Database Master Key will be created."
    $masterKeyPassword = $null
    while ($true) {
        Write-Log -Level 'PROMPT' -Message "Please enter a strong password for the new Master Key:"
        $p1 = Read-Host -AsSecureString
        Write-Log -Level 'PROMPT' -Message "Please confirm the password:"
        $p2 = Read-Host -AsSecureString
        $bstr1 = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($p1); $bstr2 = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($p2)
        $plainP1 = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr1); $plainP2 = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr2)
        if ($plainP1 -ceq $plainP2) { $masterKeyPassword = $p1; [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr2); break }
        else { Write-Log -Level 'WARN' -Message "Passwords do not match. Please try again."; [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr1); [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr2) }
    }
    Write-Log -Level 'PROMPT' -Message "Please enter the Gemini API Key value:"; $apiKeySecure = Read-Host -AsSecureString
    if ($apiKeySecure.Length -eq 0) { Write-Log -Level 'ERROR' -Message "API Key cannot be empty."; return $false }
    $bstrPass = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($masterKeyPassword); $plainTextPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstrPass)
    $bstrApi = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($apiKeySecure); $plainTextApiKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstrApi)
    $sanitizedPassword = $plainTextPassword.Replace("'", "''"); $sanitizedApiKey = $plainTextApiKey.Replace("'", "''")
    $keySetupQuery = "USE [$DbName]; CREATE MASTER KEY ENCRYPTION BY PASSWORD = '$sanitizedPassword'; CREATE CERTIFICATE Optimus_Cert WITH SUBJECT = 'Optimus Certificate'; INSERT INTO dbo.Config_ApiKey (ApiKeyName, ApiKeyValueEncrypted) VALUES ('Default Gemini Key', ENCRYPTBYCERT(CERT_ID('Optimus_Cert'), '$sanitizedApiKey'));"
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstrPass); [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstrApi); Remove-Variable plain*, san*, p1, p2, masterKey*, apiKeySecure -ErrorAction SilentlyContinue
    try {
        Write-Log -Level 'INFO' -Message "Creating master key, certificate, and storing encrypted API key..."
        Invoke-Sqlcmd -ServerInstance $TargetServer -Database $DbName -Query $keySetupQuery -TrustServerCertificate -ErrorAction Stop
        Write-Log -Level 'SUCCESS' -Message "API Key has been successfully encrypted and stored."
        return $true
    }
    catch {
        Write-Log -Level 'ERROR' -Message "Failed during API key setup. You may need to manually clean up objects. Error: $($_.Exception.Message)"
        return $false
    }
}
#endregion

#region Core Analysis Functions
function Initialize-DebugSession {
    if ($DebugMode) {
        Write-Log -Level 'WARN' -Message "Debug mode is active. Truncating the debug log table."
        $query = "TRUNCATE TABLE dbo.Analysis_Debug;"
        try {
            # We don't log this query itself, as the log is what's being cleared.
            Invoke-Sqlcmd -ServerInstance $ServerName -Database $DatabaseName -Query $query -TrustServerCertificate
        }
        catch {
            Write-Log -Level 'ERROR' -Message "Failed to truncate the debug log. Error: $($_.Exception.Message)"
        }
    }
}

function Write-DebugLog {
    param($AnalysisMetaID, $Message)
    
    if ($DebugMode) {
        $callerFunctionName = ((Get-PSCallStack)[1].FunctionName -split ':')[-1]
        if ([string]::IsNullOrWhiteSpace($callerFunctionName)) { $callerFunctionName = "Main" }

        $metaIdToLog = if ($null -eq $AnalysisMetaID -or $AnalysisMetaID -eq 0) { "NULL" } else { $AnalysisMetaID }
        $sanitizedFunction = $callerFunctionName.Replace("'", "''")
        $sanitizedMessage = $Message.Replace("'", "''")
        $query = "INSERT INTO dbo.Analysis_Debug (AnalysisMetaID, DebugFunctionName, Message) VALUES ($metaIdToLog, '$sanitizedFunction', '$sanitizedMessage');"
        try {
            Invoke-Sqlcmd -ServerInstance $ServerName -Database $DatabaseName -Query $query -TrustServerCertificate -ErrorAction Stop
        }
        catch {
            Write-Host "[CRITICAL] Failed to write to the debug log table: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

function Start-AnalysisRun {
    $query = "INSERT INTO dbo.Analysis_Run DEFAULT VALUES; SELECT SCOPE_IDENTITY();"
    try {
        Write-DebugLog -AnalysisMetaID 0 -Message "Executing SQL: $query"
        $runId = (Invoke-Sqlcmd -ServerInstance $ServerName -Database $DatabaseName -Query $query -TrustServerCertificate).Item(0)
        Write-Log -Level 'SUCCESS' -Message "Current analysis batch is RunID: $runId."
        return $runId
    }
    catch {
        Write-Log -Level 'ERROR' -Message "Failed to create a new analysis run ID. Error: $($_.Exception.Message)"
        return $null
    }
}

function Get-Configuration {
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($MasterKeyPassword); $plainTextPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr); [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    $sanitizedPassword = $plainTextPassword.Replace("'", "''")
    $metricLookupQuery = "SELECT TOP 1 MetricName FROM dbo.Config_Metric WHERE IsDefault = 1;"
    Write-DebugLog -AnalysisMetaID 0 -Message "Executing SQL: $metricLookupQuery"
    $metricLookup = if ($PSBoundParameters.ContainsKey('MetricName')) { $MetricName } else { (Invoke-Sqlcmd -ServerInstance $ServerName -Database $DatabaseName -Query $metricLookupQuery -TrustServerCertificate).MetricName }
    
    $configQuery = @"
USE [$DatabaseName];
SET NOCOUNT ON;
OPEN MASTER KEY DECRYPTION BY PASSWORD = '$sanitizedPassword';
SELECT TOP 1 CONVERT(VARCHAR(255), DECRYPTBYCERT(CERT_ID('Optimus_Cert'), ApiKeyValueEncrypted)) as ApiKey FROM dbo.Config_ApiKey;
CLOSE MASTER KEY;
SELECT ModelID, ModelName FROM dbo.Config_Model WHERE IsDefault = 1;
SELECT PromptID, PromptText FROM dbo.Config_Prompt WHERE IsDefault = 1;
SELECT MetricID, MetricFormula FROM dbo.Config_Metric WHERE MetricName = '$($metricLookup.Replace("'", "''"))';
"@
    try {
        Write-DebugLog -AnalysisMetaID 0 -Message "Executing SQL to retrieve full configuration..."
        $ds = Invoke-Sqlcmd -ServerInstance $ServerName -Database $DatabaseName -Query $configQuery -TrustServerCertificate -ErrorAction Stop
        $script:Configuration = @{
            ApiKey        = $ds[0].ApiKey; ModelID = $ds[1].ModelID; ModelName = $ds[1].ModelName
            PromptID      = $ds[2].PromptID; PromptText = $ds[2].PromptText
            MetricID      = $ds[3].MetricID; MetricFormula = $ds[3].MetricFormula
        }
        $script:MetricName = $metricLookup
        if ([string]::IsNullOrWhiteSpace($script:Configuration.ApiKey)) { throw "API Key could not be decrypted. Check password." }
        if ($null -eq $script:Configuration.MetricID) { throw "Metric '$metricLookup' not found." }
        Write-Log -Level 'SUCCESS' -Message "Configuration loaded successfully."
        return $true
    }
    catch { Write-Log -Level 'ERROR' -Message "Failed to retrieve configuration. Error: $($_.Exception.Message)"; return $false }
}

function Get-TargetDatabases {
    $query = "SELECT name FROM sys.databases WHERE is_query_store_on = 1 AND database_id > 4 AND state_desc = 'ONLINE';"
    try {
        Write-DebugLog -AnalysisMetaID 0 -Message "Executing SQL: $query"
        $dbs = Invoke-Sqlcmd -ServerInstance $ServerName -Database "master" -Query $query -TrustServerCertificate -ErrorAction Stop
        Write-Log -Level 'SUCCESS' -Message "Found $($dbs.Count) target databases."
        return $dbs.name
    }
    catch { Write-Log -Level 'ERROR' -Message "Failed to query for target databases. Error: $($_.Exception.Message)"; return @() }
}

function Get-WorstOffenders {
    param($TargetDb)
    $query = "
SET NOCOUNT ON; 
WITH R AS (
    SELECT 
        q.query_id, 
        qt.query_sql_text, 
        p.query_plan,
        DB_NAME() AS ObjectDatabase,
        s.name AS ObjectSchema,
        o.name AS ObjectName,
        $($script:Configuration.MetricFormula) AS MetricValue, 
        ROW_NUMBER() OVER(ORDER BY $($script:Configuration.MetricFormula) DESC) as rn 
    FROM sys.query_store_query AS q 
    JOIN sys.query_store_query_text AS qt ON q.query_text_id = qt.query_text_id 
    JOIN sys.query_store_plan AS p ON q.query_id = p.query_id 
    JOIN sys.query_store_runtime_stats AS rs ON p.plan_id = rs.plan_id 
    LEFT JOIN sys.objects AS o ON q.object_id = o.object_id
    LEFT JOIN sys.schemas AS s ON o.schema_id = s.schema_id
    WHERE 
        q.object_id <> 0
    GROUP BY q.query_id, qt.query_sql_text, p.query_plan, s.name, o.name
) 
SELECT TOP ($TopN) query_id, query_sql_text, query_plan, MetricValue, ObjectDatabase, ObjectSchema, ObjectName 
FROM R 
WHERE rn <= $TopN 
ORDER BY rn;
"
    try {
        Write-DebugLog -AnalysisMetaID 0 -Message "Executing SQL in '$TargetDb': $query"
        $offenders = Invoke-Sqlcmd -ServerInstance $ServerName -Database $TargetDb -Query $query -TrustServerCertificate -MaxCharLength ([int]::MaxValue)
        Write-Log -Level 'SUCCESS' -Message "Identified $($offenders.Count) offenders in '$TargetDb'."
        return $offenders
    }
    catch { Write-Log -Level 'WARN' -Message "Could not retrieve offenders from '$TargetDb'. It may have no recent QDS data."; return @() }
}

function Get-ObjectsFromPlan {
    param($ExecutionPlanXml)
    # This is a local operation, no SQL to log.
    $ns = New-Object System.Xml.XmlNamespaceManager($ExecutionPlanXml.NameTable); $ns.AddNamespace("sql", "http://schemas.microsoft.com/sqlserver/2004/07/showplan")
    $objectNodes = $ExecutionPlanXml.SelectNodes("//sql:Object", $ns); $uniqueObjectNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($node in $objectNodes) {
        $db = $node.GetAttribute("Database"); $schema = $node.GetAttribute("Schema"); $table = $node.GetAttribute("Table")
        if ($table -notlike "#*" -and -not ([string]::IsNullOrWhiteSpace($db)) -and -not ([string]::IsNullOrWhiteSpace($schema)) -and $db -ne '[mssqlsystemresource]') {
            $uniqueObjectNames.Add("$db.$schema.$table".Replace('[', '').Replace(']', '')) | Out-Null
        }
    }
    return $uniqueObjectNames | Sort-Object
}

function Get-ObjectSchemaDetails {
    param($FullObjectName, $AnalysisMetaID)
    $parts = $FullObjectName.Split('.'); $dbName = $parts[0]; $schemaName = $parts[1]; $objName = $parts[2]
    $twoPartName = "[$schemaName].[$objName]"; $sanitizedTwoPartName = $twoPartName.Replace("'", "''")
    $schemaText = "--- Schema For: [$dbName].$twoPartName ---`n"
    $colQuery = "SELECT c.name, t.name AS type, c.max_length, c.is_nullable FROM sys.columns c JOIN sys.types t ON c.user_type_id = t.user_type_id WHERE c.object_id = OBJECT_ID('$sanitizedTwoPartName') ORDER BY c.column_id;"
    $idxQuery = "SELECT i.name AS IndexName, i.type_desc AS IndexType, STUFF((SELECT ', ' + c.name FROM sys.index_columns ic JOIN sys.columns c ON ic.object_id = c.object_id AND ic.column_id = c.column_id WHERE ic.object_id = i.object_id AND ic.index_id = i.index_id ORDER BY ic.key_ordinal FOR XML PATH('')), 1, 2, '') AS KeyColumns FROM sys.indexes i WHERE i.object_id = OBJECT_ID('$sanitizedTwoPartName');"
    try { Write-DebugLog -AnalysisMetaID $AnalysisMetaID -Message "Executing SQL for Columns in '$dbName': $colQuery"; $schemaText += "COLUMNS:`n" + (Invoke-Sqlcmd -ServerInstance $ServerName -Database $dbName -Query $colQuery -TrustServerCertificate -ErrorAction Stop | ForEach-Object { "- $($_.name) $($_.type) $(if ($_.is_nullable) {'NULL'} else {'NOT NULL'})" }) -join "`n" } catch {}
    try { Write-DebugLog -AnalysisMetaID $AnalysisMetaID -Message "Executing SQL for Indexes in '$dbName': $idxQuery"; $schemaText += "`n`nINDEXES:`n" + (Invoke-Sqlcmd -ServerInstance $ServerName -Database $dbName -Query $idxQuery -TrustServerCertificate -ErrorAction Stop | ForEach-Object { "- $($_.IndexName) ($($_.IndexType)) ON ($($_.KeyColumns))" }) -join "`n" } catch {}
    return $schemaText + "`n"
}

function Invoke-GeminiAnalysis {
    param($QuerySqlText, $ExecutionPlan, $SchemaDetails, $AnalysisMetaID)
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    # FIX: Reverting to colon-syntax as it is the format used in the working client script, but ensuring stable model names are used.
    $uri = "https://generativelanguage.googleapis.com/v1beta/models/$($script:Configuration.ModelName):generateContent?key=$($script:Configuration.ApiKey)"
    $fullPrompt = "$($script:Configuration.PromptText)`n`n### T-SQL QUERY`n$QuerySqlText`n`n### EXECUTION PLAN XML`n$ExecutionPlan`n`n### OBJECT SCHEMAS`n$SchemaDetails"
    Write-DebugLog -AnalysisMetaID $AnalysisMetaID -Message "Sending Gemini Payload: $fullPrompt"
    $bodyObject = @{ contents = @( @{ parts = @( @{ text = $fullPrompt } ) } ) }; $bodyJson = $bodyObject | ConvertTo-Json -Depth 10
    try {
        $response = Invoke-RestMethod -Uri $uri -Method Post -Body $bodyJson -ContentType 'application/json' -ErrorAction Stop -TimeoutSec 300
        $rawAiResponse = $response.candidates[0].content.parts[0].text
        $stopwatch.Stop()
        Write-Log -Level 'SUCCESS' -Message "Successfully received analysis from Gemini AI in $($stopwatch.Elapsed.TotalSeconds) seconds."
        return @{ Response = $rawAiResponse.Trim(); DurationMS = [int]$stopwatch.Elapsed.TotalMilliseconds }
    }
    catch { 
        Write-Log -Level 'ERROR' -Message "Failed to get response from Gemini API. Error: $($_.Exception.Message)"
        Write-Log -Level 'DEBUG' -Message "Attempted URI: $uri" 
        return $null 
    }
}

function Write-AnalysisMeta {
    param($RunID, $QueryId, $ObjectDatabase, $ObjectSchema, $ObjectName, $MetricValue, $ObjectsParsed, $AnalysisDurationMS)
    $insertQuery = "INSERT INTO dbo.Analysis_Meta (RunID, MetricID, ModelID, PromptID, SourceQueryID, ObjectDatabase, ObjectSchema, ObjectName, MetricValue, ObjectsParsed, AnalysisDurationMS) OUTPUT INSERTED.AnalysisMetaID VALUES ($RunID, $($script:Configuration.MetricID), $($script:Configuration.ModelID), $($script:Configuration.PromptID), $QueryId, '$ObjectDatabase', '$ObjectSchema', '$ObjectName', $MetricValue, $ObjectsParsed, $AnalysisDurationMS);"
    try {
        Write-DebugLog -AnalysisMetaID 0 -Message "Executing SQL: $insertQuery"
        $metaId = (Invoke-Sqlcmd -ServerInstance $ServerName -Database $DatabaseName -Query $insertQuery -TrustServerCertificate).Item(0)
        Write-Log -Level 'INFO' -Message "Created meta record with ID $metaId for Query ID $QueryId."
        return $metaId
    }
    catch { Write-Log -Level 'ERROR' -Message "Failed to write analysis metadata to the database. Error: $($_.Exception.Message)"; return $null }
}

function Write-AnalysisResult {
    param($AnalysisMetaID, $TuningResponseType, $TuningResponse)
    $sanitizedType = $TuningResponseType.Replace("'", "''")
    $sanitizedResponse = $TuningResponse.Replace("'", "''")
    $insertQuery = "INSERT INTO dbo.Analysis_Result (AnalysisMetaID, TuningResponseType, TuningResponse) VALUES ($AnalysisMetaID, '$sanitizedType', '$sanitizedResponse');"
    try {
        Write-DebugLog -AnalysisMetaID $AnalysisMetaID -Message "Executing SQL: $insertQuery"
        Invoke-Sqlcmd -ServerInstance $ServerName -Database $DatabaseName -Query $insertQuery -TrustServerCertificate -ErrorAction Stop
        Write-Log -Level 'SUCCESS' -Message "Successfully logged [$sanitizedType] recommendation for Meta ID $AnalysisMetaID."
    }
    catch { Write-Log -Level 'ERROR' -Message "Failed to write analysis result to the database. Error: $($_.Exception.Message)" }
}

function Update-AnalysisRanks {
    param ($RunID)
    $query = @"
WITH RanksToUpdate AS (
    SELECT
        AnalysisMetaID,
        ROW_NUMBER() OVER (PARTITION BY MetricID, ObjectDatabase ORDER BY MetricValue DESC) AS DbRank,
        ROW_NUMBER() OVER (PARTITION BY MetricID ORDER BY MetricValue DESC) AS SrvRank
    FROM dbo.Analysis_Meta
    WHERE RunID = $RunID
)
UPDATE m
SET
    m.DatabaseRank = r.DbRank,
    m.ServerRank = r.SrvRank
FROM
    dbo.Analysis_Meta AS m
    JOIN RanksToUpdate AS r ON m.AnalysisMetaID = r.AnalysisMetaID;
"@
    try {
        Write-DebugLog -AnalysisMetaID 0 -Message "Executing SQL: $query"
        Invoke-Sqlcmd -ServerInstance $ServerName -Database $DatabaseName -Query $query -TrustServerCertificate -ErrorAction Stop
        Write-Log -Level 'SUCCESS' -Message "Successfully updated ranks for RunID: $RunID."
    }
    catch {
        Write-Log -Level 'ERROR' -Message "Failed to update ranks. Error: $($_.Exception.Message)"
    }
}
#endregion

# --- Script Entry Point ---
Write-Host "OptimusQDS v$($script:ScriptVersion)" -ForegroundColor Green

if (-not (Test-Prerequisites)) { exit 1 }

if (-not $PSBoundParameters.ContainsKey('ServerName')) {
    Write-Log -Level 'PROMPT' -Message "Please enter the target SQL Server instance name:"; $ServerName = Read-Host
}
if ([string]::IsNullOrWhiteSpace($ServerName)) { Write-Log -Level 'ERROR' -Message "ServerName cannot be empty."; exit 1 }

Write-Log -Level 'INFO' -Message "--- Starting OptimusQDS for server: $ServerName ---"

try {
    $dbCheckQuery = "SELECT name FROM sys.databases WHERE name = '$DatabaseName'"
    $dbCheck = Invoke-Sqlcmd -ServerInstance $ServerName -Database "master" -Query $dbCheckQuery -TrustServerCertificate -ErrorAction Stop
}
catch {
    Write-Log -Level 'ERROR' -Message "Could not connect to server '$ServerName'. Please verify the name and your permissions."; exit 1
}

if ($null -eq $dbCheck) {
    # --- FIRST RUN SETUP MODE ---
    Initialize-Database -TargetServer $ServerName -DbName $DatabaseName
    Set-InitialApiKey -TargetServer $ServerName -DbName $DatabaseName
    Write-Log -Level 'SUCCESS' -Message "--- Optimus setup complete! Please run the script again to begin an analysis. ---"
    exit 0
}
else {
    # --- SUBSEQUENT RUN ANALYSIS MODE ---
    if (-not $PSBoundParameters.ContainsKey('MasterKeyPassword')) {
        Write-Log -Level 'PROMPT' -Message "Please enter the Master Key password for the Optimus database:"; $MasterKeyPassword = Read-Host -AsSecureString
    }
    if ($MasterKeyPassword.Length -eq 0) { Write-Log -Level 'ERROR' -Message "MasterKeyPassword cannot be empty for analysis runs."; exit 1 }
    
    if (-not (Get-Configuration)) { exit 1 }
    
    Initialize-DebugSession 

    $runId = Start-AnalysisRun
    if ($null -eq $runId) { Write-Log -Level 'ERROR' -Message "Could not establish a RunID. Halting analysis."; exit 1 }

    $targetDatabases = Get-TargetDatabases
    foreach ($db in $targetDatabases) {
        $worstOffenders = Get-WorstOffenders -TargetDb $db
        foreach ($offender in $worstOffenders) {
            $metaId = Write-AnalysisMeta -RunID $runId -QueryId $offender.query_id -ObjectDatabase $offender.ObjectDatabase -ObjectSchema $offender.ObjectSchema -ObjectName $offender.ObjectName -MetricValue $offender.MetricValue -ObjectsParsed 0 -AnalysisDurationMS 0
            if ($null -eq $metaId) { Write-Log -Level 'WARN' -Message "Skipping Query ID $($offender.query_id) due to metadata insertion failure."; continue }

            $planObjects = @(); $schemaDetails = ""
            try {
                $planObjects = Get-ObjectsFromPlan -ExecutionPlanXml ([xml]$offender.query_plan)
                Write-DebugLog -AnalysisMetaID $metaId -Message "Successfully parsed $($planObjects.Count) objects from plan. Starting schema collection."
                $schemaDetails = foreach ($obj in $planObjects) { Get-ObjectSchemaDetails -FullObjectName $obj -AnalysisMetaID $metaId }
                $schemaDetails = $schemaDetails -join "`n"
                Write-DebugLog -AnalysisMetaID $metaId -Message "Schema collection complete. Payload size is $($schemaDetails.Length) characters."
            }
            catch {
                Write-Log -Level 'WARN' -Message "Could not parse execution plan for Query ID $($offender.query_id). Skipping schema collection."
                Write-DebugLog -AnalysisMetaID $metaId -Message "FATAL: Could not parse execution plan XML. Error: $($_.Exception.Message)"
            }

            $analysisResult = Invoke-GeminiAnalysis -QuerySqlText $offender.query_sql_text -ExecutionPlan $offender.query_plan -SchemaDetails $schemaDetails -AnalysisMetaID $metaId
            
            # Safely get duration, defaulting to 0 if API call failed
            $durationToLog = if ($null -ne $analysisResult) { $analysisResult.DurationMS } else { 0 }

            if ($null -ne $metaId) { # Only attempt to update if meta record was created successfully
                $updateMetaQuery = "UPDATE dbo.Analysis_Meta SET ObjectsParsed = $($planObjects.Count), AnalysisDurationMS = $durationToLog WHERE AnalysisMetaID = $metaId;"
                Write-DebugLog -AnalysisMetaID $metaId -Message "Executing SQL: $updateMetaQuery"
                try {
                    Invoke-Sqlcmd -ServerInstance $ServerName -Database $DatabaseName -Query $updateMetaQuery -TrustServerCertificate -ErrorAction Stop
                }
                catch {
                    Write-Log -Level 'ERROR' -Message "Failed to update AnalysisMetaID $metaId. Error: $($_.Exception.Message)"
                }
            }

            if ($null -ne $analysisResult -and -not [string]::IsNullOrWhiteSpace($analysisResult.Response)) {
                $cleanedResponse = $analysisResult.Response.Trim().Trim("/*").Trim("*/").Trim()
                $recommendations = $cleanedResponse -split '(?m)^\d+\.\s' | Where-Object { $_.Trim() -ne "" }
                Write-DebugLog -AnalysisMetaID $metaId -Message "Parsing $($recommendations.Count) recommendations from AI response."

                foreach ($rec in $recommendations) {
                    $type = "General"; $text = $rec.Trim()
                    if ($text -match '(?s)\[(.*?)\]\s-\s(.*)') {
                        $type = $matches[1].Trim(); $text = $matches[2].Trim()
                    }
                    Write-AnalysisResult -AnalysisMetaID $metaId -TuningResponseType $type -TuningResponse $text
                }
            }
            else {
                Write-Log -Level 'WARN' -Message "No tuning response received from Gemini for Query ID $($offender.query_id)."
                Write-DebugLog -AnalysisMetaID $metaId -Message "No valid tuning response was returned from the API."
            }
        }
    }
    Update-AnalysisRanks -RunID $runId
    Write-Log -Level 'SUCCESS' -Message "--- Optimus Analysis Run Finished ---"
}