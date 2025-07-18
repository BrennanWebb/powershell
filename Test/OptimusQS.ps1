<#
.SYNOPSIS
    OptimusQS is an automated T-SQL tuning advisor for Query Store that leverages the Gemini AI.

.DESCRIPTION
    This script runs non-interactively, connecting to a specified SQL Server to read query data
    (specifically the execution plan) from a designated table. For each query, it extracts the
    referenced database objects, collects their complete schemas (tables, columns, indexes, constraints),
    and sends this comprehensive data to the Gemini AI for performance analysis. The tuning
    recommendations are then written back to the source table. All actions and errors are logged
    to a separate, dedicated logging table in the database.

.PARAMETER ServerName
    The FQDN or IP address of the target SQL Server instance (e.g., "PROD-DB01\SQL2022"). This parameter is mandatory.

.PARAMETER FullObjectName
    The three-part name of the source table containing the Query Store data and the target column
    for the tuning response. Example: "BI_Monitoring.dbo.QueryStore_OptimusQS". This parameter is mandatory.

.PARAMETER LogTableName
    The three-part name of the table where all operational logs (INFO, WARN, ERROR) will be written.
    Example: "BI_Monitoring.dbo.OptimusQS_Log". This parameter is mandatory.

.PARAMETER ConfigTableName
    The three-part name of the table that stores the encrypted Gemini API key.
    Example: "BI_Monitoring.dbo.OptimusQS_Config". This parameter is mandatory.

.PARAMETER MasterKeyPassword
    The password for the database master key, required to decrypt the API key stored in the ConfigTableName.
    This parameter is mandatory and should be handled securely, ideally passed from a credential store.

.PARAMETER MaxRetries
    The number of times the script should re-run its analysis loop for any queries that failed on the previous pass.
    Defaults to 3.

.PARAMETER GeminiModel
    An optional parameter to specify which Gemini model to use for the analysis.
    Defaults to 'gemini-1.5-flash-latest'.

.PARAMETER DebugMode
    An optional switch that enables verbose, color-coded output to the console.

.EXAMPLE
    # Example for execution from a command line or SQL Agent Job (CmdExec step)
    $password = "YourMasterKeyPassword" | ConvertTo-SecureString -AsPlainText -Force
    
    .\OptimusQS.ps1 -ServerName "Warehouse.SelectQuote.com" `
                    -FullObjectName "BI_Monitoring.dbo.QueryStore_OptimusQS" `
                    -LogTableName "BI_Monitoring.dbo.OptimusQS_Log" `
                    -ConfigTableName "BI_Monitoring.dbo.OptimusQS_Config" `
                    -MasterKeyPassword $password `
                    -MaxRetries 3 `
                    -DebugMode

.NOTES
    Designer: Brennan Webb & Gemini
    Script Engine: Gemini
    Version: 1.1.1
    Created: 2025-07-18
    Modified: 2025-07-18
    Change Log:
    - v1.1.1: Increased AI response character limit from 4000 to 8000 in both the prompt and the truncation logic.
    - v1.1.0: Corrected the AI prompt to include the $SqlVersion variable.
    - v1.0.9: Overhauled the AI prompt to enforce a T-SQL comment block for the output format.
    - v1.0.8: Implemented a main retry loop controlled by a -MaxRetries parameter.
    - v1.0.7: Increased API call timeout in Invoke-GeminiAnalysis from 300 to 600 seconds.
    - v1.0.6: Changed Get-ObjectSchema to embed sanitized object names directly in queries.
    - v1.0.5: Corrected Get-ObjectSchema to use a two-part object name for lookups.
    - v1.0.4: Added -MaxCharLength to the main data retrieval query.
    - v1.0.3: Added a -Database parameter to the Invoke-Sqlcmd call in Get-ApiKeyFromDatabase.
    - v1.0.2: Corrected Invoke-Sqlcmd calls to properly handle parameters by embedding sanitized values.
    - v1.0.1: Removed unused variable '$dbName' to resolve PSScriptAnalyzer warning.
    Powershell Version: 5.1+
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$ServerName,

    [Parameter(Mandatory = $true)]
    [string]$FullObjectName,

    [Parameter(Mandatory = $true)]
    [string]$LogTableName,

    [Parameter(Mandatory = $true)]
    [string]$ConfigTableName,

    [Parameter(Mandatory = $true)]
    [System.Security.SecureString]$MasterKeyPassword,
    
    [Parameter(Mandatory = $false)]
    [int]$MaxRetries = 3,

    [Parameter(Mandatory = $false)]
    [ValidateSet('gemini-1.5-flash-latest', 'gemini-2.5-flash', 'gemini-2.5-pro')]
    [string]$GeminiModel = 'gemini-1.5-flash-latest',

    [Parameter(Mandatory = $false)]
    [switch]$DebugMode
)

# --- Script-level Variables ---
$script:GeminiApiKey = $null
$script:ScriptVersion = "1.1.1"

#region Logging and Prerequisite Functions

function Write-SqlLog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $true)]
        [ValidateSet('INFO', 'SUCCESS', 'WARN', 'ERROR')]
        [string]$Level
    )

    $fullMessage = "OptimusQS v$($script:ScriptVersion): $($Message -replace "'", "''")"
    $logQuery = "INSERT INTO $LogTableName (LogLevel, LogMessage) VALUES ('$Level', '$fullMessage');"
    
    try {
        Invoke-Sqlcmd -ServerInstance $ServerName -Query $logQuery -TrustServerCertificate -ErrorAction Stop
    }
    catch {
        Write-Host "CRITICAL LOGGING FAILURE: Could not write to table '$LogTableName'. Error: $($_.Exception.Message)" -ForegroundColor Red
    }

    if ($DebugMode) {
        $color = 'Cyan'
        switch ($Level) {
            'SUCCESS' { $color = 'Green' }
            'WARN'    { $color = 'Yellow' }
            'ERROR'   { $color = 'Red' }
        }
        Write-Host "[$Level] $($Message)" -ForegroundColor $color
    }
}

function Test-SqlServerModule {
    Write-SqlLog -Level 'INFO' -Message "Checking for 'SqlServer' PowerShell module."
    if (Get-Module -Name SqlServer -ListAvailable) {
        try {
            Import-Module SqlServer -ErrorAction Stop
            Write-SqlLog -Level 'SUCCESS' -Message "'SqlServer' module imported successfully."
            return $true
        }
        catch {
            Write-SqlLog -Level 'ERROR' -Message "Fatal: Failed to import 'SqlServer' module. Error: $($_.Exception.Message)"
            return $false
        }
    }
    else {
        Write-SqlLog -Level 'ERROR' -Message "Fatal: The 'SqlServer' PowerShell module is not installed. Please run 'Install-Module -Name SqlServer' and try again."
        return $false
    }
}

#endregion

#region Core SQL, Schema, and AI Functions

function Get-ApiKeyFromDatabase {
    Write-SqlLog -Level 'INFO' -Message "Attempting to retrieve Gemini API Key from database."
    
    try {
        $dbNameForApiKey = ($ConfigTableName -split '\.')[0].Trim('[]')
        Write-SqlLog -Level 'INFO' -Message "API Key Context: Using database '$dbNameForApiKey'."
    } catch {
        Write-SqlLog -Level 'ERROR' -Message "Fatal: Could not parse database name from '$ConfigTableName'. Please provide a valid three-part name (e.g., 'YourDB.dbo.YourTable')."
        return $false
    }

    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($MasterKeyPassword)
    $plainTextPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)

    $sanitizedPassword = $plainTextPassword.Replace("'", "''")

    $apiKeyQuery = @"
OPEN MASTER KEY DECRYPTION BY PASSWORD = '$sanitizedPassword';
SELECT TOP 1 CONVERT(VARCHAR(255), DECRYPTBYCERT(CERT_ID('OptimusQS_Cert'), ConfigValueEncrypted)) as ApiKey
FROM $ConfigTableName
WHERE ConfigKey = 'GeminiApiKey';
CLOSE MASTER KEY;
"@

    try {
        $result = Invoke-Sqlcmd -ServerInstance $ServerName -Database $dbNameForApiKey -Query $apiKeyQuery -TrustServerCertificate -ErrorAction Stop
        if ($null -ne $result -and -not [string]::IsNullOrWhiteSpace($result.ApiKey)) {
            Write-SqlLog -Level 'SUCCESS' -Message "Successfully retrieved and decrypted API Key."
            $script:GeminiApiKey = $result.ApiKey
            return $true
        }
        else {
            Write-SqlLog -Level 'ERROR' -Message "Fatal: API Key query returned null or empty. Check '$ConfigTableName', certificate, and master key password."
            return $false
        }
    }
    catch {
        Write-SqlLog -Level 'ERROR' -Message "Fatal: Failed to execute API Key retrieval query. Error: $($_.Exception.Message)"
        return $false
    }
}

function Get-ObjectsFromPlan {
    param([xml]$ExecutionPlanXml)
    Write-SqlLog -Level 'INFO' -Message "Parsing execution plan to identify database objects."
    try {
        $ns = New-Object System.Xml.XmlNamespaceManager($ExecutionPlanXml.NameTable)
        $ns.AddNamespace("sql", "http://schemas.microsoft.com/sqlserver/2004/07/showplan")
        
        $objectNodes = $ExecutionPlanXml.SelectNodes("//sql:Object", $ns)
        $uniqueObjectNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        foreach ($node in $objectNodes) {
            $db = $node.GetAttribute("Database")
            $schema = $node.GetAttribute("Schema")
            $table = $node.GetAttribute("Table")
            
            if ($table -notlike "#*" -and -not ([string]::IsNullOrWhiteSpace($db)) -and -not ([string]::IsNullOrWhiteSpace($schema)) -and $db -ne '[mssqlsystemresource]') {
                $fullName = "$db.$schema.$table".Replace('[', '').Replace(']', '')
                $uniqueObjectNames.Add($fullName) | Out-Null
            }
        }

        $finalList = $uniqueObjectNames | Sort-Object
        Write-SqlLog -Level 'INFO' -Message "Identified $($finalList.Count) unique user objects in the plan."
        return $finalList
    }
    catch {
        Write-SqlLog -Level 'ERROR' -Message "Failed to parse objects from execution plan XML. Error: $($_.Exception.Message)"
        return @()
    }
}

function Get-ObjectSchema {
    param(
        [string]$TargetServer,
        [string]$DatabaseName,
        [string]$SchemaName,
        [string]$ObjectName
    )
    
    $twoPartObjectName = "[$SchemaName].[$ObjectName]"
    $sanitizedTwoPartName = $twoPartObjectName.Replace("'", "''")
    $schemaText = "--- Schema For: [$DatabaseName].$twoPartObjectName ---`n"

    try {
        $colQuery = "SELECT c.name, t.name AS system_type_name, c.max_length, c.precision, c.scale, c.is_nullable FROM sys.columns c JOIN sys.types t ON c.user_type_id = t.user_type_id WHERE c.object_id = OBJECT_ID('$sanitizedTwoPartName') ORDER BY c.column_id;"
        $columnResult = Invoke-Sqlcmd -ServerInstance $TargetServer -Database $DatabaseName -TrustServerCertificate -Query $colQuery -ErrorAction Stop
        
        if ($columnResult) {
            $schemaText += "COLUMNS:`n"
            foreach ($col in $columnResult) {
                $isNullable = if ($col.is_nullable) { 'NULL' } else { 'NOT NULL' }
                $schemaText += "- $($col.name) $($col.system_type_name) $isNullable`n"
            }
        }
    } catch { Write-SqlLog -Level 'WARN' -Message "Could not get COLUMN schema for '[$DatabaseName].$twoPartObjectName'. Error: $($_.Exception.Message)" }

    try {
        $idxQuery = @"
SELECT i.name AS IndexName, i.type_desc AS IndexType, i.is_unique,
STUFF((SELECT ', ' + c.name FROM sys.index_columns ic JOIN sys.columns c ON ic.object_id = c.object_id AND ic.column_id = c.column_id WHERE ic.object_id = i.object_id AND ic.index_id = i.index_id AND ic.is_included_column = 0 ORDER BY ic.key_ordinal FOR XML PATH('')), 1, 2, '') AS KeyColumns,
STUFF((SELECT ', ' + c.name FROM sys.index_columns ic JOIN sys.columns c ON ic.object_id = c.object_id AND ic.column_id = c.column_id WHERE ic.object_id = i.object_id AND ic.index_id = i.index_id AND ic.is_included_column = 1 ORDER BY ic.key_ordinal FOR XML PATH('')), 1, 2, '') AS IncludedColumns
FROM sys.indexes i WHERE i.object_id = OBJECT_ID('$sanitizedTwoPartName');
"@
        $indexResult = Invoke-Sqlcmd -ServerInstance $TargetServer -Database $DatabaseName -TrustServerCertificate -Query $idxQuery -ErrorAction Stop
        if ($indexResult) {
            $schemaText += "`nINDEXES:`n"
            foreach ($idx in $indexResult) {
                $uniqueText = if ($idx.is_unique) { "UNIQUE " } else { "" }
                $idxLine = "- $($idx.IndexName) ($($idx.IndexType) / $($uniqueText)KEYS: $($idx.KeyColumns))"
                if (-not [string]::IsNullOrWhiteSpace($idx.IncludedColumns)) { $idxLine += " (INCLUDES: $($idx.IncludedColumns))" }
                $schemaText += $idxLine + "`n"
            }
        }
    } catch { Write-SqlLog -Level 'WARN' -Message "Could not get INDEX info for '[$DatabaseName].$twoPartObjectName'. Error: $($_.Exception.Message)" }

    try {
        $conQuery = @"
SELECT name, type_desc FROM sys.key_constraints WHERE parent_object_id = OBJECT_ID('$sanitizedTwoPartName');
SELECT name, type_desc FROM sys.check_constraints WHERE parent_object_id = OBJECT_ID('$sanitizedTwoPartName');
SELECT name, type_desc FROM sys.default_constraints WHERE parent_object_id = OBJECT_ID('$sanitizedTwoPartName');
"@
        $constraintResult = Invoke-Sqlcmd -ServerInstance $TargetServer -Database $DatabaseName -TrustServerCertificate -Query $conQuery -ErrorAction Stop
        if ($constraintResult) {
            $schemaText += "`nCONSTRAINTS:`n"
            foreach($con in $constraintResult) {
                $schemaText += "- $($con.name) ($($con.type_desc))`n"
            }
        }
    } catch { Write-SqlLog -Level 'WARN' -Message "Could not get CONSTRAINT info for '[$DatabaseName].$twoPartObjectName'. Error: $($_.Exception.Message)" }

    return $schemaText + "`n"
}

function Invoke-GeminiAnalysis {
    param(
        [string]$ExecutionPlan,
        [string]$SchemaDocument,
        [string]$SqlVersion
    )
    Write-SqlLog -Level 'INFO' -Message "Sending context to Gemini AI for analysis."

    $uri = "https://generativelanguage.googleapis.com/v1beta/models/$($GeminiModel):generateContent?key=$($script:GeminiApiKey)"

    $prompt = @"
### 1. Overview
You are an expert T-SQL performance tuning assistant for Microsoft SQL Server. Your task is to analyze the provided execution plan and object schemas to identify performance bottlenecks and provide actionable tuning recommendations. 

### 2. Rules for Analysis
* Analyze the execution plan for costly operators such as Table Scans, Clustered Index Scans, Key Lookups, and large Sort operations.
* Examine the provided object schemas (columns, data types, indexes) to formulate precise `CREATE INDEX` or `ALTER INDEX` recommendations. Ensure your recommendations are covering indexes where possible to avoid lookups.
* Identify non-SARGable predicates in the query (e.g., functions applied to columns in a `WHERE` clause) and suggest alternative, SARGable T-SQL syntax.

### 3. Rules for Output
* Your ENTIRE response MUST be valid T-SQL syntax, not exceeding 8000 characters. Do not include any markdown, conversational text, or explanations outside of T-SQL comments.
* Your output must be a single T-SQL block comment (`/* ... */`).
* Inside the block comment, format your findings as a numbered list. For each distinct issue you identify, you MUST use the following three-part structure:

1.  [A short, high-level description of the issue, e.g., 'Costly Index Scan on Table X']
    a. Problem: A clear and concise explanation of the specific performance issue.
    b. Recommendation: Provide the recommended T-SQL code (e.g., a complete CREATE INDEX statement or a rewritten query snippet). This code should be ready to be copied and executed.
    c. Expected Result: Briefly describe the expected performance gain (e.g., 'This will change the operation from an Index Scan to an Index Seek, significantly reducing logical reads.').

* If you have multiple recommendations, create a new numbered entry for each one inside the same single T-SQL block comment.


---
### 4. SQL SERVER VERSION
You must ensure all generated T-SQL syntax is compatible with this version:
$SqlVersion

---
### 5. EXECUTION PLAN XML
$ExecutionPlan

---
### 6. OBJECT SCHEMAS
$SchemaDocument
"@

    $bodyObject = @{ contents = @( @{ parts = @( @{ text = $prompt } ) } ) }
    $bodyJson = $bodyObject | ConvertTo-Json -Depth 10

    try {
        $response = Invoke-RestMethod -Uri $uri -Method Post -Body $bodyJson -ContentType 'application/json' -ErrorAction Stop -TimeoutSec 600
        $rawAiResponse = $response.candidates[0].content.parts[0].text
        Write-SqlLog -Level 'SUCCESS' -Message "Successfully received analysis from Gemini AI."
        return $rawAiResponse.Trim()
    }
    catch {
        $errorMessage = "Failed to get response from Gemini API. Error: $($_.Exception.Message)"
        if ($_.Exception.Response) {
            $errorDetails = $_.Exception.Response.GetResponseStream()
            $streamReader = New-Object System.IO.StreamReader($errorDetails)
            $errorText = $streamReader.ReadToEnd()
            $errorMessage += " | API Error Details: $errorText"
        }
        Write-SqlLog -Level 'ERROR' -Message $errorMessage
        return $null
    }
}

# --- Main Application Logic ---
function Start-OptimusQS {
    Write-SqlLog -Level 'INFO' -Message "--- OptimusQS Analysis Run Started ---"

    if (-not (Test-SqlServerModule)) {
        Write-SqlLog -Level 'ERROR' -Message "Halting script due to missing module dependency."
        return
    }

    if (-not (Get-ApiKeyFromDatabase)) {
        Write-SqlLog -Level 'ERROR' -Message "Halting script due to API Key retrieval failure."
        return
    }

    $sqlVersion = (Invoke-Sqlcmd -ServerInstance $ServerName -Query "SELECT @@VERSION" -TrustServerCertificate).Item(0)
    Write-SqlLog -Level 'INFO' -Message "Connected to SQL Server Version: $sqlVersion"
    $dbNameForData = ($FullObjectName -split '\.')[0].Trim('[]')

    for ($retryCount = 1; $retryCount -le ($MaxRetries + 1); $retryCount++) {
        $dataQuery = "SELECT ID, execution_plan_xml FROM $FullObjectName WHERE tuning_response IS NULL AND execution_plan_xml IS NOT NULL;"
        
        try {
            $queriesToProcess = Invoke-Sqlcmd -ServerInstance $ServerName `
                                              -Database $dbNameForData `
                                              -Query $dataQuery `
                                              -TrustServerCertificate `
                                              -MaxCharLength ([int]::MaxValue) `
                                              -ErrorAction Stop
        }
        catch {
            Write-SqlLog -Level 'ERROR' -Message "Fatal: Could not retrieve data from '$FullObjectName'. Error: $($_.Exception.Message)"
            return
        }

        if ($null -eq $queriesToProcess -or $queriesToProcess.Count -eq 0) {
            Write-SqlLog -Level 'SUCCESS' -Message "No queries remaining for analysis. All tasks complete."
            break
        }

        Write-SqlLog -Level 'INFO' -Message "--- Starting Analysis Pass #$($retryCount). Found $($queriesToProcess.Count) queries to process. ---"

        foreach ($query in $queriesToProcess) {
            $currentId = $query.ID
            Write-SqlLog -Level 'INFO' -Message "--- Analyzing ID: $currentId ---"
            
            try {
                [xml]$planXml = $query.execution_plan_xml
            }
            catch {
                Write-SqlLog -Level 'ERROR' -Message "ID: $currentId - Skipping due to invalid XML in 'execution_plan_xml'. Error: $($_.Exception.Message)"
                continue
            }

            $objectList = Get-ObjectsFromPlan -ExecutionPlanXml $planXml
            if ($objectList.Count -eq 0) {
                Write-SqlLog -Level 'WARN' -Message "ID: $currentId - Skipping as no user objects were found in the execution plan."
                continue
            }

            $consolidatedSchema = ""
            Write-SqlLog -Level 'INFO' -Message "ID: $currentId - Collecting schema for $($objectList.Count) objects."
            foreach ($obj in $objectList) {
                $parts = $obj.Split('.')
                $consolidatedSchema += Get-ObjectSchema -TargetServer $ServerName -DatabaseName $parts[0] -SchemaName $parts[1] -ObjectName $parts[2]
            }
            
            if ([string]::IsNullOrWhiteSpace($consolidatedSchema)) {
                Write-SqlLog -Level 'WARN' -Message "ID: $currentId - Skipping as schema collection resulted in an empty document. Check permissions."
                continue
            }

            $tuningResponse = Invoke-GeminiAnalysis -ExecutionPlan $planXml.OuterXml -SchemaDocument $consolidatedSchema -SqlVersion $sqlVersion
            
            if ($null -ne $tuningResponse -and $tuningResponse -match 'model is overloaded|try again later') {
                Write-SqlLog -Level 'WARN' -Message "ID: $currentId - Gemini reported it is overloaded. Will retry on the next pass."
                $tuningResponse = $null
            }

            if (-not [string]::IsNullOrWhiteSpace($tuningResponse)) {
                # **FIX:** Increased truncation length and updated warning message.
                if ($tuningResponse.Length -gt 8000) {
                    $tuningResponse = $tuningResponse.Substring(0, 8000)
                    Write-SqlLog -Level 'WARN' -Message "ID: $currentId - AI response was truncated to 8000 characters."
                }

                $sanitizedResponse = $tuningResponse.Replace("'", "''")
                $updateQuery = "UPDATE $FullObjectName SET tuning_response = '$sanitizedResponse' WHERE ID = $currentId;"

                try {
                    Invoke-Sqlcmd -ServerInstance $ServerName -Database $dbNameForData -Query $updateQuery -TrustServerCertificate -ErrorAction Stop
                    Write-SqlLog -Level 'SUCCESS' -Message "ID: $currentId - Successfully updated with tuning recommendations."
                }
                catch {
                    Write-SqlLog -Level 'ERROR' -Message "ID: $currentId - Failed to write tuning response to database. Error: $($_.Exception.Message)"
                }
            }
            else {
                Write-SqlLog -Level 'ERROR' -Message "ID: $currentId - Skipping update as AI analysis returned no response or a retryable error."
            }
        } 

        if ($retryCount -le $MaxRetries -and (Invoke-Sqlcmd -ServerInstance $ServerName -Database $dbNameForData -Query "SELECT COUNT(*) FROM $FullObjectName WHERE tuning_response IS NULL AND execution_plan_xml IS NOT NULL;").Item(0) -gt 0) {
            $delaySeconds = 60
            Write-SqlLog -Level 'INFO' -Message "Analysis Pass #$($retryCount) complete. Waiting for $delaySeconds seconds before next attempt."
            Start-Sleep -Seconds $delaySeconds
        }

    }

    Write-SqlLog -Level 'INFO' -Message "--- OptimusQS Analysis Run Finished ---"
}

# --- Script Entry Point ---
Start-OptimusQS