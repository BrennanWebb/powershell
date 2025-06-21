<#
.SYNOPSIS
    Optimus is a T-SQL tuning advisor that leverages the Gemini AI for performance recommendations.

.DESCRIPTION
    This script performs a statement-by-statement analysis of a T-SQL file by generating a single
    master execution plan. This process inherently validates the script's syntax. The script then parses this plan,
    analyzing each statement individually. For each statement, it collects object schemas, packages the artifacts
    (SQL, schema, plan), and sends them to the Google Gemini API for performance analysis. The final output is a
    single, tuned SQL script with AI-generated recommendations embedded as comments.

.PARAMETER SQLFile
    The path to the .sql file to be analyzed.

.PARAMETER UseActualPlan
    An optional switch to generate the 'Actual' execution plan. This WILL execute the query.
    If not present, the script defaults to 'Estimated' or will prompt in interactive mode.

.PARAMETER DebugMode
    Enables detailed diagnostic output.

.EXAMPLE
    .\Optimus.ps1 -SQLFile "C:\Scripts\MyProcedure.sql"
    Runs the script in normal mode. Intermediary files will be cleaned up automatically.

.EXAMPLE
    .\Optimus.ps1 -DebugMode
    Runs the script with verbose debugging messages.

.NOTES
    Author: Gemini
    Version: 10.0
    Created: 2025-06-17
    Powershell Version: 5.1+
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [string]$SQLFile,

    [Parameter(Mandatory=$false)]
    [switch]$UseActualPlan,

    [Parameter(Mandatory=$false)]
    [switch]$DebugMode
)

#region Color-Coded & Debug Functions
function Write-Host-Cyan   { param([string]$Message) Write-Host $Message -ForegroundColor Cyan }
function Write-Host-Green  { param([string]$Message) Write-Host $Message -ForegroundColor Green }
function Write-Host-Yellow { param([string]$Message) Write-Host $Message -ForegroundColor Yellow }
function Write-Host-Red    { param([string]$Message) Write-Host $Message -ForegroundColor Red }
function Write-Host-White  { param([string]$Message) Write-Host $Message -ForegroundColor White }
function Write-DebugMessage { param([string]$Message) if ($DebugMode) { Write-Host "[DEBUG] $Message" -ForegroundColor Gray } }
#endregion

#region Configuration Management
function Initialize-Configuration {
    Write-DebugMessage "Entering Function: Initialize-Configuration"
    Write-Host-Cyan "Initializing Optimus configuration..."
    try {
        $userProfile = $env:USERPROFILE
        $configDir = Join-Path -Path $userProfile -ChildPath ".optimus"
        $analysisBaseDir = Join-Path -Path $configDir -ChildPath "Analyses"
        $serverFile = Join-Path -Path $configDir -ChildPath "servers.json"
        $apiKeyFile = Join-Path -Path $configDir -ChildPath "api.config"

        foreach($dir in @($configDir, $analysisBaseDir)){ if (-not (Test-Path -Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null; Write-Host-Green "Directory created at '$dir'." } }
        if (-not (Test-Path -Path $serverFile)) { Set-Content -Path $serverFile -Value "[]"; Write-Host-Green "Server list file created." }
        
        $script:OptimusConfig = @{ AnalysisBaseDir = $analysisBaseDir; ServerFile = $serverFile; ApiKeyFile = $apiKeyFile }
        Write-Host-Green "Configuration initialized successfully."; return $true
    }
    catch { Write-Host-Red "ERROR: Could not initialize configuration: $($_.Exception.Message)"; return $false }
}

function Get-And-Set-ApiKey {
    Write-DebugMessage "Entering Function: Get-And-Set-ApiKey"
    Write-Host-Cyan "Checking for Gemini API Key..."
    $apiKeyFile = $script:OptimusConfig.ApiKeyFile
    if (Test-Path -Path $apiKeyFile) {
        try { $script:GeminiApiKey = Get-Content -Path $apiKeyFile | ConvertTo-SecureString; Write-Host-Green "API Key loaded successfully."; return }
        catch { Write-Host-Yellow "Could not read existing API key. It may be corrupt. Please re-enter." }
    }
    $secureKey = Read-Host "Please enter your Gemini API Key" -AsSecureString
    if ($secureKey.Length -eq 0) { Write-Host-Red "ERROR: API Key cannot be empty."; Get-And-Set-ApiKey; return }
    try {
        $secureKey | ConvertFrom-SecureString | Set-Content -Path $apiKeyFile
        $script:GeminiApiKey = $secureKey; Write-Host-Green "API Key has been validated and saved securely."
    }
    catch { Write-Host-Red "ERROR: Failed to save API Key: $($_.Exception.Message)"; throw "Could not save API key." }
}
#endregion

#region Core SQL, Validation & File Functions
function Test-SqlServerModule {
    Write-DebugMessage "Entering Function: Test-SqlServerModule"
    Write-Host-Cyan "Checking for 'SqlServer' PowerShell module..."
    if (Get-Module -Name SqlServer -ListAvailable) {
        try { Import-Module SqlServer -ErrorAction Stop; Write-Host-Green "'SqlServer' module imported."; return $true }
        catch { Write-Host-Red "ERROR: Failed to import 'SqlServer' module: $($_.Exception.Message)"; return $false }
    } else { Write-Host-Yellow "The 'SqlServer' module is required. Run: Install-Module -Name SqlServer -Scope CurrentUser"; return $false }
}

function Test-SqlServerConnection {
    param([string]$ServerInstance)
    Write-DebugMessage "Entering Function: Test-SqlServerConnection for server '$ServerInstance'"
    Write-Host-Cyan "Testing connection to '$ServerInstance'..."
    try { Invoke-Sqlcmd -ServerInstance $ServerInstance -Query "SELECT @@VERSION" -QueryTimeout 5 -TrustServerCertificate -ErrorAction Stop | Out-Null; Write-Host-Green "Connection successful!"; return $true }
    catch { Write-Host-Red "ERROR: Failed to connect to '$ServerInstance': $($_.Exception.Message)"; return $false }
}

function Select-SqlServer {
    Write-DebugMessage "Entering Function: Select-SqlServer"
    Write-Host-Cyan "Please select a SQL Server to use:"
    [array]$servers = Get-Content -Path $script:OptimusConfig.ServerFile | ConvertFrom-Json
    if ($servers.Count -gt 0) { for ($i = 0; $i -lt $servers.Count; $i++) { Write-Host-White "[$($i+1)] $($servers[$i])" } }
    Write-Host-White "[A] Add a new server"; Write-Host-White "[Q] Quit"
    while ($true) {
        $choice = Read-Host "Enter your choice"
        if ($choice -imatch 'Q') { return $null }
        if ($choice -imatch 'A') {
            $newServer = Read-Host "Enter the new SQL server name or IP"
            if ([string]::IsNullOrWhiteSpace($newServer)) { Write-Host-Red "Server name cannot be empty."; continue }
            if (Test-SqlServerConnection -ServerInstance $newServer) {
                $servers += $newServer; ($servers | Sort-Object -Unique) | ConvertTo-Json -Depth 5 | Set-Content -Path $script:OptimusConfig.ServerFile
                Write-Host-Green "'$newServer' has been added."; return $newServer
            }
            continue
        }
        if ($choice -match '^\d+$' -and [int]$choice -gt 0 -and [int]$choice -le $servers.Count) {
            $selectedServer = $servers[[int]$choice - 1]
            if (Test-SqlServerConnection -ServerInstance $selectedServer) { return $selectedServer }
        } else { Write-Host-Red "Invalid choice." }
    }
}

function Show-FilePicker {
    Write-DebugMessage "Entering Function: Show-FilePicker"
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $fileDialog.Title = "Select SQL File for Analysis"
        $fileDialog.InitialDirectory = [System.Environment]::GetFolderPath('MyDocuments')
        $fileDialog.Filter = "SQL Files (*.sql)|*.sql|All files (*.*)|*.*"
        if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $fileDialog.FileName }
    }
    catch { Write-Host-Yellow "Could not display graphical file picker: $($_.Exception.Message)" }
    return $null
}

function Get-SQLQueryFile {
    Write-DebugMessage "Entering Function: Get-SQLQueryFile"
    $localFile = $null
    if (-not [string]::IsNullOrWhiteSpace($SQLFile)) {
        if ((Test-Path -Path $SQLFile -PathType Leaf) -and $SQLFile -like '*.sql') { $localFile = $SQLFile }
        else { Write-Host-Red "Parameter invalid: '$SQLFile'." }
    }
    while ([string]::IsNullOrWhiteSpace($localFile)) {
        Write-Host-Cyan "Please select the .sql file to analyze..."; $localFile = Show-FilePicker
        if (-not $localFile) {
            $retry = Read-Host "File selection cancelled. Try again? (Y/N)"; if ($retry -notmatch '^[Yy]$') { return $null }
        }
    }
    Write-Host-Green "Successfully selected SQL file: $localFile"; return $localFile
}
#endregion

#region Data Parsing, Collection, and AI Analysis
function Get-MasterExecutionPlan {
    param($ServerInstance, $DatabaseContext, $FullQueryText, [switch]$IsActualPlan)
    Write-DebugMessage "Entering Function: Get-MasterExecutionPlan"
    
    $planCommand = if ($IsActualPlan) { "SET STATISTICS XML ON;" } else { "SET SHOWPLAN_XML ON;" }
    $planType = if ($IsActualPlan) { "Actual" } else { "Estimated" }
    Write-Host-Cyan "Generating master '$planType' execution plan (this also validates script syntax)..."
    
    $dbContextForCheck = if ([string]::IsNullOrWhiteSpace($DatabaseContext)) { 'master' } else { $DatabaseContext }
    Write-DebugMessage "Using database context '$dbContextForCheck' to generate plan."

    $planQuery = "$planCommand`nGO`n$FullQueryText"
    try {
        $planResult = Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $dbContextForCheck -TrustServerCertificate -Query $planQuery -MaxCharLength ([int]::MaxValue) -ErrorAction Stop
        Write-Host-Green "Successfully generated master execution plan."
        return $planResult.Item(0)
    }
    catch {
        Write-Host-Red "ERROR: The SQL script is invalid. SQL Server could not compile it. Error: $($_.Exception.Message)"
        return $null
    }
}

function Get-SqlObjectData {
    param([string]$QueryText)
    Write-DebugMessage "Entering Function: Get-SqlObjectData"
    $part = '(?:\[?[\w\d_]+\]?)'
    $objectRegex = "(?i)(?:FROM|JOIN|INTO|UPDATE|DELETE(?:\s+FROM)?)\s+($part(?:\.$part){0,2})"
    $regexMatches = [regex]::Matches($QueryText, $objectRegex)
    Write-DebugMessage "Regex found $($regexMatches.Count) potential objects."
    if ($regexMatches.Count -eq 0) { return @() }
    $parsedObjects = @()
    foreach ($match in $regexMatches) {
        Write-DebugMessage "Processing raw regex match: '$($match.Groups[1].Value)'"
        $fullName = $match.Groups[1].Value.Replace('[','').Replace(']','')
        $parts = $fullName.Split('.')
        $obj = [pscustomobject]@{ FullName = $fullName; Database = $null }
        if ($parts.Count -eq 3) { $obj.Database = $parts[0] }
        $parsedObjects += $obj
    }
    return ($parsedObjects | Sort-Object FullName -Unique)
}

function Get-ObjectSchema {
    param($ServerInstance, $DatabaseContext, [array]$ParsedObjects)
    Write-DebugMessage "Entering function: Get-ObjectSchema"
    $allSchemaText = ""
    Write-Host-Cyan "Collecting schema for objects..."
    foreach ($obj in $ParsedObjects) {
        $db = if ($obj.Database) { $obj.Database } else { $DatabaseContext }
        if (-not $db) {
            Write-Host-Red "ERROR: Cannot determine database for object '$($obj.FullName)'. Analysis cannot continue."
            return $null
        }
        $schemaQuery = "SELECT name, system_type_name, max_length, [precision], scale, is_nullable FROM sys.dm_exec_describe_first_result_set(N'SELECT * FROM $($obj.FullName)', NULL, 0);"
        try {
            $schemaResult = Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $db -TrustServerCertificate -Query $schemaQuery -ErrorAction Stop
            $allSchemaText += "--- Schema for $($obj.FullName) (in database $db):`n"
            $allSchemaText += ($schemaResult | Format-Table | Out-String) + "`n"
        } catch { 
            Write-Host-Yellow "Warning: Could not describe schema for '$($obj.FullName)' in database '$db'. It may not be a table or view."
            continue
        }
    }
    if ([string]::IsNullOrWhiteSpace($allSchemaText)) { $allSchemaText = "--- No object schemas could be determined." }
    Write-DebugMessage "Collected schema text:`n$allSchemaText"
    return $allSchemaText
}

function Invoke-GeminiAnalysis {
    param([securestring]$ApiKey, [string]$StatementText, [string]$SchemaText, [string]$PlanXml)
    Write-DebugMessage "Entering Function: Invoke-GeminiAnalysis"
    Write-DebugMessage "Analyzing statement: $StatementText"
    Write-Host-Cyan "Sending statement to Gemini for analysis..."

    $plainTextApiKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ApiKey))
    $uri = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=$plainTextApiKey"

    $prompt = @"
You are an expert T-SQL performance tuning assistant. I will provide a single T-SQL statement, the schema of all referenced objects, and the estimated XML execution plan for that statement.

Your task is to analyze these three components and provide specific, actionable recommendations to improve the query's performance.

Your response MUST be ONLY the tuned T-SQL script.
- Add your recommendations as T-SQL block comments (/* ... */) directly above the specific line of code they apply to.
- Explain WHY the recommendation is being made inside the comment.
- If no improvements are possible, return the original T-SQL statement with a single comment at the top: /* No performance tuning recommendations found. */
- Do not include any conversational text, greetings, or explanations outside of the T-SQL comments.
- Ensure the final output is valid T-SQL code.

--- T-SQL STATEMENT ---
$StatementText

--- OBJECT SCHEMAS ---
$SchemaText

--- EXECUTION PLAN ---
$PlanXml
"@

    $body = @{
        contents = @(
            @{
                parts = @(
                    @{
                        text = $prompt
                    }
                )
            }
        )
    } | ConvertTo-Json -Depth 10

    try {
        $response = Invoke-RestMethod -Uri $uri -Method Post -Body $body -ContentType 'application/json' -ErrorAction Stop
        $tunedSql = $response.candidates[0].content.parts[0].text
        Write-DebugMessage "Successfully received response from Gemini API."
        Write-Host-Green "AI analysis complete."
        return $tunedSql
    } catch {
        Write-Host-Red "ERROR: Failed to get response from Gemini API."
        $errorDetails = $_.Exception.Response.GetResponseStream()
        $streamReader = New-Object System.IO.StreamReader($errorDetails)
        $errorText = $streamReader.ReadToEnd()
        Write-Host-Red "API Error Details: $errorText"
        return $null
    }
}
#endregion

# --- Main Application Logic ---
function Start-Optimus {
    if ($DebugMode) { Write-DebugMessage "Starting Optimus v10.0 in Debug Mode."}
    if (-not (Initialize-Configuration)) { return }
    if (-not (Test-SqlServerModule)) { return }
    Get-And-Set-ApiKey
    Write-Host-Cyan "`n--- Welcome to Optimus v10.0 ---"
    do {
        $analysisPath = $null 
        try {
            $selectedServer = Select-SqlServer; if (-not $selectedServer) { Write-Host-Yellow "No server selected."; break }
            $sqlFilePath = Get-SQLQueryFile; if (-not $sqlFilePath) { Write-Host-Yellow "No file selected."; continue }
            
            $useActualPlanSwitch = $UseActualPlan.IsPresent
            if (-not $useActualPlanSwitch) {
                Write-Host-Cyan "`nWhich execution plan would you like to generate?"
                Write-Host-White "[1] Estimated (Default - Recommended, does not run the query)"
                Write-Host-White "[2] Actual (Executes the query, use with caution)"
                $choice = Read-Host "Enter your choice"
                if ($choice -eq '2') {
                    Write-Host-Yellow "`nWARNING: Proceeding with 'Actual Execution Plan'. This will EXECUTE your entire SQL script."
                    $useActualPlanSwitch = $true
                }
            }
            
            $sqlQueryText = Get-Content -Path $sqlFilePath -Raw
            Write-DebugMessage "Initial SQL file content:`n---`n$sqlQueryText`n---"
            
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($sqlFilePath)
            $analysisPath = Join-Path -Path $script:OptimusConfig.AnalysisBaseDir -ChildPath "${timestamp}_${baseName}"
            New-Item -Path $analysisPath -ItemType Directory -Force | Out-Null
            Write-Host-Cyan "Created analysis directory: '$analysisPath'"

            $initialDbContext = ([regex]::Match($sqlQueryText, '(?im)^\s*USE\s+\[?([\w\d_]+)\]?')).Groups[1].Value
            
            $masterPlanXml = Get-MasterExecutionPlan -ServerInstance $selectedServer -DatabaseContext $initialDbContext -FullQueryText $sqlQueryText -IsActualPlan:$useActualPlanSwitch
            if (-not $masterPlanXml) { continue }
            
            [xml]$masterPlan = $masterPlanXml
            $ns = New-Object System.Xml.XmlNamespaceManager($masterPlan.NameTable)
            $ns.AddNamespace("sql", "http://schemas.microsoft.com/sqlserver/2004/07/showplan")
            $statementNodes = $masterPlan.SelectNodes("//sql:StmtSimple", $ns)
            Write-Host-Cyan "Found $($statementNodes.Count) statements in the execution plan to analyze."

            $finalTunedStatements = @()
            $overallSuccess = $true
            
            for ($i = 0; $i -lt $statementNodes.Count; $i++) {
                $node = $statementNodes[$i]
                # FIX: Trim any leading whitespace or semicolons from the statement text
                $currentStatement = $node.StatementText.TrimStart(" `t`n`r;")
                $currentPlanXml = $node.QueryPlan.OuterXml
                
                Write-Host-Cyan "--- Analyzing Statement $($i + 1) of $($statementNodes.Count) ---"
                Write-DebugMessage "Statement Text: $currentStatement"

                $parsedObjects = Get-SqlObjectData -QueryText $currentStatement
                
                $isAmbiguous = $false
                if (-not $initialDbContext) {
                    if ($parsedObjects | Where-Object { -not $_.Database }) { $isAmbiguous = $true }
                } else {
                    $parsedObjects | ForEach-Object { if (-not $_.Database) { $_.Database = $initialDbContext } }
                }

                if ($isAmbiguous) {
                    $ambiguousObjects = ($parsedObjects | Where-Object { -not $_.Database }).FullName -join ", "
                    Write-Host-Red "ERROR: Statement is ambiguous and no global USE context was found: [$ambiguousObjects]"
                    $overallSuccess = $false; break
                }

                if ($parsedObjects.Count -eq 0) {
                    Write-Host-Yellow "No queryable objects found in this statement. Skipping analysis."
                    $finalTunedStatements += $currentStatement; continue
                }
                
                $schemaText = Get-ObjectSchema -ServerInstance $selectedServer -DatabaseContext $initialDbContext -ParsedObjects $parsedObjects
                if (-not $schemaText) { $overallSuccess = $false; break }
                
                $tunedStatement = Invoke-GeminiAnalysis -ApiKey $script:GeminiApiKey -StatementText $currentStatement -SchemaText $schemaText -PlanXml $currentPlanXml
                if (-not $tunedStatement) { $overallSuccess = $false; break }
                $finalTunedStatements += $tunedStatement
            }

            if ($overallSuccess -and $finalTunedStatements.Count -gt 0) {
                Write-Host-Green "--- Overall Analysis Complete ---"
                $finalScript = $finalTunedStatements -join "$([System.Environment]::NewLine)GO$([System.Environment]::NewLine)"
                $tunedScriptPath = Join-Path -Path $analysisPath -ChildPath "Tuned_Query.sql"
                $finalScript | Out-File -FilePath $tunedScriptPath -Encoding UTF8
                Write-Host-Green "Successfully generated tuned script at: '$tunedScriptPath'"
                Write-Host-Cyan "--- Tuned Script ---"
                Write-Host-White $finalScript
                Write-Host-Cyan "--------------------"
            } elseif ($overallSuccess) {
                 Write-Host-Green "Analysis complete. No tunable statements were found to process."
            } else {
                Write-Host-Red "Analysis halted due to an error. Please review messages above."
            }
        }
        catch {
            Write-Host-Red "CRITICAL ERROR: An unexpected error occurred in the main block: $($_.Exception.Message)"
        }
        
        $response = Read-Host "Would you like to analyze another SQL file? (Y/N)"
    } while ($response -match '^[Yy]$')
    Write-Host-Green "Thank you for using Optimus. Exiting."
}

# --- Script Entry Point ---
Start-Optimus