<#
.SYNOPSIS
    Optimus is a T-SQL tuning advisor with a Graphical User Interface that leverages the Gemini AI for performance recommendations.

.DESCRIPTION
    This script provides a WPF user interface to perform a holistic analysis of T-SQL files. It can process a single file, 
    a list of files, or all .sql files in a specified folder. It generates a master execution plan to validate syntax and 
    identify database objects. It then queries each required database to build a comprehensive schema document for all 
    user-defined objects. All execution steps, messages, and recommendations are recorded in a detailed log file for each analysis.

.PARAMETER DebugMode
    Enables detailed diagnostic output to the console and the UI log. All messages are always written to the execution log file regardless of this setting.

.NOTES
    Designer: Brennan Webb
    Script Engine: Gemini
    Version: 3.0
    Created: 2025-06-21
    Modified: 2025-06-22
    Change Log:
    - v3.0: Complete overhaul to implement a WPF graphical user interface for an improved user experience.
    - v3.0: Refactored core logic to decouple it from the UI for better maintainability.
    - v3.0: Enhanced logging to support real-time updates to the UI's output window.
    - v2.0: Updated output directory to be segmented by model name.
    - v2.0: Promoted to major version after preview cycle.
    - v1.9-preview: Updated AI prompt to explicitly forbid Markdown within recommendation comments.
    - v1.8-preview: Implemented a one-time, persistent model selection configuration.
    - v1.8-preview: Updated reset logic to include clearing the selected model.
    - v1.7-preview: Corrected and hardened the AI prompt to ensure the full original script is always returned.
    - v1.6-preview: Added logic to skip schema collection for the mssqlsystemresource database.
    - v1.5-preview: Modified plan collection to capture all statements in a script.
    - v1.5-preview: Hardened schema collection to support system tables.
    - v1.4-preview: Removed the logic that excludes 'sys' schema objects from analysis.
    - v1.3-preview: Added an option to Reset-OptimusConfiguration to remove only analysis reports.
    - v1.2-preview: Renamed function to use an approved PowerShell verb (Invoke-OptimusVersionCheck).
    - v1.1-preview: Added an automatic version check at startup.
    - v1.0-preview: Initial preview release. Re-versioned from legacy builds.
    Powershell Version: 5.1+
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [switch]$DebugMode
)

# --- SCRIPT-WIDE CONFIGURATION ---
$script:CurrentVersion = "3.0"
$script:uiMode = $true # Flag to indicate we are running in UI mode.

#region Centralized Logging
function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [Parameter(Mandatory=$false)]
        [ValidateSet('DEBUG', 'INFO', 'SUCCESS', 'WARN', 'ERROR', 'PROMPT', 'RESULT')]
        [string]$Level = 'INFO',
        [Parameter(Mandatory=$false)]
        [switch]$NoNewLine
    )

    # 1. Always write the full message to the log file first if the path is set.
    if ($script:LogFilePath) {
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logMessageToFile = "$timestamp [$Level] - $Message"
        try {
            Add-Content -Path $script:LogFilePath -Value $logMessageToFile -Encoding UTF8
        } catch {
            Write-Warning "CRITICAL: Failed to write to log file $($script:LogFilePath): $($_.Exception.Message)"
        }
    }

    # 2. Handle all console/UI output.
    # For DEBUG level, only write to console if DebugMode switch is present.
    if ($Level -eq 'DEBUG' -and -not $DebugMode) {
        return
    }

    $consoleMessage = $Message
    $color = 'White'
    switch ($Level) {
        'DEBUG'   { $color = 'Gray';   $consoleMessage = "[DEBUG] $Message" }
        'INFO'    { $color = 'Cyan';   }
        'SUCCESS' { $color = 'Green';  }
        'WARN'    { $color = 'Yellow'; $consoleMessage = "   $Message" }
        'ERROR'   { $color = 'Red';    $consoleMessage = "   $Message" }
        'PROMPT'  { $color = 'White';  $consoleMessage = "   $Message" }
        'RESULT'  { $color = 'White';  }
    }

    if ($script:uiMode) {
        # UI Logging requires using the dispatcher to safely update the UI from the script's thread.
        try {
            $script:TxtOutputLog.Dispatcher.InvokeAsync({
                param($logMsg, $logColor)
                
                $run = New-Object System.Windows.Documents.Run($logMsg + "`n")
                $run.Foreground = $logColor
                $paragraph = New-Object System.Windows.Documents.Paragraph($run)
                $paragraph.Margin = "0"

                $script:TxtOutputLog.Document.Blocks.Add($paragraph)
                $script:TxtOutputLog.ScrollToEnd()
            }, "Normal", @($consoleMessage, $color)) | Out-Null
        } catch {
            Write-Warning "Failed to write to UI log: $($_.Exception.Message)"
        }
    } else {
        # Fallback to console if not in UI mode
        if ($NoNewLine) {
            Write-Host $consoleMessage -ForegroundColor $color -NoNewline
        } else {
            Write-Host $consoleMessage -ForegroundColor $color
        }
    }
}
#endregion

#region Configuration Management
function Reset-OptimusConfiguration {
    Write-Log -Message "Entering Function: Reset-OptimusConfiguration" -Level 'DEBUG'
    $configDir = Join-Path -Path $env:USERPROFILE -ChildPath ".optimus"
    if (-not (Test-Path -Path $configDir)) {
        Write-Log -Message "Configuration directory does not exist. No reset needed." -Level 'DEBUG'
        return $true
    }

    # In UI mode, we would pop a dialog. For now, using Write-Log for prompts.
    Write-Log -Message "`n--- Optimus Configuration Reset ---" -Level 'WARN'
    Write-Log -Message "[1] Reset Configuration Only (deletes API key, server list, and model selection)" -Level 'PROMPT'
    Write-Log -Message "[2] Remove all Analysis Reports" -Level 'PROMPT'
    Write-Log -Message "[3] Full Reset (deletes configuration AND all past analysis reports)" -Level 'PROMPT'
    Write-Log -Message "[Q] Quit / Cancel" -Level 'PROMPT'

    Write-Log -Message "Enter your choice: " -Level 'PROMPT' -NoNewLine
    $choice = Read-Host
    Write-Host ""
    Write-Log -Message "User Input: $choice" -Level 'DEBUG'

    switch ($choice) {
        '1' {
            Write-Log -Message "Are you sure you want to delete the API key, server list, and model selection? (Y/N): " -Level 'PROMPT' -NoNewLine
            $confirm = Read-Host
            Write-Host ""
            Write-Log -Message "User Input: $confirm" -Level 'DEBUG'
            if ($confirm -match '^[Yy]$') {
                try {
                    $serverFile = Join-Path -Path $configDir -ChildPath "servers.json"
                    $apiKeyFile = Join-Path -Path $configDir -ChildPath "api.config"
                    $lastPathFile = Join-Path -Path $configDir -ChildPath "lastpath.config"
                    $modelFile = Join-Path -Path $configDir -ChildPath "model.config"
                    if (Test-Path $serverFile) { Remove-Item -Path $serverFile -Force; Write-Log -Message "Server list deleted." -Level 'DEBUG' }
                    if (Test-Path $apiKeyFile) { Remove-Item -Path $apiKeyFile -Force; Write-Log -Message "API Key deleted." -Level 'DEBUG' }
                    if (Test-Path $lastPathFile) { Remove-Item -Path $lastPathFile -Force; Write-Log -Message "Last path config deleted." -Level 'DEBUG' }
                    if (Test-Path $modelFile) { Remove-Item -Path $modelFile -Force; Write-Log -Message "Model configuration deleted." -Level 'DEBUG' }
                    Write-Log -Message "Configuration reset successfully." -Level 'SUCCESS'
                    return $true
                } catch {
                    Write-Log -Message "Failed to delete configuration files: $($_.Exception.Message)" -Level 'ERROR'
                    return $false
                }
            } else {
                return $false
            }
        }
        '2' {
            Write-Log -Message "Are you sure you want to delete ALL past analysis reports? This action cannot be undone. (Y/N): " -Level 'PROMPT' -NoNewLine
            $confirm = Read-Host
            Write-Host ""
            Write-Log -Message "User Input: $confirm" -Level 'DEBUG'
            if ($confirm -match '^[Yy]$') {
                try {
                    $analysisDir = Join-Path -Path $configDir -ChildPath "Analysis"
                    if (Test-Path $analysisDir) {
                        Remove-Item -Path $analysisDir -Recurse -Force
                        Write-Log -Message "All analysis reports have been deleted." -Level 'SUCCESS'
                    } else {
                        Write-Log -Message "Analysis reports directory not found. Nothing to delete." -Level 'INFO'
                    }
                    return $true
                } catch {
                    Write-Log -Message "Failed to remove the analysis reports directory: $($_.Exception.Message)" -Level 'ERROR'
                    return $false
                }
            } else {
                return $false
            }
        }
        '3' {
            Write-Log -Message "WARNING: This will delete ALL configuration AND all saved analysis reports. This action cannot be undone. Are you absolutely sure? (Y/N): " -Level 'PROMPT' -NoNewLine
            $confirm = Read-Host
            Write-Host ""
            Write-Log -Message "User Input: $confirm" -Level 'DEBUG'
            if ($confirm -match '^[Yy]$') {
                try {
                    Remove-Item -Path $configDir -Recurse -Force
                    Write-Log -Message "Full reset complete. The '.optimus' directory has been removed." -Level 'SUCCESS'
                    return $true
                } catch {
                    Write-Log -Message "Failed to remove the .optimus directory: $($_.Exception.Message)" -Level 'ERROR'
                    return $false
                }
            } else {
                return $false
            }
        }
        default {
            return $false
        }
    }
}

function Initialize-Configuration {
    Write-Log -Message "Entering Function: Initialize-Configuration" -Level 'DEBUG'
    Write-Log -Message "Initializing Optimus configuration..." -Level 'DEBUG'
    try {
        $userProfile = $env:USERPROFILE
        $configDir = Join-Path -Path $userProfile -ChildPath ".optimus"
        $analysisBaseDir = Join-Path -Path $configDir -ChildPath "Analysis"
        $serverFile = Join-Path -Path $configDir -ChildPath "servers.json"
        $apiKeyFile = Join-Path -Path $configDir -ChildPath "api.config"
        $lastPathFile = Join-Path -Path $configDir -ChildPath "lastpath.config"
        $modelFile = Join-Path -Path $configDir -ChildPath "model.config"

        foreach($dir in @($configDir, $analysisBaseDir)){ if (-not (Test-Path -Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null } }
        if (-not (Test-Path -Path $serverFile)) { Set-Content -Path $serverFile -Value "[]" | Out-Null }
        
        $script:OptimusConfig = @{
            AnalysisBaseDir = $analysisBaseDir
            ServerFile      = $serverFile
            ApiKeyFile      = $apiKeyFile
            LastPathFile    = $lastPathFile
            ModelFile       = $modelFile
        }
        Write-Log -Message "Configuration initialized successfully." -Level 'DEBUG'
        return $true
    }
    catch { Write-Log -Message "Could not initialize configuration: $($_.Exception.Message)" -Level 'ERROR'; return $false }
}

function Get-And-Set-ApiKey {
    param([switch]$ForcePrompt)
    Write-Log -Message "Entering Function: Get-And-Set-ApiKey" -Level 'DEBUG'
    Write-Log -Message "Checking for Gemini API Key..." -Level 'DEBUG'
    $apiKeyFile = $script:OptimusConfig.ApiKeyFile
    if ((Test-Path -Path $apiKeyFile) -and (-not $ForcePrompt)) {
        try {
            $keyContent = Get-Content -Path $apiKeyFile
            if (-not [string]::IsNullOrWhiteSpace($keyContent)) {
                $script:GeminiApiKey = $keyContent | ConvertTo-SecureString
                Write-Log -Message "API Key loaded successfully." -Level 'DEBUG'
                return $true
            }
        }
        catch { Write-Log -Message "Could not read existing API key. It may be corrupt. Please re-enter." -Level 'WARN' }
    }

    Write-Log -Message "`nTo use Optimus, you need a Gemini API key." -Level 'INFO'
    Write-Log -Message "You can create one for free at Google AI Studio:" -Level 'INFO'
    Write-Log -Message "https://aistudio.google.com/app/apikey" -Level 'PROMPT'

    while ($true) {
        Write-Log -Message "`nPlease enter your Gemini API Key: " -Level 'PROMPT' -NoNewLine
        $secureKey = Read-Host -AsSecureString
        Write-Host ""
        Write-Log -Message "User Input: [SECURE KEY ENTERED]" -Level 'DEBUG'
        if ($secureKey.Length -gt 0) {
            try {
                $secureKey | ConvertFrom-SecureString | Set-Content -Path $apiKeyFile
                $script:GeminiApiKey = $secureKey
                Write-Log -Message "API Key has been validated and saved securely." -Level 'SUCCESS'
                return $true
            }
            catch {
                Write-Log -Message "Failed to save API Key: $($_.Exception.Message)" -Level 'ERROR'
                return $false
            }
        } else {
            Write-Log -Message "API Key cannot be empty." -Level 'ERROR'
            Write-Log -Message "Try again? (Y/N): " -Level 'PROMPT' -NoNewLine
            $retry = Read-Host
            Write-Host ""
            Write-Log -Message "User Input: $retry" -Level 'DEBUG'
            if ($retry -notmatch '^[Yy]$') { return $false }
        }
    }
}

function Get-And-Set-Model {
    param([switch]$ForcePrompt)
    Write-Log -Message "Entering Function: Get-And-Set-Model" -Level 'DEBUG'
    $modelFile = $script:OptimusConfig.ModelFile

    if ((Test-Path -Path $modelFile) -and (-not $ForcePrompt)) {
        $modelName = Get-Content -Path $modelFile
        if (-not [string]::IsNullOrWhiteSpace($modelName)) {
            Write-Log -Message "Using configured model: '$modelName'" -Level 'DEBUG'
            return $modelName
        }
    }

    Write-Log -Message "`nPlease select the Gemini model to use for all future analyses:" -Level 'INFO'
    Write-Log -Message "This can be changed later using the Configuration menu." -Level 'INFO'
    Write-Log -Message "   [1] Gemini 1.5 Flash (Fastest, good for general use - Default)" -Level 'PROMPT'
    Write-Log -Message "   [2] Gemini 1.5 Pro (Most powerful, for complex analysis)" -Level 'PROMPT'
    
    $modelChoice = $null
    while (-not $modelChoice) {
        Write-Log -Message "   Enter your choice: " -Level 'PROMPT' -NoNewLine
        $choice = Read-Host
        Write-Host ""
        Write-Log -Message "User Input: $choice" -Level 'DEBUG'
        switch ($choice) {
            '1' { $modelChoice = 'gemini-1.5-flash-latest' }
            '2' { $modelChoice = 'gemini-1.5-pro-latest' }
            default { Write-Log -Message "Invalid selection. Please enter 1 or 2." -Level 'ERROR' }
        }
    }

    try {
        Set-Content -Path $modelFile -Value $modelChoice
        Write-Log -Message "Model set to '$modelChoice'. This will be used for all future runs." -Level 'SUCCESS'
        return $modelChoice
    } catch {
        Write-Log -Message "Failed to save model configuration: $($_.Exception.Message)" -Level 'ERROR'
        return $null
    }
}
#endregion

#region Environment & Prerequisite Checks
function Invoke-OptimusVersionCheck {
    param(
        [string]$CurrentVersion
    )
    Write-Log -Message "Entering Function: Invoke-OptimusVersionCheck" -Level 'DEBUG'
    
    try {
        $repoUrl = "https://raw.githubusercontent.com/BrennanWebb/powershell/main/Production/Optimus.ps1"
        Write-Log -Message "Checking for new version at: $repoUrl" -Level 'DEBUG'

        $webContent = Invoke-WebRequest -Uri $repoUrl -UseBasicParsing -TimeoutSec 10 | Select-Object -ExpandProperty Content
        if ($webContent -match "Version:\s*([^\s]+)") {
            $latestVersionStr = $matches[1]
            Write-Log -Message "Latest version found online: '$latestVersionStr'" -Level 'DEBUG'
            
            $cleanCurrent = ($CurrentVersion -split '-')[0]
            $cleanLatest = ($latestVersionStr -split '-')[0]

            if ([System.Version]$cleanLatest -gt [System.Version]$cleanCurrent) {
                Write-Log -Message "A new version of Optimus is available! (Current: v$CurrentVersion, Latest: v$latestVersionStr)" -Level 'WARN'
                Write-Log -Message "You can download it from: https://github.com/BrennanWebb/powershell/blob/main/Production/Optimus.ps1" -Level 'WARN'
            } else {
                Write-Log -Message "Optimus is up to date." -Level 'DEBUG'
            }
        }
    }
    catch {
        Write-Log -Message "Could not check for a new version. This can happen if GitHub is unreachable or there is no internet connection." -Level 'DEBUG'
    }
}

function Test-PowerShellVersion {
    Write-Log -Message "Entering Function: Test-PowerShellVersion" -Level 'DEBUG'
    if ($PSVersionTable.PSVersion.Major -lt 5 -or ($PSVersionTable.PSVersion.Major -eq 5 -and $PSVersionTable.PSVersion.Minor -lt 1)) {
        Write-Log -Message "This script requires PowerShell version 5.1 or higher. You are running version $($PSVersionTable.PSVersion)." -Level 'ERROR'
        return $false
    }
    return $true
}

function Test-WindowsEnvironment {
    Write-Log -Message "Entering Function: Test-WindowsEnvironment" -Level 'DEBUG'
    if ($env:OS -ne 'Windows_NT' -or $PSVersionTable.PSEdition -ne 'Desktop') {
        Write-Log -Message "This script requires a Windows Desktop environment (PowerShell 5.1+) to run the graphical interface." -Level 'ERROR'
        return $false
    }
    return $true
}

function Test-InternetConnection {
    param([string]$HostName = "googleapis.com")
    Write-Log -Message "Entering Function: Test-InternetConnection" -Level 'DEBUG'
    try {
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        $asyncResult = $tcpClient.BeginConnect($HostName, 443, $null, $null)
        $success = $asyncResult.AsyncWaitHandle.WaitOne(3000, $true)
        if ($success) { $tcpClient.EndConnect($asyncResult); $tcpClient.Close(); return $true } 
        else { $tcpClient.Close(); throw "Connection to $($HostName):443 timed out." }
    }
    catch {
        Write-Log -Message "Could not establish an internet connection to '$($HostName)'. The script will likely fail when contacting the Gemini API." -Level 'WARN'
        return $false
    }
}

function Test-SqlServerModule {
    Write-Log -Message "Entering Function: Test-SqlServerModule" -Level 'DEBUG'
    if (Get-Module -Name SqlServer -ListAvailable) {
        try { Import-Module SqlServer -ErrorAction Stop; return $true }
        catch { Write-Log -Message "Failed to import 'SqlServer' module: $($_.Exception.Message)" -Level 'ERROR'; return $false }
    } else {
        Write-Log -Message "The 'SqlServer' module is not installed." -Level 'WARN'
        # In UI mode, we can't do an interactive install prompt easily, so we just fail.
        Write-Log -Message "Please install it manually using: Install-Module -Name SqlServer -Scope CurrentUser" -Level 'ERROR'
        return $false
    }
}
#endregion

#region Core SQL, Validation & File Functions
function Test-SqlServerConnection {
    param([string]$ServerInstance)
    Write-Log -Message "Entering Function: Test-SqlServerConnection for server '$ServerInstance'" -Level 'DEBUG'
    Write-Log -Message "Testing connection to '$ServerInstance'..." -Level 'INFO'
    try { Invoke-Sqlcmd -ServerInstance $ServerInstance -Query "SELECT @@VERSION" -QueryTimeout 5 -TrustServerCertificate -ErrorAction Stop | Out-Null; Write-Log -Message "Connection successful!" -Level 'SUCCESS'; return $true }
    catch { Write-Log -Message "Failed to connect to '$ServerInstance': $($_.Exception.Message)" -Level 'ERROR'; return $false }
}

function Get-SqlServerVersion {
    param([string]$ServerInstance)
    Write-Log -Message "Entering Function: Get-SqlServerVersion" -Level 'DEBUG'
    try {
        return (Invoke-Sqlcmd -ServerInstance $ServerInstance -Query "SELECT @@VERSION" -TrustServerCertificate).Item(0)
    } catch {
        Write-Log -Message "Could not retrieve SQL Server version details." -Level 'WARN'
        return "Unknown"
    }
}

function Get-MasterExecutionPlan {
    param($ServerInstance, $DatabaseContext, $FullQueryText, [switch]$IsActualPlan)
    Write-Log -Message "Entering Function: Get-MasterExecutionPlan" -Level 'DEBUG'
    
    $planCommand = if ($IsActualPlan) { "SET STATISTICS XML ON;" } else { "SET SHOWPLAN_XML ON;" }
    $planType = if ($IsActualPlan) { "Actual" } else { "Estimated" }
    Write-Log -Message "`nGenerating master '$planType' execution plan (this also validates script syntax)..." -Level 'INFO'
    
    $dbContextForCheck = if ([string]::IsNullOrWhiteSpace($DatabaseContext)) { 'master' } else { $DatabaseContext }
    Write-Log -Message "Using database context '$dbContextForCheck' to generate plan." -Level 'DEBUG'

    $cleanQueryText = $FullQueryText.Trim()
    if ($cleanQueryText.ToUpper().EndsWith('GO')) {
        $cleanQueryText = $cleanQueryText.Substring(0, $cleanQueryText.Length - 2).Trim()
    }
    
    $planQuery = "$planCommand`nGO`n$cleanQueryText"
    try {
        $planResult = Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $dbContextForCheck -TrustServerCertificate -Query $planQuery -MaxCharLength ([int]::MaxValue) -ErrorAction Stop
        
        $planFragments = @()
        foreach ($resultSet in $planResult) {
            if ($resultSet) {
                $potentialPlan = $resultSet.Item(0)
                if ($potentialPlan -is [string] -and $potentialPlan -like '<*showplan*>') {
                    $planFragments += $potentialPlan
                }
            }
        }
        
        if ($planFragments.Count -eq 0) { throw "Could not find a valid execution plan string in the results from SQL Server." }

        $masterPlanXml = "<MasterShowPlan>" + ($planFragments -join '') + "</MasterShowPlan>"
        [xml]$masterPlanXml | Out-Null # Validate XML
        Write-Log -Message "Successfully generated and validated master execution plan for all statements." -Level 'SUCCESS'
        return $masterPlanXml
    }
    catch {
        Write-Log -Message "The SQL script is invalid or failed to execute. SQL Server could not compile it. Error: $($_.Exception.Message)" -Level 'ERROR'
        return $null
    }
}

function Get-ObjectsFromPlan {
    param([xml]$MasterPlan, [object]$NamespaceManager)
    Write-Log -Message "Entering Function: Get-ObjectsFromPlan" -Level 'DEBUG'
    Write-Log -Message "Parsing execution plan to identify unique database objects..." -Level 'INFO'
    try {
        $objectNodes = $MasterPlan.SelectNodes("//sql:Object", $NamespaceManager)
        $uniqueObjectNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        foreach($node in $objectNodes) {
            $db = $node.GetAttribute("Database")
            $schema = $node.GetAttribute("Schema")
            $table = $node.GetAttribute("Table")
            
            if ($table -notlike "#*" -and -not ([string]::IsNullOrWhiteSpace($db)) -and -not ([string]::IsNullOrWhiteSpace($schema))) {
                $uniqueObjectNames.Add("$db.$schema.$table".Replace('[','').Replace(']','')) | Out-Null
            }
        }

        $finalList = $uniqueObjectNames | Sort-Object
        
        if ($DebugMode -and $finalList.Count -gt 0) {
            $displayObjects = $finalList | ForEach-Object { $parts = $_.Split('.'); [pscustomobject]@{ Database = $parts[0]; Schema = $parts[1]; Name = $parts[2] } }
            Write-Log -Message "The following unique user objects were found in the execution plan:`n$($displayObjects | Format-Table -AutoSize | Out-String)" -Level 'DEBUG'
        }

        Write-Log -Message "Identified $($finalList.Count) unique objects for schema collection." -Level 'SUCCESS'
        return $finalList
    } catch {
        Write-Log -Message "Failed to parse objects from the execution plan: $($_.Exception.Message)" -Level 'ERROR'
        return @()
    }
}

function Get-ObjectSchema {
    param([string]$ServerInstance, [string]$DatabaseName, [string]$SchemaName, [string]$ObjectName)
    Write-Log -Message "Now collecting schema for: $DatabaseName.$SchemaName.$ObjectName" -Level 'DEBUG'
    
    $fullObjectName = "[$SchemaName].[$ObjectName]"
    $schemaText = "--- Schema For Table: $SchemaName.$ObjectName ---`n"
    $columnResult = $null

    try {
        $columnQuery = "SELECT name, system_type_name, max_length, [precision], scale, is_nullable FROM sys.dm_exec_describe_first_result_set(N'SELECT * FROM $fullObjectName', NULL, 0);"
        $columnResult = Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $DatabaseName -TrustServerCertificate -Query $columnQuery -ErrorAction Stop
    } catch {
        Write-Log -Message "Primary schema collection method failed for '$fullObjectName'. Attempting fallback." -Level 'DEBUG'
        try {
            $fallbackQuery = "SELECT c.name, t.name AS system_type_name, c.max_length, c.precision, c.scale, c.is_nullable FROM sys.columns c JOIN sys.types t ON c.user_type_id = t.user_type_id WHERE c.object_id = OBJECT_ID(@FullObjectName) ORDER BY c.column_id;"
            $params = @{ FullObjectName = "$DatabaseName.$fullObjectName" }
            $columnResult = Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $DatabaseName -TrustServerCertificate -Query $fallbackQuery -Variable $params -ErrorAction Stop
        } catch { Write-Log -Message "Could not get COLUMN schema for '$fullObjectName' in db '$DatabaseName'. Error: $($_.Exception.Message)" -Level 'WARN' }
    }

    if ($columnResult) {
        $schemaText += "COLUMNS:`n"
        foreach($col in $columnResult) {
            $isNullable = if ($col.is_nullable) { 'YES' } else { 'NO' }
            $schemaText += "name: $($col.name), type: $($col.system_type_name), length: $($col.max_length), nullable: $isNullable`n"
        }
    }

    try {
        $indexQuery = "SELECT i.name AS IndexName, i.type_desc AS IndexType, STUFF((SELECT ', ' + c.name FROM sys.index_columns ic JOIN sys.columns c ON ic.object_id = c.object_id AND ic.column_id = c.column_id WHERE ic.object_id = i.object_id AND ic.index_id = i.index_id AND ic.is_included_column = 0 ORDER BY ic.key_ordinal FOR XML PATH('')), 1, 2, '') AS KeyColumns, STUFF((SELECT ', ' + c.name FROM sys.index_columns ic JOIN sys.columns c ON ic.object_id = c.object_id AND ic.column_id = c.column_id WHERE ic.object_id = i.object_id AND ic.index_id = i.index_id AND ic.is_included_column = 1 ORDER BY ic.key_ordinal FOR XML PATH('')), 1, 2, '') AS IncludedColumns FROM sys.indexes i WHERE i.object_id = OBJECT_ID('$fullObjectName');"
        $indexResult = Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $DatabaseName -TrustServerCertificate -Query $indexQuery -ErrorAction Stop
        if ($indexResult -and $indexResult.Count -gt 0) {
            $schemaText += "`nINDEXES:`n"
            foreach($idx in $indexResult) {
                $idxLine = "IndexName: $($idx.IndexName), Type: $($idx.IndexType), KeyColumns: $($idx.KeyColumns)"
                if (-not [string]::IsNullOrWhiteSpace($idx.IncludedColumns)) { $idxLine += ", IncludedColumns: $($idx.IncludedColumns)" }
                $schemaText += $idxLine + "`n"
            }
        }
    } catch { Write-Log -Message "Could not get INDEX information for '$fullObjectName'." -Level 'WARN' }
    
    return $schemaText + "`n"
}

function Invoke-GeminiAnalysis {
    param([string]$ModelName, [securestring]$ApiKey, [string]$FullSqlText, [string]$ConsolidatedSchema, [string]$MasterPlanXml, [string]$SqlServerVersion)
    Write-Log -Message "Entering Function: Invoke-GeminiAnalysis" -Level 'DEBUG'
    Write-Log -Message "Sending full script to Gemini for holistic analysis..." -Level 'INFO'

    $plainTextApiKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ApiKey))
    $uri = "https://generativelanguage.googleapis.com/v1beta/models/$($ModelName):generateContent?key=$plainTextApiKey"

    $prompt = @"
You are an expert T-SQL performance tuning assistant. You will be provided with the specific SQL Server version. You MUST ensure that any T-SQL syntax you generate is valid for that version.

Your Core Mandate:
Your ONLY task is to return the complete, original T-SQL script provided below. You will add T-SQL comment blocks containing your analysis directly above any statements you identify for improvement. Think of this as a text transformation task: the original text is your input, and the identical text with your added comments is the only output.

Your Golden Rule:
You MUST NOT change the original T-SQL code. The final output must contain the original, unmodified T-SQL script with only your comments added.

Your Task:
Your primary goal is to identify ALL potential performance improvements for each T-SQL statement. For any statement that can be improved, you will add a single T-SQL block comment immediately above it. Within this block comment, you may provide one or more recommendations. If a statement is already optimal, do not add a comment for it.

Recommendation Categories:
You should consider the following categories of recommendations. For any given T-SQL statement, multiple recommendations from different categories might be valid. For example, a statement could be improved with both a query rewrite AND a new index. Present all valid options.
1.  **Non-Invasive Query Rewrites:** These are the most desirable. Look for opportunities to make predicates SARGable, simplify logic, or use more efficient patterns that are valid for the specified SQL Server version.
2.  **Indexing Improvements:**
    * **Alter Existing Index:** If an existing index can be modified (e.g., adding an INCLUDE column) to better serve the query, provide the necessary `DROP` and `CREATE` DDL.
    * **Create New Index:** If no existing index is a suitable candidate for alteration, recommend a new, covering index. Provide the complete `CREATE INDEX` DDL.

Comment Formatting:
Every analysis comment block you add MUST use the following structure. The block starts with a general "Optimus Analysis" header. Inside, each distinct recommendation is numbered and contains the three required sections. This allows for multiple, independent suggestions for the same statement. Important: All text inside the comment block must be plain text. Do not use any Markdown formatting like **bolding** or `backticks`.

/*
--- Optimus Analysis ---

[1] Recommendation
    - Problem: A brief, clear explanation of the first performance issue.
    - Recommended Code: The suggested T-SQL query rewrite or DDL syntax for the first issue.
    - Reasoning: An explanation of why this specific recommendation improves performance.

[2] Recommendation
    - Problem: A brief, clear explanation of a second, distinct performance issue.
    - Recommended Code: The alternative or additional code for the second recommendation.
    - Reasoning: An explanation of why this second recommendation is also a valid performance improvement.
*/

Final Output Rules:
- Your response MUST be the complete, original T-SQL script from start to finish. Do not omit any part of the original script for any reason.
- For statements that require improvement, insert your formatted analysis comment block directly above the statement.
- For statements that are already optimal, include the original T-SQL for that statement without any comment.
- Your entire response must be ONLY the T-SQL script text. Do not include any conversational text, greetings, or explanations outside of the T-SQL comments.

--- SQL SERVER VERSION ---
$SqlServerVersion

--- FULL T-SQL SCRIPT ---
$FullSqlText

--- CONSOLIDATED OBJECT SCHEMAS AND DEFINITIONS ---
$ConsolidatedSchema

--- MASTER EXECUTION PLAN ---
$MasterPlanXml
"@
    $promptPath = Join-Path -Path $script:AnalysisPath -ChildPath "_FinalAIPrompt.txt"
    try { $prompt | Set-Content -Path $promptPath -Encoding UTF8; Write-Log -Message "Final AI prompt saved for review at: $promptPath" -Level 'DEBUG' } catch { }

    $bodyObject = @{ contents = @( @{ parts = @( @{ text = $prompt } ) } ) }
    $bodyJson = $bodyObject | ConvertTo-Json -Depth 10

    try {
        $response = Invoke-RestMethod -Uri $uri -Method Post -Body $bodyJson -ContentType 'application/json' -ErrorAction Stop
        $rawAiResponse = $response.candidates[0].content.parts[0].text
        $cleanedScript = $rawAiResponse -replace '(?i)^```sql\s*','' -replace '```\s*$',''
        Write-Log -Message "AI analysis complete." -Level 'SUCCESS'
        return $cleanedScript
    } catch {
        Write-Log -Message "Failed to get response from Gemini API." -Level 'ERROR'
        $errorText = ($_.Exception.Response.GetResponseStream() | New-Object System.IO.StreamReader).ReadToEnd()
        Write-Log -Message "API Error Details: $errorText" -Level 'ERROR'
        return $null
    }
}


function New-AnalysisSummary {
    param([string]$TunedScript, [int]$TotalStatementCount, [string]$AnalysisPath)
    Write-Log -Message "Entering Function: New-AnalysisSummary" -Level 'DEBUG'
    
    try {
        $summaryContent = "--- Optimus Analysis Summary ---`nTimestamp: $(Get-Date)`n`n"
        $recommendationBlocks = [regex]::Matches($TunedScript, '(?s)\/\*\s*--- Optimus Analysis ---(.*?)\*\/')
        $totalRecommendations = 0; foreach ($block in $recommendationBlocks) { $totalRecommendations += ([regex]::Matches($block.Value, '\[\d+\]\s*Recommendation')).Count }

        $summaryContent += "Total Statements Analyzed: $TotalStatementCount`n"
        $summaryContent += "Statements with Recommendations: $($recommendationBlocks.Count)`n"
        $summaryContent += "Total Individual Recommendations: $totalRecommendations`n`n"

        if ($recommendationBlocks.Count -gt 0) {
            $summaryContent += "--- Summary of Findings ---`n"
            $findingIndex = 1
            foreach ($block in $recommendationBlocks) {
                foreach ($problem in [regex]::Matches($block.Value, 'Problem:(.*?)(?=\s*-\s*Recommended Code:|\s*$)')) {
                    $summaryContent += "$($findingIndex). $($problem.Groups[1].Value.Trim())`n"; $findingIndex++
                }
            }
        }
        $summaryPath = Join-Path -Path $AnalysisPath -ChildPath "_AnalysisSummary.txt"
        $summaryContent | Out-File -FilePath $summaryPath -Encoding UTF8
        Write-Log -Message "Analysis summary report generated at: '$summaryPath'" -Level 'DEBUG'
    } catch {
        Write-Log -Message "Could not generate analysis summary report. Error: $($_.Exception.Message)" -Level 'WARN'
    }
}
#endregion

#region Core Analysis Workflow
function Invoke-AnalysisWorkflow {
    param(
        [Parameter(Mandatory=$true)] [string[]]$SqlFilePaths,
        [Parameter(Mandatory=$true)] [string]$SelectedServer,
        [Parameter(Mandatory=$true)] [string]$ChosenModel,
        [Parameter(Mandatory=$true)] [securestring]$GeminiApiKey,
        [Parameter(Mandatory=$true)] [bool]$UseActualPlan
    )

    # Create the model-specific and batch parent folders
    $sanitizedModelName = $ChosenModel -replace '[.-]', '_'
    $modelSpecificPath = Join-Path -Path $script:OptimusConfig.AnalysisBaseDir -ChildPath $sanitizedModelName
    if (-not (Test-Path -Path $modelSpecificPath)) { New-Item -Path $modelSpecificPath -ItemType Directory -Force | Out-Null }

    $batchTimestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $batchFolderPath = Join-Path -Path $modelSpecificPath -ChildPath $batchTimestamp
    New-Item -Path $batchFolderPath -ItemType Directory -Force | Out-Null
    Write-Log -Message "`nCreated batch analysis folder: $batchFolderPath" -Level 'SUCCESS'

    # Loop through each selected file
    foreach ($sqlFilePath in $sqlFilePaths) {
        try {
            $fileNameOnly = [System.IO.Path]::GetFileName($sqlFilePath)
            Write-Log -Message "`n--- Starting Analysis for: $fileNameOnly ---" -Level 'SUCCESS'

            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($sqlFilePath)
            $script:AnalysisPath = Join-Path -Path $batchFolderPath -ChildPath $baseName
            New-Item -Path $script:AnalysisPath -ItemType Directory -Force | Out-Null

            $script:LogFilePath = Join-Path -Path $script:AnalysisPath -ChildPath "ExecutionLog.txt"
            "# Optimus v$($script:CurrentVersion) Execution Log | File: $fileNameOnly | Started: $(Get-Date)" | Out-File -FilePath $script:LogFilePath -Encoding utf8
            
            Write-Log -Message "Created analysis directory: '$($script:AnalysisPath)'" -Level 'INFO'
            $sqlVersion = Get-SqlServerVersion -ServerInstance $selectedServer
            Write-Log -Message "Detected SQL Server Version: $sqlVersion" -Level 'DEBUG'
            
            $sqlQueryText = Get-Content -Path $sqlFilePath -Raw
        
            # 1. Get Master Plan
            $initialDbContext = ([regex]::Match($sqlQueryText, '(?im)^\s*USE\s+\[?([\w\d_]+)\]?')).Groups[1].Value
            $masterPlanXml = Get-MasterExecutionPlan -ServerInstance $selectedServer -DatabaseContext $initialDbContext -FullQueryText $sqlQueryText -IsActualPlan:$UseActualPlan
            if (-not $masterPlanXml) { Write-Log -Message "Could not generate a master plan for $fileNameOnly. Skipping." -Level 'ERROR'; continue }
            $masterPlanXml | Set-Content -Path (Join-Path $script:AnalysisPath -ChildPath "_MasterPlan.xml") -Encoding UTF8

            # 2. Parse unique object names from the plan
            [xml]$masterPlan = $masterPlanXml
            $ns = New-Object System.Xml.XmlNamespaceManager($masterPlan.NameTable)
            $ns.AddNamespace("sql", "http://schemas.microsoft.com/sqlserver/2004/07/showplan")
            [string[]]$uniqueObjectNames = @(Get-ObjectsFromPlan -MasterPlan $masterPlan -NamespaceManager $ns)
            $statementNodes = $masterPlan.SelectNodes("//sql:StmtSimple", $ns)
            
            # 3. Build the consolidated schema document
            $consolidatedSchema = ""
            if ($uniqueObjectNames.Count -gt 0) {
                Write-Log -Message "Starting schema collection for all objects..." -Level 'INFO'
                foreach ($dbGroup in ($uniqueObjectNames | Group-Object { ($_ -split '\.')[0] })) {
                    if ($dbGroup.Name -eq 'mssqlsystemresource') { Write-Log -Message "Skipping schema collection for internal db: 'mssqlsystemresource'." -Level 'DEBUG'; continue }
                    Write-Log -Message "Querying database '$($dbGroup.Name)'..." -Level 'INFO'
                    foreach ($objName in $dbGroup.Group) {
                        $parts = $objName.Split('.'); $consolidatedSchema += Get-ObjectSchema -ServerInstance $selectedServer -DatabaseName $parts[0] -SchemaName $parts[1] -ObjectName $parts[2]
                    }
                }
            } else { Write-Log -Message "No user database objects were found in the execution plan. Halting analysis." -Level 'WARN'; continue }

            if ([string]::IsNullOrWhiteSpace($consolidatedSchema)) { Write-Log -Message "Schema collection resulted in an empty document. Halting analysis." -Level 'WARN'; continue }
            $consolidatedSchema | Set-Content -Path (Join-Path $script:AnalysisPath -ChildPath "_ConsolidatedSchema.txt") -Encoding UTF8

            # 4. Make single call to AI
            $finalScript = Invoke-GeminiAnalysis -ModelName $ChosenModel -ApiKey $GeminiApiKey -FullSqlText $sqlQueryText -ConsolidatedSchema $consolidatedSchema -MasterPlanXml $masterPlanXml -SqlServerVersion $sqlVersion
            
            # 5. Process and save the final result
            if ($finalScript) {
                $finalScript = $finalScript.Trim()
                $tunedScriptPath = Join-Path -Path $script:AnalysisPath -ChildPath "${baseName}_tuned.sql"
                $finalScript | Out-File -FilePath $tunedScriptPath -Encoding UTF8
                New-AnalysisSummary -TunedScript $finalScript -TotalStatementCount $statementNodes.Count -AnalysisPath $script:AnalysisPath
                Write-Log -Message "Analysis complete for $fileNameOnly." -Level 'SUCCESS'
            } else { Write-Log -Message "Analysis halted for $fileNameOnly due to an error or empty response from the AI." -Level 'ERROR' }
        }
        catch { Write-Log -Message "CRITICAL UNHANDLED ERROR during analysis of '$fileNameOnly': $($_.Exception.Message). Moving to next file." -Level 'ERROR' }
    } # End foreach loop

    Write-Log -Message "`n--- Batch Analysis Complete ---" -Level 'SUCCESS'
    Write-Log -Message "All analysis folders for this batch are located in:" -Level 'SUCCESS'
    Write-Log -Message "$batchFolderPath" -Level 'RESULT'

    try { Invoke-Item -Path $batchFolderPath } catch { Write-Log -Message "Could not automatically open the batch folder. Please navigate to the path above manually." -Level 'WARN' }
}
#endregion

# --- Main Application Logic ---
function Start-OptimusGUI {
    # Perform prerequisite checks before showing the UI
    $checksPassed = {
        if (-not (Test-WindowsEnvironment)) { return $false }
        if (-not (Test-PowerShellVersion)) { return $false }
        if (-not (Initialize-Configuration)) { return $false }
        if (-not (Test-SqlServerModule)) { return $false }
        return $true
    }.Invoke()

    if (-not $checksPassed) {
        Write-Log -Message "Prerequisite checks failed. Please resolve the issues above. Exiting." -Level 'ERROR'
        Read-Host "Press Enter to exit"
        return
    }

    # Load WPF Assemblies
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

    # Define the UI in XAML
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        x:Name="Window" Title="Optimus T-SQL Tuning Advisor v$($script:CurrentVersion)" Height="700" Width="800" MinHeight="500" MinWidth="600" WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Menu Grid.Row="0" Margin="0,0,0,10">
            <MenuItem Header="_File">
                <MenuItem x:Name="MenuExit" Header="E_xit"/>
            </MenuItem>
            <MenuItem Header="_Configuration">
                <MenuItem x:Name="MenuReset" Header="_Reset Configuration..."/>
                <MenuItem x:Name="MenuSetApiKey" Header="Set API _Key..."/>
                <MenuItem x:Name="MenuSetModel" Header="Set AI _Model..."/>
            </MenuItem>
             <MenuItem Header="_Help">
                <MenuItem x:Name="MenuAbout" Header="_About..."/>
            </MenuItem>
        </Menu>

        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <GroupBox Grid.Column="0" Header="1. Select SQL Source" Margin="0,0,5,0" Padding="5">
                <StackPanel>
                    <RadioButton x:Name="RadioFiles" Content="Analyze Individual File(s)" IsChecked="True"/>
                    <RadioButton x:Name="RadioFolder" Content="Analyze an Entire Folder" Margin="0,5,0,0"/>
                    <TextBox x:Name="TxtSourcePath" Margin="0,10,0,5" IsReadOnly="True" Text="No source selected..."/>
                    <Button x:Name="BtnBrowse" Content="Browse..." HorizontalAlignment="Right" Width="100" Padding="5"/>
                </StackPanel>
            </GroupBox>

            <GroupBox Grid.Column="1" Header="2. Analysis Configuration" Margin="5,0,0,0" Padding="5">
                <Grid>
                     <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <StackPanel Grid.Row="0" Margin="0,0,0,10">
                        <Label Content="SQL Server:"/>
                        <ComboBox x:Name="ComboServers"/>
                    </StackPanel>
                    <StackPanel Grid.Row="1" Margin="0,5,0,10">
                        <Label Content="Execution Plan Type:"/>
                        <RadioButton x:Name="RadioPlanEstimated" Content="Estimated (Recommended)" IsChecked="True" Margin="5,0,0,0"/>
                        <RadioButton x:Name="RadioPlanActual" Content="Actual (Executes Script)" Margin="5,5,0,0"/>
                    </StackPanel>
                     <CheckBox Grid.Row="2" x:Name="CheckDebug" Content="Enable Verbose Debug Logging" VerticalAlignment="Bottom"/>
                </Grid>
            </GroupBox>
        </Grid>
        
        <GroupBox Grid.Row="2" Header="Live Output Log" Margin="0,10,0,0">
             <ScrollViewer VerticalScrollBarVisibility="Auto">
                 <RichTextBox x:Name="TxtOutputLog" IsReadOnly="True" FontFamily="Consolas" FontSize="12" VerticalScrollBarVisibility="Auto">
                     <FlowDocument/>
                 </RichTextBox>
            </ScrollViewer>
        </GroupBox>

        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
             <Button x:Name="BtnStartAnalysis" Content="Start Analysis" IsEnabled="False" FontWeight="Bold" Width="150" Height="30" Padding="5"/>
        </StackPanel>
    </Grid>
</Window>
"@

    # Create UI objects from XAML
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [System.Windows.Markup.XamlReader]::Load($reader)

    # Automatically create PowerShell variables for all named XAML controls
    $xaml.SelectNodes("//*[@*[local-name()='Name']]") | ForEach-Object {
        Set-Variable -Name ($_.Name) -Value $window.FindName($_.Name) -Scope "script"
    }

    # --- UI Event Handlers ---
    $Window.Add_SourceInitialized({
        Write-Log -Message "--- Welcome to Optimus v$($script:CurrentVersion) ---" -Level 'SUCCESS'
        Write-Log -Message "Initializing..." -Level 'INFO'
        Invoke-OptimusVersionCheck -CurrentVersion $script:CurrentVersion
        
        $script:ChosenModel = Get-And-Set-Model
        if (-not $script:ChosenModel) {
            Write-Log -Message "A Gemini model must be configured. Please use the Configuration menu." -Level 'ERROR'
        } else {
             Write-Log -Message "AI Model: '$($script:ChosenModel)'" -Level 'DEBUG'
        }

        if (-not (Get-And-Set-ApiKey)) {
            Write-Log -Message "A Gemini API Key must be configured. Please use the Configuration menu." -Level 'ERROR'
        }
        
        if (-not (Test-InternetConnection)) {
            Write-Log -Message "Internet connection failed. Analysis may not succeed." -Level 'WARN'
        }
        
        # Populate server list
        $script:ComboServers.ItemsSource = Get-Content -Path $script:OptimusConfig.ServerFile | ConvertFrom-Json
    })

    $BtnBrowse.Add_Click({
        $initialDir = [System.Environment]::getFolderPath('MyDocuments')
        if (Test-Path -Path $script:OptimusConfig.LastPathFile) {
            $lastPath = Get-Content -Path $script:OptimusConfig.LastPathFile
            if (Test-Path -Path $lastPath -PathType Container) { $initialDir = $lastPath }
        }

        if ($script:RadioFiles.IsChecked) {
            $dialog = New-Object Microsoft.Win32.OpenFileDialog
            $dialog.InitialDirectory = $initialDir
            $dialog.Filter = "SQL Files (*.sql)|*.sql"
            $dialog.Multiselect = $true
            if ($dialog.ShowDialog()) {
                $script:TxtSourcePath.Text = $dialog.FileNames -join '; '
                $script:TxtSourcePath.Tag = $dialog.FileNames # Store array in Tag property
            }
        } else {
            # Folder browser is more complex, using a simpler one for now
            Add-Type -AssemblyName System.Windows.Forms
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $dialog.SelectedPath = $initialDir
            if ($dialog.ShowDialog() -eq 'OK') {
                $script:TxtSourcePath.Text = $dialog.SelectedPath
                $script:TxtSourcePath.Tag = (Get-ChildItem -Path $dialog.SelectedPath -Filter *.sql).FullName
            }
        }

        # Enable start button if all conditions met
        if (($script:TxtSourcePath.Tag -ne $null) -and ($script:ComboServers.SelectedItem -ne $null)) {
            $script:BtnStartAnalysis.IsEnabled = $true
        }
    })

    $ComboServers.Add_SelectionChanged({
        # Enable start button if all conditions met
        if (($script:TxtSourcePath.Tag -ne $null) -and ($script:ComboServers.SelectedItem -ne $null)) {
            $script:BtnStartAnalysis.IsEnabled = $true
        }
    })

    $BtnStartAnalysis.Add_Click({
        $script:BtnStartAnalysis.IsEnabled = $false
        try {
            # Clear log for new run
            $script:TxtOutputLog.Document.Blocks.Clear()

            # Gather inputs from UI
            $filePaths = $script:TxtSourcePath.Tag
            $server = $script:ComboServers.SelectedItem
            $useActual = $script:RadioPlanActual.IsChecked
            $script:DebugMode = $script:CheckDebug.IsChecked

            # Validation
            if (($null -eq $filePaths) -or ($filePaths.Count -eq 0)) { Write-Log -Message "No valid .sql files selected or found in folder." -Level "ERROR"; return }
            if ([string]::IsNullOrWhiteSpace($server)) { Write-Log -Message "No SQL Server selected." -Level "ERROR"; return }
            if ([string]::IsNullOrWhiteSpace($script:ChosenModel)) { Write-Log -Message "No AI Model is configured. Use the menu to set one." -Level "ERROR"; return }
            if ([string]::IsNullOrWhiteSpace($script:GeminiApiKey)) { Write-Log -Message "No API Key is configured. Use the menu to set one." -Level "ERROR"; return }

            # Run the main analysis workflow
            Invoke-AnalysisWorkflow -SqlFilePaths $filePaths -SelectedServer $server -ChosenModel $script:ChosenModel -GeminiApiKey $script:GeminiApiKey -UseActualPlan:$useActual
        }
        catch {
            Write-Log -Message "An unexpected error occurred in the UI: $($_.Exception.Message)" -Level 'ERROR'
        }
        finally {
            $script:BtnStartAnalysis.IsEnabled = $true
        }
    })
    
    $MenuExit.Add_Click({ $window.Close() })
    $MenuAbout.Add_Click({ [System.Windows.MessageBox]::Show("Optimus T-SQL Tuning Advisor`nVersion: $($script:CurrentVersion)`nDesigner: Brennan Webb", "About Optimus", "OK", "Information") })
    $MenuReset.Add_Click({ Reset-OptimusConfiguration })
    $MenuSetApiKey.Add_Click({ Get-And-Set-ApiKey -ForcePrompt })
    $MenuSetModel.Add_Click({ 
        $script:ChosenModel = Get-And-Set-Model -ForcePrompt
        Write-Log -Message "Model has been updated to '$($script:ChosenModel)'." -Level 'SUCCESS'
    })

    # Show the window
    $window.ShowDialog() | Out-Null
}

# --- Script Entry Point ---
Start-OptimusGUI