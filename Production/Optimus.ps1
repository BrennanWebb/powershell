<#
.SYNOPSIS
    Optimus is a T-SQL advisor that leverages the Gemini AI for performance tuning or standardized code reviews.

.DESCRIPTION
    This script performs a holistic analysis of a T-SQL query. It has two modes:
    1. Tuning: It generates a master execution plan, builds a comprehensive schema document for all referenced objects, and provides performance tuning recommendations.
    2. Code Review: It performs a static analysis of the T-SQL script against a set of standardized best practices for readability, maintainability, and security.

    It can process a T-SQL script from a single file, a list of files, all .sql files in a specified folder, or a raw T-SQL string.
    All execution steps, messages, and recommendations are recorded in a detailed log file for each analysis.

    It also includes a utility to easily import new prompts from .txt files.

.PARAMETER SQLFile
    The path to one or more .sql files to be analyzed. For multiple files, provide a comma-separated list. This parameter
    cannot be used with -FolderPath or -AdhocSQL.

.PARAMETER FolderPath
    The path to a single folder. The script will analyze all .sql files found in this folder (non-recursively). 
    This parameter cannot be used with -SQLFile or -AdhocSQL.

.PARAMETER AdhocSQL
    A string containing the T-SQL query to be analyzed. This is useful for passing queries directly from other applications.
    This parameter cannot be used with -SQLFile or -FolderPath.

.PARAMETER CodeReview
    An optional switch to perform a standardized code review instead of the default performance tuning analysis.
    When this switch is used, the -ServerName and -UseActualPlan parameters are ignored.

.PARAMETER PromptName
    An optional parameter to specify the exact name of a prompt from the 'prompts.json' configuration file, bypassing the interactive prompt selection menu.

.PARAMETER ServerName
    An optional parameter to specify the SQL Server instance for the analysis, bypassing the interactive menu.
    This parameter is only used for performance tuning analysis.

.PARAMETER UseActualPlan
    An optional switch to generate the 'Actual' execution plan during a tuning analysis. This WILL execute the query.
    If not present, the script defaults to 'Estimated' or will prompt in interactive mode. This parameter is only used for performance tuning analysis.

.PARAMETER OpenTunedFile
    An optional switch that opens the final analyzed .sql file using the default OS application (e.g., SSMS) after analysis is complete.

.PARAMETER OpenPlanFile
    An optional switch that opens the generated .sqlplan file using the default OS application (e.g., SSMS or Sentry Plan Explorer) after it is created.
    This parameter is only used for performance tuning analysis.

.PARAMETER ResetConfiguration
    An optional switch to trigger an interactive menu that allows for resetting user configurations.
    This can be used to clear the saved API key, server list, and optionally, all past analysis reports.

.PARAMETER ImportPrompt
    A switch to activate the prompt import utility. Must be used with the -Path parameter. This bypasses all analysis workflows.

.PARAMETER Path
    When used with -ImportPrompt, specifies the path to the .txt file containing the body of the new prompt to be imported.

.PARAMETER DebugMode
    Enables detailed diagnostic output to the console. All messages are always written to the execution log file regardless of this setting.

.EXAMPLE
    .\Optimus.ps1 -FolderPath "C:\My TSQL Projects\Batch1" -OpenPlanFile
    Runs a tuning analysis on all .sql files in the folder and automatically opens the generated .sqlplan file for each script in the default application.

.EXAMPLE
    .\Optimus.ps1 -ImportPrompt -Path "C:\MyPrompts\NewTuningPrompt.txt"
    Starts the prompt import utility to add the prompt from the specified .txt file to the configuration.

.NOTES
    Designer: Brennan Webb & Gemini
    Script Engine: Gemini   
    Version: 3.4.0
    Created: 2025-06-21
    Modified: 2025-10-07
    Change Log:
    - v3.4.0: Enhanced the default tuning prompt to guide the AI to critically evaluate index suggestions and avoid duplicates.
    - v3.3.3: Renamed output files to be based on the source script name for better clarity (e.g., YourScript.sqlplan and YourScript_analyzed.sql).
    - v3.3.2: Implemented a robust XML node selection method (`SelectSingleNode`) to fix a crash in the plan merging logic.
    - v3.3.1: Implemented logic to merge multiple XML plans into a single valid .sqlplan file. Refined interactive prompt to only show in true interactive mode.
    - v3.3.0: Added ability to save a raw .sqlplan file and a new -OpenPlanFile switch to automatically open it in a default application like Sentry Plan Explorer.
    - v3.2.5: Removed literal Markdown from default prompt text to prevent AI confusion.
    - v3.2.4: Added backspace escaping to the manual JSON serializer for full compliance.
    - v3.2.3: Centralized JSON creation to use the manual serializer, ensuring valid output on first run.
    - v3.2.2: Added tab character escaping to the manual JSON serializer to ensure full JSON compliance.
    - v3.2.1: Corrected string escaping logic in the manual JSON serializer.

    Powershell Version: 5.1+
#>
[CmdletBinding(DefaultParameterSetName = 'Interactive')]
param (
    [Parameter(Mandatory=$true, ParameterSetName='Files')]
    [string[]]$SQLFile,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$FolderPath,

    [Parameter(Mandatory=$true, ParameterSetName='Adhoc')]
    [string]$AdhocSQL,

    [Parameter(Mandatory=$false, ParameterSetName='Files')]
    [Parameter(Mandatory=$false, ParameterSetName='Folder')]
    [Parameter(Mandatory=$false, ParameterSetName='Adhoc')]
    [Parameter(Mandatory=$false, ParameterSetName='Interactive')]
    [switch]$CodeReview,

    [Parameter(Mandatory=$false, ParameterSetName='Files')]
    [Parameter(Mandatory=$false, ParameterSetName='Folder')]
    [Parameter(Mandatory=$false, ParameterSetName='Adhoc')]
    [Parameter(Mandatory=$false, ParameterSetName='Interactive')]
    [string]$PromptName,

    [Parameter(Mandatory=$false, ParameterSetName='Files')]
    [Parameter(Mandatory=$false, ParameterSetName='Folder')]
    [Parameter(Mandatory=$false, ParameterSetName='Adhoc')]
    [Parameter(Mandatory=$false, ParameterSetName='Interactive')]
    [string]$ServerName,

    [Parameter(Mandatory=$false, ParameterSetName='Files')]
    [Parameter(Mandatory=$false, ParameterSetName='Folder')]
    [Parameter(Mandatory=$false, ParameterSetName='Adhoc')]
    [Parameter(Mandatory=$false, ParameterSetName='Interactive')]
    [switch]$UseActualPlan,

    [Parameter(Mandatory=$false, ParameterSetName='Files')]
    [Parameter(Mandatory=$false, ParameterSetName='Folder')]
    [Parameter(Mandatory=$false, ParameterSetName='Adhoc')]
    [Parameter(Mandatory=$false, ParameterSetName='Interactive')]
    [switch]$OpenTunedFile,
    
    [Parameter(Mandatory=$false, ParameterSetName='Files')]
    [Parameter(Mandatory=$false, ParameterSetName='Folder')]
    [Parameter(Mandatory=$false, ParameterSetName='Adhoc')]
    [Parameter(Mandatory=$false, ParameterSetName='Interactive')]
    [switch]$OpenPlanFile,

    [Parameter(Mandatory=$false)]
    [switch]$ResetConfiguration,

    [Parameter(Mandatory=$true, ParameterSetName='ImportPrompt')]
    [switch]$ImportPrompt,

    [Parameter(Mandatory=$true, ParameterSetName='ImportPrompt')]
    [ValidateScript({
        if (Test-Path -Path $_ -PathType Leaf) {
            if ($_ -like '*.txt') {
                return $true
            } else {
                throw "The file '$_' is not a .txt file."
            }
        } else {
            throw "The path '$_' does not exist or is not a file."
        }
    })]
    [string]$Path,

    [Parameter(Mandatory=$false)]
    [switch]$DebugMode
)

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

    # 2. Handle all console output.
    # For DEBUG level, only write to console if DebugMode switch is present.
    if ($Level -eq 'DEBUG' -and -not $DebugMode) {
        return
    }

    # Prepare message and color specifically for the console
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

    # Write the formatted message to the console
    if ($NoNewLine) {
        Write-Host $consoleMessage -ForegroundColor $color -NoNewline
    } else {
        Write-Host $consoleMessage -ForegroundColor $color
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
                    $configFiles = @(
                        Join-Path -Path $configDir -ChildPath "servers.json",
                        Join-Path -Path $configDir -ChildPath "api.config",
                        Join-Path -Path $configDir -ChildPath "lastpath.config",
                        Join-Path -Path $configDir -ChildPath "model.config",
                        Join-Path -Path $configDir -ChildPath "prompts.json"
                    )
                    foreach($file in $configFiles) { if (Test-Path $file) { Remove-Item -Path $file -Force; Write-Log -Message "Deleted: $(Split-Path $file -Leaf)" -Level 'DEBUG' } }
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

function Initialize-PromptConfiguration {
    Write-Log -Message "Entering Function: Initialize-PromptConfiguration" -Level 'DEBUG'
    $promptFile = $script:OptimusConfig.PromptFile
    if (Test-Path $promptFile) {
        Write-Log -Message "prompts.json already exists. Skipping creation." -Level 'DEBUG'
        return
    }

    Write-Log -Message "prompts.json not found. Creating and seeding with default prompts." -Level 'DEBUG'
    $defaultPrompts = @{
        prompts = @(
            @{
                name = "Default Tuning Prompt"
                type = "Tuning"
                description = "The standard, built-in prompt for performance tuning analysis."
                body = @"
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
    Before recommending a new index, you must first verify that no existing index can be slightly altered (e.g., by adding an `INCLUDE` column) to satisfy the query. Do not suggest an index if it is redundant or a near-duplicate of an existing one.
    * **Alter Existing Index:** If an existing index can be modified (e.g., adding an INCLUDE column) to better serve the query, provide the necessary `DROP` and `CREATE` DDL.
    * **Create New Index:** If no existing index is a suitable candidate for alteration, recommend a new, covering index. The execution plan's "missing index" suggestion should be treated as a starting point, not the final answer. You must critically evaluate this suggestion and improve upon it by adding necessary `INCLUDE` columns to create a true covering index or by optimizing the key column order. Provide the complete `CREATE INDEX` DDL.

Comment Formatting:
Every analysis comment block you add MUST use the following structure. The block starts with a general "Optimus Analysis" header. Inside, each distinct recommendation is numbered and contains the three required sections. This allows for multiple, independent suggestions for the same statement. Important: All text inside the comment block must be plain text. Do not use any Markdown formatting like bolding or backticks.

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


(Add more numbered recommendations as needed for the same statement)
*/

Final Output Rules:
- Your response MUST be the complete, original T-SQL script from start to finish. Do not omit any part of the original script for any reason.
- For statements that require improvement, insert your formatted analysis comment block directly above the statement.
- For statements that are already optimal, include the original T-SQL for that statement without any comment.
- Your entire response must be ONLY the T-SQL script text. Do not include any conversational text, greetings, or explanations outside of the T-SQL comments.
"@
            },
            @{
                name = "Default Code Review Prompt"
                type = "CodeReview"
                description = "The standard, built-in prompt for code style and best practices review."
                body = @"
You are an expert T-SQL code reviewer.

Your Core Mandate:
Your ONLY task is to return the complete, original T-SQL script provided below. You will add T-SQL comment blocks containing your review items directly above any code you identify for improvement. Think of this as a text transformation task: the original text is your input, and the identical text with your added comments is the only output.

Your Golden Rule:
You MUST NOT change the original T-SQL code. The final output must contain the original, unmodified T-SQL script with only your comments added.

Your Task:
Your goal is to perform a static code analysis on the provided T-SQL script. You will identify areas that deviate from established best practices. For any section of code that can be improved, add a single T-SQL block comment immediately above it. If a statement or section is already optimal, do not add a comment for it.

Code Review Categories:
Analyze the script against these standardized best practices:
1.  **Readability & Formatting:** Consistent casing, clear comments, logical code structure, and proper indentation.
2.  **Best Practices:** Correct use of `SET NOCOUNT ON`, schema-qualification for all database objects (e.g., dbo.MyTable), avoiding `SELECT *` in production code, using `EXISTS` instead of `IN` where appropriate.
3.  **Error Handling:** Presence and correct implementation of `TRY...CATCH` blocks for DML statements. Verification that transaction management (`BEGIN TRAN`, `COMMIT`, `ROLLBACK`) is handled correctly within the `TRY...CATCH` structure.
4.  **Maintainability:** Avoiding "magic numbers" or hard-coded strings that should be parameters, use of meaningful and consistent object and variable names.

Comment Formatting:
Every review comment block you add MUST use the following structure. The block starts with a "Optimus Code Review" header. Inside, each distinct review item is numbered. Important: All text inside the comment block must be plain text. Do not use any Markdown formatting like bolding or backticks.

/*
--- Optimus Code Review ---

[1] Review Item: (e.g., Missing SET NOCOUNT ON)
    - Finding: A brief explanation of the issue found in the code.
    - Recommendation: The suggested code change or addition to fix the issue.
    - Justification: An explanation of why the change aligns with T-SQL best practices (e.g., "Reduces unnecessary network traffic by stopping the 'rows affected' messages from being sent to the client.").

(Add more numbered review items as needed for the same code block)
*/

Final Output Rules:
- Your response MUST be the complete, original T-SQL script from start to finish. Do not omit any part of the original script.
- For code that requires improvement, insert your formatted review comment block directly above it.
- For code that is already optimal, include the original T-SQL for that section without any comment.
- Your entire response must be ONLY the T-SQL script text. Do not include any conversational text, greetings, or explanations outside of the T-SQL comments.
"@
            }
        )
    }

    try {
        # Use our robust manual serializer to ensure a valid JSON file is always created
        $jsonOutputString = ConvertTo-JsonManual -prompts $defaultPrompts.prompts
        $jsonOutputString | Out-File -FilePath $promptFile -Encoding UTF8
        Write-Log -Message "Successfully created and seeded prompts.json." -Level 'SUCCESS'
    } catch {
        Write-Log -Message "Failed to create prompts.json: $($_.Exception.Message)" -Level 'ERROR'
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
        $promptFile = Join-Path -Path $configDir -ChildPath "prompts.json"

        foreach($dir in @($configDir, $analysisBaseDir)){ if (-not (Test-Path -Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null } }
        if (-not (Test-Path -Path $serverFile)) { Set-Content -Path $serverFile -Value "[]" | Out-Null }
        
        $script:OptimusConfig = @{
            AnalysisBaseDir = $analysisBaseDir
            ServerFile      = $serverFile
            ApiKeyFile      = $apiKeyFile
            LastPathFile    = $lastPathFile
            ModelFile       = $modelFile
            PromptFile      = $promptFile
        }

        # After core config is set, initialize the prompts file.
        Initialize-PromptConfiguration

        Write-Log -Message "Configuration initialized successfully." -Level 'DEBUG'
        return $true
    }
    catch { Write-Log -Message "Could not initialize configuration: $($_.Exception.Message)" -Level 'ERROR'; return $false }
}

function Get-And-Set-ApiKey {
    Write-Log -Message "Entering Function: Get-And-Set-ApiKey" -Level 'DEBUG'
    Write-Log -Message "Checking for Gemini API Key..." -Level 'DEBUG'
    $apiKeyFile = $script:OptimusConfig.ApiKeyFile
    if (Test-Path -Path $apiKeyFile) {
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
    Write-Log -Message "Entering Function: Get-And-Set-Model" -Level 'DEBUG'
    $modelFile = $script:OptimusConfig.ModelFile

    if (Test-Path -Path $modelFile) {
        $modelName = Get-Content -Path $modelFile
        if (-not [string]::IsNullOrWhiteSpace($modelName)) {
            Write-Log -Message "Using configured model: '$modelName'" -Level 'DEBUG'
            return $modelName
        }
    }

    Write-Log -Message "`nPlease select the Gemini model to use for all future analyses:" -Level 'INFO'
    Write-Log -Message "This can be changed later using the -ResetConfiguration parameter." -Level 'INFO'
    Write-Log -Message "   [1] Gemini 1.5 Flash (Fastest, good for general use - Default)" -Level 'PROMPT'
    Write-Log -Message "   [2] Gemini 2.5 Flash (Next-gen speed and efficiency)" -Level 'PROMPT'
    Write-Log -Message "   [3] Gemini 2.5 Pro (Most powerful, for complex analysis)" -Level 'PROMPT'
    
    $modelChoice = $null
    while (-not $modelChoice) {
        Write-Log -Message "   Enter your choice: " -Level 'PROMPT' -NoNewLine
        $choice = Read-Host
        Write-Host ""
        Write-Log -Message "User Input: $choice" -Level 'DEBUG'
        switch ($choice) {
            '1' { $modelChoice = 'gemini-1.5-flash-latest' }
            '2' { $modelChoice = 'gemini-2.5-flash' }
            '3' { $modelChoice = 'gemini-2.5-pro' }
            default { Write-Log -Message "Invalid selection. Please enter 1, 2, or 3." -Level 'ERROR' }
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

function ConvertTo-JsonManual {
    param($prompts)

    # Use .NET StringBuilder for robust, efficient string construction
    $sb = New-Object System.Text.StringBuilder
    $sb.AppendLine("{") | Out-Null
    $sb.AppendLine("  `"prompts`": [") | Out-Null

    for ($i = 0; $i -lt $prompts.Count; $i++) {
        $prompt = $prompts[$i]

        # Correctly escape all string properties for JSON compliance.
        # The order of replacement is critical: backslashes must be escaped first.
        $name = $prompt.name -replace '\\', '\\' -replace '"', '\"' -replace "`r", '\r' -replace "`n", '\n' -replace "`t", '\t' -replace "`b", '\b'
        $type = $prompt.type -replace '\\', '\\' -replace '"', '\"' -replace "`r", '\r' -replace "`n", '\n' -replace "`t", '\t' -replace "`b", '\b'
        $desc = $prompt.description -replace '\\', '\\' -replace '"', '\"' -replace "`r", '\r' -replace "`n", '\n' -replace "`t", '\t' -replace "`b", '\b'
        $body = $prompt.body -replace '\\', '\\' -replace '"', '\"' -replace "`r", '\r' -replace "`n", '\n' -replace "`t", '\t' -replace "`b", '\b'

        # Append the JSON structure for a single prompt
        $sb.AppendLine("    {") | Out-Null
        $sb.AppendLine("      `"name`": `"$name`",") | Out-Null
        $sb.AppendLine("      `"type`": `"$type`",") | Out-Null
        $sb.AppendLine("      `"description`": `"$desc`",") | Out-Null
        $sb.AppendLine("      `"body`": `"$body`"") | Out-Null
        $sb.Append("    }") | Out-Null

        # Add a comma if it's not the last item in the array
        if ($i -lt $prompts.Count - 1) {
            $sb.AppendLine(",") | Out-Null
        } else {
            $sb.AppendLine() | Out-Null
        }
    }

    $sb.AppendLine("  ]") | Out-Null
    $sb.Append("}") | Out-Null

    return $sb.ToString()
}

function Import-UserPrompt {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )
    Write-Log -Message "Entering Function: Import-UserPrompt" -Level 'DEBUG'
    Write-Log -Message "`n--- Optimus Prompt Import Utility ---" -Level 'SUCCESS'

    try {
        $fileContent = Get-Content -Path $FilePath -Raw
        Write-Log -Message "Successfully read file: $(Split-Path $FilePath -Leaf)" -Level 'INFO'
    } catch {
        Write-Log -Message "Failed to read the file at '$FilePath'. Error: $($_.Exception.Message)" -Level 'ERROR'
        return
    }

    # Gather Metadata
    $defaultName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    Write-Log -Message "Please provide a name for this new prompt (press Enter to use '$defaultName'): " -Level 'PROMPT' -NoNewLine
    $promptName = Read-Host
    if ([string]::IsNullOrWhiteSpace($promptName)) { $promptName = $defaultName }
    Write-Host ""
    Write-Log -Message "User Input: $promptName" -Level 'DEBUG'

    $promptType = $null
    while (-not $promptType) {
        Write-Log -Message "What type of prompt is this?" -Level 'PROMPT'
        Write-Log -Message "  [1] Tuning" -Level 'PROMPT'
        Write-Log -Message "  [2] CodeReview" -Level 'PROMPT'
        Write-Log -Message "Enter your choice: " -Level 'PROMPT' -NoNewLine
        $choice = Read-Host
        Write-Host ""
        Write-Log -Message "User Input: $choice" -Level 'DEBUG'
        switch ($choice) {
            '1' { $promptType = 'Tuning' }
            '2' { $promptType = 'CodeReview' }
            default { Write-Log -Message "Invalid selection. Please enter 1 or 2." -Level 'ERROR' }
        }
    }

    Write-Log -Message "Please provide a brief description for this prompt: " -Level 'PROMPT' -NoNewLine
    $promptDescription = Read-Host
    Write-Host ""
    Write-Log -Message "User Input: $promptDescription" -Level 'DEBUG'

    # Confirmation
    Write-Log -Message "`n--- Confirmation ---" -Level 'WARN'
    Write-Log -Message "A new prompt will be added with the following details:" -Level 'INFO'
    Write-Log -Message "  Name:        $promptName" -Level 'PROMPT'
    Write-Log -Message "  Type:        $promptType" -Level 'PROMPT'
    Write-Log -Message "  Description: $promptDescription" -Level 'PROMPT'
    
    Write-Log -Message "`nAre you sure you want to add this prompt? (Y/N): " -Level 'PROMPT' -NoNewLine
    $confirm = Read-Host
    Write-Host ""
    Write-Log -Message "User Input: $confirm" -Level 'DEBUG'

    if ($confirm -notmatch '^[Yy]$') {
        Write-Log -Message "Import cancelled by user." -Level 'WARN'
        return
    }

    # Add to JSON
    try {
        $promptFile = $script:OptimusConfig.PromptFile
        Write-Log -Message "Attempting to read prompts configuration from: $promptFile" -Level 'DEBUG'
        $promptList = [System.Collections.Generic.List[pscustomobject]](Get-Content -Path $promptFile -Raw | ConvertFrom-Json).prompts
        Write-Log -Message "Current prompt count: $($promptList.Count)" -Level 'DEBUG'
        
        $newPrompt = [pscustomobject]@{
            name        = $promptName
            type        = $promptType
            description = $promptDescription
            body        = $fileContent
        }
        Write-Log -Message "New prompt object created. Adding to configuration." -Level 'DEBUG'

        # Add the new prompt to the .NET List
        $promptList.Add($newPrompt)

        # Rebuild the final configuration object using the stable list
        $finalConfig = @{
            prompts = $promptList
        }
        Write-Log -Message "New prompt count: $($finalConfig.prompts.Count)" -Level 'DEBUG'
        
        Write-Log -Message "Attempting to convert the final configuration object to a JSON string using manual serializer." -Level 'DEBUG'
        $jsonOutputString = ConvertTo-JsonManual -prompts $finalConfig.prompts
        Write-Log -Message "Successfully converted object to JSON string. Length: $($jsonOutputString.Length)" -Level 'DEBUG'
        
        Write-Log -Message "Attempting to write JSON string to file: $promptFile" -Level 'DEBUG'
        $jsonOutputString | Out-File -FilePath $promptFile -Encoding UTF8
        Write-Log -Message "Success! The new prompt '$promptName' has been added to your configuration." -Level 'SUCCESS'
    } catch {
        Write-Log -Message "Failed to update prompts.json. Error: $($_.Exception.Message)" -Level 'ERROR'
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
        # The URL for the raw script file on GitHub
        $repoUrl = "https://raw.githubusercontent.com/BrennanWebb/powershell/main/Production/Optimus.ps1"
        Write-Log -Message "Checking for new version at: $repoUrl" -Level 'DEBUG'

        # Download the latest script content as a string
        $webContent = Invoke-WebRequest -Uri $repoUrl -UseBasicParsing -TimeoutSec 10 | Select-Object -ExpandProperty Content

        # Use regex to find the version number in the script's header
        if ($webContent -match "Version:\s*([^\s]+)") {
            $latestVersionStr = $matches[1]
            Write-Log -Message "Latest version found online: '$latestVersionStr'" -Level 'DEBUG'
            
            # Sanitize versions for comparison by removing suffixes like '-preview'
            $cleanCurrent = ($CurrentVersion -split '-')[0]
            $cleanLatest = ($latestVersionStr -split '-')[0]

            # Compare the versions
            if ([System.Version]$cleanLatest -gt [System.Version]$cleanCurrent) {
                Write-Log -Message "A new version of Optimus is available! (Current: v$CurrentVersion, Latest: v$latestVersionStr)" -Level 'WARN'
                Write-Log -Message "You can download it from: https://github.com/BrennanWebb/powershell/blob/main/Production/Optimus.ps1" -Level 'WARN'
            } else {
                Write-Log -Message "Optimus is up to date." -Level 'DEBUG'
            }
        }
    }
    catch {
        # Fail silently if the check doesn't work. This is a non-essential feature.
        Write-Log -Message "Could not check for a new version. This can happen if GitHub is unreachable or there is no internet connection." -Level 'DEBUG'
        Write-Log -Message "Version check error: $($_.Exception.Message)" -Level 'DEBUG'
    }
}

function Test-PowerShellVersion {
    Write-Log -Message "Entering Function: Test-PowerShellVersion" -Level 'DEBUG'
    $currentVersion = $PSVersionTable.PSVersion
    Write-Log -Message "Checking current version: '$($currentVersion)' against required version '5.1'" -Level 'DEBUG'
    if ($currentVersion.Major -lt 5 -or ($currentVersion.Major -eq 5 -and $currentVersion.Minor -lt 1)) {
        Write-Log -Message "This script requires PowerShell version 5.1 or higher. You are running version $currentVersion." -Level 'ERROR'
        return $false
    }
    Write-Log -Message "PowerShell version $currentVersion is compatible." -Level 'DEBUG'
    return $true
}

function Test-WindowsEnvironment {
    Write-Log -Message "Entering Function: Test-WindowsEnvironment" -Level 'DEBUG'
    Write-Log -Message "Value of `$env:OS: $($env:OS)" -Level 'DEBUG'
    Write-Log -Message "Value of `$PSVersionTable.PSEdition: $($PSVersionTable.PSEdition)" -Level 'DEBUG'

    if ($env:OS -ne 'Windows_NT') {
        Write-Log -Message "This script requires a Windows operating system." -Level 'ERROR'
        return $false
    }

    if ($PSVersionTable.PSEdition -ne 'Desktop') {
        Write-Log -Message "You are running PowerShell $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition) edition)." -Level 'WARN'
        Write-Log -Message "The graphical file picker will not be available. Please use the -SQLFile parameter instead." -Level 'WARN'
    } else {
        Write-Log -Message "Windows environment with PowerShell $($PSVersionTable.PSEdition) edition confirmed." -Level 'DEBUG'
    }
    return $true
}

function Test-InternetConnection {
    param([string]$HostName = "googleapis.com")
    Write-Log -Message "Entering Function: Test-InternetConnection" -Level 'DEBUG'
    Write-Log -Message "Checking Internet connectivity..." -Level 'DEBUG'
    try {
        # Use .NET TcpClient for a completely silent connection test, eliminating console flicker.
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        $asyncResult = $tcpClient.BeginConnect($HostName, 443, $null, $null)
        # Wait for up to 3 seconds for the connection to succeed.
        $success = $asyncResult.AsyncWaitHandle.WaitOne(3000, $true)

        if ($success) {
            $tcpClient.EndConnect($asyncResult)
            $tcpClient.Close()
            Write-Log -Message "Internet connection to '$($HostName)' successful." -Level 'DEBUG'
            return $true
        } else {
            $tcpClient.Close()
            throw "Connection to $($HostName):443 timed out."
        }
    }
    catch {
        Write-Log -Message "Could not establish an internet connection to '$($HostName)'. The script will likely fail when contacting the Gemini API." -Level 'WARN'
        Write-Log -Message "Internet check failed with error: $($_.Exception.Message)" -Level 'DEBUG'
        Write-Log -Message "Would you like to continue anyway? (Y/N): " -Level 'PROMPT' -NoNewLine
        $choice = Read-Host
        Write-Host ""
        Write-Log -Message "User Input: $choice" -Level 'DEBUG'
        if ($choice -match '^[Yy]$') {
            Write-Log -Message "Continuing without a verified internet connection..." -Level 'WARN'
            return $true
        } else {
            return $false
        }
    }
}

function Test-SqlServerModule {
    Write-Log -Message "Entering Function: Test-SqlServerModule" -Level 'DEBUG'
    Write-Log -Message "Checking for 'SqlServer' PowerShell module..." -Level 'DEBUG'
    if (Get-Module -Name SqlServer -ListAvailable) {
        try {
            Import-Module SqlServer -ErrorAction Stop
            Write-Log -Message "'SqlServer' module imported." -Level 'DEBUG'
            return $true
        }
        catch { Write-Log -Message "Failed to import 'SqlServer' module: $($_.Exception.Message)" -Level 'ERROR'; return $false }
    } else {
        Write-Log -Message "The 'SqlServer' module is not installed." -Level 'WARN'
        Write-Log -Message "Would you like to attempt to install it now for the current user? (Y/N): " -Level 'PROMPT' -NoNewLine
        $choice = Read-Host
        Write-Host ""
        Write-Log -Message "User Input: $choice" -Level 'DEBUG'
        if ($choice -match '^[Yy]$') {
            Write-Log -Message "Installing 'SqlServer' module. This may take a moment..." -Level 'INFO'
            try {
                Install-Module -Name SqlServer -Scope CurrentUser -AllowClobber -Force -ErrorAction Stop
                Write-Log -Message "Module installed successfully. Importing..." -Level 'DEBUG'
                Import-Module SqlServer -ErrorAction Stop
                Write-Log -Message "'SqlServer' module is now ready." -Level 'SUCCESS'
                return $true
            } catch {
                Write-Log -Message "Failed to install or import the 'SqlServer' module. Please install it manually using: Install-Module -Name SqlServer -Scope CurrentUser" -Level 'ERROR'
                Write-Log -Message "Error details: $($_.Exception.Message)" -Level 'ERROR'
                return $false
            }
        } else {
             Write-Log -Message "The 'SqlServer' module is required to continue. Exiting." -Level 'ERROR'
             return $false
        }
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
        $result = Invoke-Sqlcmd -ServerInstance $ServerInstance -Query "SELECT @@VERSION" -TrustServerCertificate
        return $result.Item(0)
    } catch {
        Write-Log -Message "Could not retrieve SQL Server version details." -Level 'WARN'
        return "Unknown"
    }
}

function Select-SqlServer {
    Write-Log -Message "Entering Function: Select-SqlServer" -Level 'DEBUG'
    Write-Log -Message "`nPlease select a SQL Server to use:" -Level 'INFO'
    [array]$servers = Get-Content -Path $script:OptimusConfig.ServerFile | ConvertFrom-Json
    if ($servers.Count -gt 0) { for ($i = 0; $i -lt $servers.Count; $i++) { Write-Log -Message "   [$($i+1)] $($servers[$i])" -Level 'PROMPT' } }
    Write-Log -Message "   [A] Add a new server" -Level 'PROMPT'; Write-Log -Message "   [Q] Quit" -Level 'PROMPT'
    while ($true) {
        Write-Log -Message "   Enter your choice: " -Level 'PROMPT' -NoNewLine
        $choice = Read-Host
        Write-Host ""
        Write-Log -Message "User Input: $choice" -Level 'DEBUG'
        if ($choice -imatch 'Q') { return $null }
        if ($choice -imatch 'A') {
            Write-Log -Message "   Enter the new SQL server name or IP: " -Level 'PROMPT' -NoNewLine
            $newServer = Read-Host
            Write-Host ""
            Write-Log -Message "User Input: $newServer" -Level 'DEBUG'
            if ([string]::IsNullOrWhiteSpace($newServer)) { Write-Log -Message "Server name cannot be empty." -Level 'ERROR'; continue }
            if (Test-SqlServerConnection -ServerInstance $newServer) {
                $servers += $newServer; ($servers | Sort-Object -Unique) | ConvertTo-Json -Depth 5 | Set-Content -Path $script:OptimusConfig.ServerFile
                Write-Log -Message "'$newServer' has been added." -Level 'SUCCESS'; return $newServer
            }
            continue
        }
        if ($choice -match '^\d+$' -and [int]$choice -gt 0 -and [int]$choice -le $servers.Count) {
            $selectedServer = $servers[[int]$choice - 1]
            if (Test-SqlServerConnection -ServerInstance $selectedServer) { return $selectedServer }
        } else { Write-Log -Message "Invalid choice." -Level 'ERROR' }
    }
}

function Show-FilePicker {
    Write-Log -Message "Entering Function: Show-FilePicker" -Level 'DEBUG'
    
    $initialDir = [System.Environment]::getFolderPath('MyDocuments')
    $lastPathFile = $script:OptimusConfig.LastPathFile
    
    if (Test-Path -Path $lastPathFile) {
        $lastPath = Get-Content -Path $lastPathFile
        if ((-not [string]::IsNullOrWhiteSpace($lastPath)) -and (Test-Path -Path $lastPath -PathType Container)) {
            $initialDir = $lastPath
            Write-Log -Message "Setting initial file dialog directory to last used path: $initialDir" -Level 'DEBUG'
        }
    }

    try {
        Add-Type -AssemblyName System.Windows.Forms
        $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $fileDialog.Title = "Select one or more SQL Files for Analysis"
        $fileDialog.InitialDirectory = $initialDir
        $fileDialog.Filter = "SQL Files (*.sql)|*.sql|All files (*.*)|*.*"
        $fileDialog.Multiselect = $true
        if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { 
            # Save the directory of the selected file(s) for next time
            try {
                $directory = [System.IO.Path]::GetDirectoryName($fileDialog.FileNames[0])
                Set-Content -Path $script:OptimusConfig.LastPathFile -Value $directory
                Write-Log -Message "Saved last used directory: $directory" -Level 'DEBUG'
            } catch {
                Write-Log -Message "Could not save the last used directory path." -Level 'WARN'
            }
            return $fileDialog.FileNames 
        }
    }
    catch { Write-Log -Message "Could not display graphical file picker: $($_.Exception.Message)" -Level 'WARN' }
    return $null
}

function Get-AnalysisInputs {
    Write-Log -Message "Entering Function: Get-AnalysisInputs" -Level 'DEBUG'
    Write-Log -Message "Parameter Set Name: $($PSCmdlet.ParameterSetName)" -Level 'DEBUG'
    
    $inputObjects = [System.collections.generic.List[object]]::new()

    switch ($PSCmdlet.ParameterSetName) {
        'Adhoc' {
            if (-not [string]::IsNullOrWhiteSpace($AdhocSQL)) {
                $baseName = "AdhocQuery_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
                $inputObjects.Add([pscustomobject]@{
                    SqlText  = $AdhocSQL
                    BaseName = $baseName
                })
                Write-Log -Message "Received Ad-hoc SQL for analysis." -Level 'SUCCESS'
            } else {
                Write-Log -Message "The -AdhocSQL parameter was used but contained no query." -Level 'WARN'
                return $null
            }
        }
        'Folder' {
            if (-not (Test-Path -Path $FolderPath -PathType Container)) {
                Write-Log -Message "The path provided for -FolderPath is not a valid directory: '$FolderPath'." -Level 'ERROR'
                return $null
            }
            $filesToAnalyze = (Get-ChildItem -Path $FolderPath -Filter *.sql).FullName
            if ($filesToAnalyze.Count -eq 0) {
                Write-Log -Message "No .sql files were found in the specified folder: '$FolderPath'." -Level 'WARN'
                return $null
            }
            Write-Log -Message "Found $($filesToAnalyze.Count) file(s) in folder '$FolderPath' for analysis." -Level 'SUCCESS'
            foreach ($file in $filesToAnalyze) {
                $inputObjects.Add([pscustomobject]@{
                    SqlText  = Get-Content -Path $file -Raw
                    BaseName = [System.IO.Path]::GetFileNameWithoutExtension($file)
                })
            }
        }
        'Files' {
            [string[]]$validFiles = @()
            foreach ($file in $SQLFile) {
                if ((Test-Path -Path $file -PathType Leaf) -and $file -like '*.sql') {
                    $validFiles += $file
                } else {
                    Write-Log -Message "Parameter invalid, file not found or not a .sql file: '$file'. Skipping." -Level 'WARN'
                }
            }
            if ($validFiles.Count -eq 0) {
                 Write-Log -Message "No valid .sql files were provided via the -SQLFile parameter." -Level 'WARN'
                 return $null
            }
            Write-Log -Message "Successfully targeted $($validFiles.Count) file(s) for analysis." -Level 'SUCCESS'
            foreach ($file in $validFiles) {
                $inputObjects.Add([pscustomobject]@{
                    SqlText  = Get-Content -Path $file -Raw
                    BaseName = [System.IO.Path]::GetFileNameWithoutExtension($file)
                })
            }
        }
        default { # Interactive Mode
            while ($inputObjects.Count -eq 0) {
                Write-Log -Message "`nPlease select one or more .sql files to analyze..." -Level 'INFO'
                $selectedFiles = Show-FilePicker
                
                if ($null -ne $selectedFiles -and $selectedFiles.Count -gt 0) {
                     Write-Log -Message "Successfully selected $($selectedFiles.Count) file(s) for analysis." -Level 'SUCCESS'
                     foreach ($file in $selectedFiles) {
                        $inputObjects.Add([pscustomobject]@{
                            SqlText  = Get-Content -Path $file -Raw
                            BaseName = [System.IO.Path]::GetFileNameWithoutExtension($file)
                        })
                    }
                } else {
                    Write-Log -Message "   File selection cancelled. Try again? (Y/N): " -Level 'PROMPT' -NoNewLine
                    $retry = Read-Host
                    Write-Host ""
                    Write-Log -Message "User Input: $retry" -Level 'DEBUG'
                    if ($retry -notmatch '^[Yy]$') { return $null }
                }
            }
        }
    }

    return $inputObjects
}

function Select-AIPrompt {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet('Tuning', 'CodeReview')]
        [string]$AnalysisType
    )
    Write-Log -Message "Entering Function: Select-AIPrompt" -Level 'DEBUG'
    $promptFile = $script:OptimusConfig.PromptFile
    if (-not (Test-Path $promptFile)) {
        Write-Log -Message "CRITICAL: The prompts.json file is missing at '$promptFile'. Please run with -ResetConfiguration to regenerate it." -Level 'ERROR'
        return $null
    }

    $allPrompts = (Get-Content -Path $promptFile -Raw | ConvertFrom-Json).prompts
    $availablePrompts = @($allPrompts | Where-Object { $_.type -eq $AnalysisType })

    if ($availablePrompts.Count -eq 0) {
        Write-Log -Message "Could not find any prompts of type '$AnalysisType' in '$promptFile'." -Level 'ERROR'
        return $null
    }

    # Automated selection via -PromptName parameter
    if (-not [string]::IsNullOrWhiteSpace($PromptName)) {
        Write-Log -Message "PromptName parameter used. Searching for prompt: '$PromptName' of type '$AnalysisType'" -Level 'DEBUG'
        $selectedPrompt = $availablePrompts | Where-Object { $_.name -eq $PromptName }
        if ($selectedPrompt) {
            Write-Log -Message "Found matching prompt: '$($selectedPrompt.Name)'." -Level 'SUCCESS'
            return $selectedPrompt.body
        } else {
            Write-Log -Message "A prompt with the name '$PromptName' and type '$AnalysisType' was not found in prompts.json." -Level 'ERROR'
            return $null
        }
    }

    # For non-interactive runs without a specified prompt name, select the default and exit.
    if ($PSCmdlet.ParameterSetName -ne 'Interactive' -and -not $DebugMode.IsPresent) {
        Write-Log -Message "Non-interactive mode detected. Automatically selecting the default prompt." -Level 'DEBUG'
        return $availablePrompts[0].body
    }

    # Interactive selection
    Write-Log -Message "`nPlease select a prompt to use for this $AnalysisType batch:" -Level 'INFO'
    for ($i = 0; $i -lt $availablePrompts.Count; $i++) {
        Write-Log -Message "   [$($i+1)] $($availablePrompts[$i].name)" -Level 'PROMPT'
        Write-Log -Message "        $($availablePrompts[$i].description)" -Level 'PROMPT'
    }

    while ($true) {
        Write-Log -Message "   Enter your choice (default is 1): " -Level 'PROMPT' -NoNewLine
        $choice = Read-Host
        Write-Host ""
        Write-Log -Message "User Input: $choice" -Level 'DEBUG'
        if ([string]::IsNullOrWhiteSpace($choice)) { $choice = 1 }
        if ($choice -match '^\d+$' -and [int]$choice -ge 1 -and [int]$choice -le $availablePrompts.Count) {
            $selectedPrompt = $availablePrompts[[int]$choice - 1]
            Write-Log -Message "User selected prompt: '$($selectedPrompt.name)'" -Level 'DEBUG'
            return $selectedPrompt.body
        } else {
            Write-Log -Message "Invalid selection. Please enter a number between 1 and $($availablePrompts.Count)." -Level 'ERROR'
        }
    }
}
#endregion

#region Data Parsing, Collection, and AI Analysis

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
    
    # Log the exact query text being sent to the log file for debugging and auditing.
    Write-Log -Message "Cleaned T-SQL for execution plan generation:`n$cleanQueryText" -Level 'DEBUG'

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
        
        if ($planFragments.Count -eq 0) {
            Write-Log -Message "Could not find a valid execution plan string in the results from SQL Server." -Level 'ERROR'
            return $null
        }

        # Create the wrapped plan for internal parsing.
        $masterPlanXml = "<MasterShowPlan>" + ($planFragments -join '') + "</MasterShowPlan>"

        # Create a valid, merged .sqlplan file by properly merging XML nodes.
        $rawPlanContent = ''
        try {
            # Create a valid, empty master plan structure.
            $masterShowPlan = [xml]'<ShowPlanXML Version="1.539" Build="16.0.1110.1" xmlns="http://schemas.microsoft.com/sqlserver/2004/07/showplan"><BatchSequence><Batch></Batch></BatchSequence></ShowPlanXML>'
            
            # Use a namespace manager for reliable node selection, which is critical for handling the default namespace.
            $nsm = New-Object System.Xml.XmlNamespaceManager($masterShowPlan.NameTable)
            $nsm.AddNamespace("sp", "http://schemas.microsoft.com/sqlserver/2004/07/showplan")

            # Robustly select the target <Batch> node.
            $batchNode = $masterShowPlan.SelectSingleNode("//sp:Batch", $nsm)

            # Loop through each fragment and merge its statement contents.
            foreach ($fragment in $planFragments) {
                $planXml = [xml]$fragment
                $statementNodes = $planXml.ShowPlanXML.BatchSequence.Batch.ChildNodes
                foreach ($node in $statementNodes) {
                    # Import and append each statement into the single master document.
                    $importedNode = $masterShowPlan.ImportNode($node, $true)
                    $batchNode.AppendChild($importedNode) | Out-Null
                }
            }
            # Save the final, valid XML.
            $rawPlanContent = $masterShowPlan.OuterXml
        } catch {
             Write-Log -Message "Could not merge XML fragments into a valid .sqlplan file. Error: $($_.Exception.Message)" -Level 'WARN'
             # Fallback to the potentially invalid raw content if merging fails.
             $rawPlanContent = $planFragments -join ''
        }
        
        try {
            [xml]$masterPlanXml | Out-Null
            Write-Log -Message "Successfully generated and validated master execution plan for all statements." -Level 'SUCCESS'
            
            # Return an object containing both the raw plan for the .sqlplan file
            # and the wrapped XML for internal script processing.
            return [pscustomobject]@{
                MasterPlanXml    = $masterPlanXml
                RawPlanContent = $rawPlanContent
            }
        } catch {
            Write-Log -Message "The combined execution plan string is not valid XML. Error: $($_.Exception.Message)" -Level 'ERROR'
            return $null
        }
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
                $fullName = "$db.$schema.$table".Replace('[','').Replace(']','')
                $uniqueObjectNames.Add($fullName) | Out-Null
            }
        }

        $finalList = $uniqueObjectNames | Sort-Object
        
        if ($DebugMode -and $finalList.Count -gt 0) {
            $displayObjects = $finalList | ForEach-Object {
                $parts = $_.Split('.')
                [pscustomobject]@{
                    Database = $parts[0]
                    Schema   = $parts[1]
                    Name     = $parts[2]
                }
            }
            $tableOutput = $displayObjects | Format-Table -AutoSize | Out-String
            Write-Log -Message "The following unique user objects were found in the execution plan:`n$tableOutput" -Level 'DEBUG'
        }

        Write-Log -Message "Identified $($finalList.Count) unique objects for schema collection." -Level 'SUCCESS'
        Write-Log -Message "Returning from Get-ObjectsFromPlan." -Level 'DEBUG'
        return $finalList
    } catch {
        Write-Log -Message "Failed to parse objects from the execution plan: $($_.Exception.Message)" -Level 'ERROR'
        return @()
    }
}

function Get-ObjectSchema {
    param(
        [string]$ServerInstance,
        [string]$DatabaseName,
        [string]$SchemaName,
        [string]$ObjectName
    )
    Write-Log -Message "Now collecting schema for: $DatabaseName.$SchemaName.$ObjectName" -Level 'DEBUG'
    
    $fullObjectName = "[$SchemaName].[$ObjectName]"
    $schemaText = "--- Schema For Table: $SchemaName.$ObjectName ---`n"
    $columnResult = $null

    # Get Column Info - Primary Method
    try {
        $columnQuery = "SELECT name, system_type_name, max_length, [precision], scale, is_nullable FROM sys.dm_exec_describe_first_result_set(N'SELECT * FROM $fullObjectName', NULL, 0);"
        $columnResult = Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $DatabaseName -TrustServerCertificate -Query $columnQuery -ErrorAction Stop
    } catch {
        Write-Log -Message "Primary schema collection method failed for '$fullObjectName'. Attempting fallback." -Level 'DEBUG'
        # Fallback Method
        try {
            $fallbackQuery = @"
SELECT c.name, t.name AS system_type_name, c.max_length, c.precision, c.scale, c.is_nullable
FROM sys.columns c JOIN sys.types t ON c.user_type_id = t.user_type_id
WHERE c.object_id = OBJECT_ID(@FullObjectName) ORDER BY c.column_id;
"@
            $params = @{ FullObjectName = "$DatabaseName.$fullObjectName" }
            $columnResult = Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $DatabaseName -TrustServerCertificate -Query $fallbackQuery -Variable $params -ErrorAction Stop
        } catch {
            Write-Log -Message "Could not get COLUMN schema for '$fullObjectName' in db '$DatabaseName' using any method. Error: $($_.Exception.Message)" -Level 'WARN'
        }
    }

    if ($columnResult) {
        $schemaText += "COLUMNS:`n"
        foreach($col in $columnResult) {
            $isNullable = if ($col.is_nullable) { 'YES' } else { 'NO' }
            $schemaText += "name: $($col.name), type: $($col.system_type_name), length: $($col.max_length), nullable: $isNullable`n"
        }
    }

    # Get Index Info
    try {
        $indexQuery = @"
SELECT i.name AS IndexName, i.type_desc AS IndexType,
STUFF((SELECT ', ' + c.name FROM sys.index_columns ic JOIN sys.columns c ON ic.object_id = c.object_id AND ic.column_id = c.column_id WHERE ic.object_id = i.object_id AND ic.index_id = i.index_id AND ic.is_included_column = 0 ORDER BY ic.key_ordinal FOR XML PATH('')), 1, 2, '') AS KeyColumns,
STUFF((SELECT ', ' + c.name FROM sys.index_columns ic JOIN sys.columns c ON ic.object_id = c.object_id AND ic.column_id = c.column_id WHERE ic.object_id = i.object_id AND ic.index_id = i.index_id AND ic.is_included_column = 1 ORDER BY ic.key_ordinal FOR XML PATH('')), 1, 2, '') AS IncludedColumns
FROM sys.indexes i WHERE i.object_id = OBJECT_ID('$fullObjectName');
"@
        $indexResult = Invoke-Sqlcmd -ServerInstance $ServerInstance -Database $DatabaseName -TrustServerCertificate -Query $indexQuery -ErrorAction Stop
        if ($indexResult -and $indexResult.Count -gt 0) {
            $schemaText += "`nINDEXES:`n"
            foreach($idx in $indexResult) {
                $idxLine = "IndexName: $($idx.IndexName), Type: $($idx.IndexType), KeyColumns: $($idx.KeyColumns)"
                if (-not [string]::IsNullOrWhiteSpace($idx.IncludedColumns)) { $idxLine += ", IncludedColumns: $($idx.IncludedColumns)" }
                $schemaText += $idxLine + "`n"
            }
        }
    } catch {
        Write-Log -Message "Could not get INDEX information for '$fullObjectName'." -Level 'WARN'
    }
    
    return $schemaText + "`n"
}

function Invoke-GeminiAnalysis {
    param(
        [Parameter(Mandatory=$true)] [string]$ModelName,
        [Parameter(Mandatory=$true)] [securestring]$ApiKey, 
        [Parameter(Mandatory=$true)] [string]$PromptBody,
        [Parameter(Mandatory=$true)] [string]$FullSqlText, 
        [Parameter(Mandatory=$true)] [string]$ConsolidatedSchema, 
        [Parameter(Mandatory=$true)] [string]$MasterPlanXml,
        [Parameter(Mandatory=$true)] [string]$SqlServerVersion
    )
    Write-Log -Message "Entering Function: Invoke-GeminiAnalysis" -Level 'DEBUG'
    Write-Log -Message "Sending full script to Gemini for performance tuning analysis..." -Level 'INFO'

    $plainTextApiKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ApiKey))
    $uri = "https://generativelanguage.googleapis.com/v1beta/models/$($ModelName):generateContent?key=$plainTextApiKey"

    $prompt = @"
$PromptBody

--- SQL SERVER VERSION ---
$SqlServerVersion

--- FULL T-SQL SCRIPT ---
$FullSqlText

--- CONSOLIDATED OBJECT SCHEMAS AND DEFINITIONS ---
$ConsolidatedSchema

--- MASTER EXECUTION PLAN ---
$MasterPlanXml
"@

    $promptPath = Join-Path -Path $script:AnalysisPath -ChildPath "_FinalAIPrompt_Tuning.txt"
    try { $prompt | Set-Content -Path $promptPath -Encoding UTF8; Write-Log -Message "Final AI prompt saved for review at: $promptPath" -Level 'DEBUG' } catch { Write-Log -Message "Could not save final AI prompt file." -Level 'WARN' }

    $bodyObject = @{ contents = @( @{ parts = @( @{ text = $prompt } ) } ) }
    $bodyJson = $bodyObject | ConvertTo-Json -Depth 10

    try {
        $response = Invoke-RestMethod -Uri $uri -Method Post -Body $bodyJson -ContentType 'application/json' -ErrorAction Stop
        $rawAiResponse = $response.candidates[0].content.parts[0].text
        Write-Log -Message "Successfully received raw response from Gemini API." -Level 'DEBUG'
        
        $cleanedScript = $rawAiResponse -replace '(?i)^```sql\s*','' -replace '```\s*$',''
        Write-Log -Message "Cleaned response received from AI." -Level 'DEBUG'
        
        Write-Log -Message "AI performance analysis complete." -Level 'SUCCESS'
        return $cleanedScript
    } catch {
        Write-Log -Message "Failed to get response from Gemini API." -Level 'ERROR'
        $errorDetails = $_.Exception.Response.GetResponseStream()
        $streamReader = New-Object System.IO.StreamReader($errorDetails)
        $errorText = $streamReader.ReadToEnd()
        Write-Log -Message "API Error Details: $errorText" -Level 'ERROR'
        return $null
    }
}

function Invoke-GeminiCodeReview {
    param(
        [Parameter(Mandatory=$true)] [string]$ModelName,
        [Parameter(Mandatory=$true)] [securestring]$ApiKey,
        [Parameter(Mandatory=$true)] [string]$PromptBody,
        [Parameter(Mandatory=$true)] [string]$FullSqlText
    )
    Write-Log -Message "Entering Function: Invoke-GeminiCodeReview" -Level 'DEBUG'
    Write-Log -Message "Sending full script to Gemini for standardized code review..." -Level 'INFO'

    $plainTextApiKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ApiKey))
    $uri = "https://generativelanguage.googleapis.com/v1beta/models/$($ModelName):generateContent?key=$plainTextApiKey"

    $prompt = @"
$PromptBody

--- FULL T-SQL SCRIPT ---
$FullSqlText
"@

    $promptPath = Join-Path -Path $script:AnalysisPath -ChildPath "_FinalAIPrompt_CodeReview.txt"
    try { $prompt | Set-Content -Path $promptPath -Encoding UTF8; Write-Log -Message "Final AI prompt saved for review at: $promptPath" -Level 'DEBUG' } catch { Write-Log -Message "Could not save final AI prompt file." -Level 'WARN' }

    $bodyObject = @{ contents = @( @{ parts = @( @{ text = $prompt } ) } ) }
    $bodyJson = $bodyObject | ConvertTo-Json -Depth 10

    try {
        $response = Invoke-RestMethod -Uri $uri -Method Post -Body $bodyJson -ContentType 'application/json' -ErrorAction Stop
        $rawAiResponse = $response.candidates[0].content.parts[0].text
        Write-Log -Message "Successfully received raw response from Gemini API." -Level 'DEBUG'
        
        $cleanedScript = $rawAiResponse -replace '(?i)^```sql\s*','' -replace '```\s*$',''
        Write-Log -Message "Cleaned response received from AI." -Level 'DEBUG'
        
        Write-Log -Message "AI code review complete." -Level 'SUCCESS'
        return $cleanedScript
    } catch {
        Write-Log -Message "Failed to get response from Gemini API." -Level 'ERROR'
        $errorDetails = $_.Exception.Response.GetResponseStream()
        $streamReader = New-Object System.IO.StreamReader($errorDetails)
        $errorText = $streamReader.ReadToEnd()
        Write-Log -Message "API Error Details: $errorText" -Level 'ERROR'
        return $null
    }
}

function New-AnalysisSummary {
    param(
        [Parameter(Mandatory=$true)] [string]$TunedScript,
        [Parameter(Mandatory=$true)] [int]$TotalStatementCount,
        [Parameter(Mandatory=$true)] [string]$AnalysisPath
    )
    Write-Log -Message "Entering Function: New-AnalysisSummary" -Level 'DEBUG'
    
    try {
        $summaryContent = @"
--- Optimus Analysis Summary ---
Timestamp: $(Get-Date)

"@
        # Regex to find either Tuning or Code Review blocks
        $recommendationBlockRegex = '(?s)\/\*\s*--- Optimus (Analysis|Code Review) ---(.*?)\*\/';
        $recommendationBlocks = [regex]::Matches($TunedScript, $recommendationBlockRegex)
        
        # Regex to count individual recommendations or review items within a block
        $individualRecommendationRegex = '\[\d+\]\s*(Recommendation|Review Item)'
        $totalRecommendations = 0
        foreach ($block in $recommendationBlocks) {
            $totalRecommendations += ([regex]::Matches($block.Value, $individualRecommendationRegex)).Count
        }
        
        $summaryContent += "Total Statements Analyzed (Approximation): $TotalStatementCount`n"
        $summaryContent += "Code Blocks with Recommendations/Reviews: $($recommendationBlocks.Count)`n"
        $summaryContent += "Total Individual Recommendations/Reviews: $totalRecommendations`n`n"

        if ($recommendationBlocks.Count -gt 0) {
            $summaryContent += "--- Summary of Findings ---`n"
            $problemRegex = '(Problem|Finding):(.*?)(?=\s*-\s*(Recommended Code|Recommendation|Justification):|\s*$)';
            
            $findingIndex = 1
            foreach ($block in $recommendationBlocks) {
                $problemMatches = [regex]::Matches($block.Value, $problemRegex)
                foreach ($problem in $problemMatches) {
                    $problemText = $problem.Groups[2].Value.Trim()
                    $summaryContent += "$($findingIndex). $problemText`n"
                    $findingIndex++
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

# --- Analysis Workflow Handlers ---

function Start-TuningAnalysis {
    Write-Log -Message "Starting new batch in Performance Tuning mode." -Level 'INFO'

    # Server Selection Logic
    $selectedServer = $null
    if (-not [string]::IsNullOrWhiteSpace($ServerName)) {
        Write-Log -Message "ServerName parameter provided, attempting to connect to '$ServerName'." -Level 'INFO'
        if (Test-SqlServerConnection -ServerInstance $ServerName) {
            $selectedServer = $ServerName
            try {
                [array]$servers = Get-Content -Path $script:OptimusConfig.ServerFile | ConvertFrom-Json
                if ($selectedServer -notin $servers) {
                    $servers += $selectedServer
                    ($servers | Sort-Object -Unique) | ConvertTo-Json -Depth 5 | Set-Content -Path $script:OptimusConfig.ServerFile
                    Write-Log -Message "'$selectedServer' has been validated and saved to the configuration." -Level 'DEBUG'
                }
            } catch { Write-Log -Message "Could not save the provided server name to the configuration file." -Level 'WARN' }
        } else {
            Write-Log -Message "Connection test to the server '$ServerName' failed. Please check the server name and permissions." -Level 'ERROR'
            return # Exit the function
        }
    } else {
        $selectedServer = Select-SqlServer
    }

    if (-not $selectedServer) {
        Write-Log -Message "No valid SQL Server was selected or provided. Halting tuning analysis." -Level 'WARN'
        return 
    }
    
    # Select the AI prompt once for the entire batch.
    $selectedPromptBody = Select-AIPrompt -AnalysisType 'Tuning'
    if (-not $selectedPromptBody) {
        Write-Log -Message "No valid prompt was selected. Halting batch." -Level 'ERROR'
        return
    }

    [array]$analysisInputs = Get-AnalysisInputs
    if ($null -eq $analysisInputs -or $analysisInputs.Count -eq 0) {
        Write-Log -Message "No valid inputs were found for analysis." -Level 'WARN'
        return
    }

    $useActualPlanSwitch = $UseActualPlan.IsPresent
    if ($PSCmdlet.ParameterSetName -eq 'Interactive' -and -not $UseActualPlan.IsPresent) {
        Write-Log -Message "`nWhich execution plan would you like to generate for this batch?" -Level 'INFO'
        Write-Log -Message "   [1] Estimated (Default - Recommended, does not run the query)" -Level 'PROMPT'
        Write-Log -Message "   [2] Actual (Executes the query, use with caution on all files)" -Level 'PROMPT'
        Write-Log -Message "   Enter your choice: " -Level 'PROMPT' -NoNewLine
        $choice = Read-Host
        Write-Host ""
        Write-Log -Message "User Input: $choice" -Level 'DEBUG'
        if ($choice -eq '2') {
            Write-Log -Message "Proceeding with 'Actual Execution Plan'. This will EXECUTE every SQL script." -Level 'WARN'
            $useActualPlanSwitch = $true
        } else {
            Write-Log -Message "Defaulting to 'Estimated Execution Plan' for this batch." -Level 'INFO'
        }
    }

    $sanitizedModelName = $script:ChosenModel -replace '[.-]', '_'
    $modelSpecificPath = Join-Path -Path $script:OptimusConfig.AnalysisBaseDir -ChildPath $sanitizedModelName
    if (-not (Test-Path -Path $modelSpecificPath)) { New-Item -Path $modelSpecificPath -ItemType Directory -Force | Out-Null }
    $batchTimestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $batchFolderPath = Join-Path -Path $modelSpecificPath -ChildPath $batchTimestamp
    New-Item -Path $batchFolderPath -ItemType Directory -Force | Out-Null
    Write-Log -Message "`nCreated batch analysis folder: $batchFolderPath" -Level 'SUCCESS'

    foreach ($analysisItem in $analysisInputs) {
        try {
            $baseName = $analysisItem.BaseName
            $sqlQueryText = $analysisItem.SqlText
            Write-Log -Message "`n--- Starting Tuning Analysis for: $baseName ---" -Level 'SUCCESS'
            $script:AnalysisPath = Join-Path -Path $batchFolderPath -ChildPath $baseName
            New-Item -Path $script:AnalysisPath -ItemType Directory -Force | Out-Null
            $script:LogFilePath = Join-Path -Path $script:AnalysisPath -ChildPath "ExecutionLog.txt"
            "# Optimus v$($script:CurrentVersion) Tuning Log | File: $baseName | Started: $(Get-Date)" | Out-File -FilePath $script:LogFilePath -Encoding utf8
            Write-Log -Message "Created analysis directory: '$($script:AnalysisPath)'" -Level 'INFO'
            $sqlVersion = Get-SqlServerVersion -ServerInstance $selectedServer
            Write-Log -Message "Detected SQL Server Version: $sqlVersion" -Level 'DEBUG'
            
            $initialDbContext = ([regex]::Match($sqlQueryText, '(?im)^\s*USE\s+\[?([\w\d_]+)\]?')).Groups[1].Value
            if ([string]::IsNullOrWhiteSpace($initialDbContext)) { $initialDbContext = 'master' }
            
            $planObject = Get-MasterExecutionPlan -ServerInstance $selectedServer -DatabaseContext $initialDbContext -FullQueryText $sqlQueryText -IsActualPlan:$useActualPlanSwitch
            if (-not $planObject) { Write-Log -Message "Could not generate a master plan for $baseName. Skipping." -Level 'ERROR'; continue }
            
            $masterPlanXml = $planObject.MasterPlanXml
            $xmlPlanPath = Join-Path -Path $script:AnalysisPath -ChildPath "_MasterPlan.xml"
            try { $masterPlanXml | Set-Content -Path $xmlPlanPath -Encoding UTF8; Write-Log -Message "Master execution plan saved to .xml file." -Level 'DEBUG' } catch { Write-Log -Message "Could not save master plan .xml file." -Level 'WARN' }

            # Save the raw .sqlplan file with the script's base name
            $sqlPlanPath = Join-Path -Path $script:AnalysisPath -ChildPath "${baseName}.sqlplan"
            try {
                $planObject.RawPlanContent | Set-Content -Path $sqlPlanPath -Encoding UTF8
                Write-Log -Message "Raw execution plan saved to .sqlplan file." -Level 'DEBUG'
                # Check the new switch and open the file
                if ($OpenPlanFile.IsPresent) {
                    try {
                        Write-Log -Message "Opening .sqlplan file in default application..." -Level 'INFO'
                        Invoke-Item -Path $sqlPlanPath
                    } catch {
                        Write-Log -Message "Failed to open the .sqlplan file automatically: $($_.Exception.Message)" -Level 'WARN'
                    }
                }
            } catch {
                 Write-Log -Message "Could not save raw .sqlplan file." -Level 'WARN'
            }

            [xml]$masterPlan = $masterPlanXml
            $ns = New-Object System.Xml.XmlNamespaceManager($masterPlan.NameTable)
            $ns.AddNamespace("sql", "http://schemas.microsoft.com/sqlserver/2004/07/showplan")
            [string[]]$uniqueObjectNames = @(Get-ObjectsFromPlan -MasterPlan $masterPlan -NamespaceManager $ns)
            $statementNodes = $masterPlan.SelectNodes("//sql:StmtSimple", $ns)
            
            $consolidatedSchema = ""
            if ($null -ne $uniqueObjectNames -and $uniqueObjectNames.Count -gt 0) {
                Write-Log -Message "Starting schema collection for all objects..." -Level 'INFO'
                $objectsByDb = $uniqueObjectNames | Group-Object { ($_ -split '\.')[0] }
                foreach ($dbGroup in $objectsByDb) {
                    $dbName = $dbGroup.Name
                    if ($dbName -eq 'mssqlsystemresource') { Write-Log -Message "Skipping schema collection for internal database: 'mssqlsystemresource'." -Level 'DEBUG'; continue }
                    Write-Log -Message "Querying database '$dbName'..." -Level 'INFO'
                    foreach ($objName in $dbGroup.Group) {
                        $parts = $objName.Split('.'); $consolidatedSchema += Get-ObjectSchema -ServerInstance $selectedServer -DatabaseName $parts[0] -SchemaName $parts[1] -ObjectName $parts[2]
                    }
                }
            } else { Write-Log -Message "No user database objects were found for $baseName. Halting." -Level 'WARN'; continue }

            if ([string]::IsNullOrWhiteSpace($consolidatedSchema)) { Write-Log -Message "Schema collection resulted empty. Halting." -Level 'WARN'; continue }
            $schemaPath = Join-Path -Path $script:AnalysisPath -ChildPath "_ConsolidatedSchema.txt"
            try { $consolidatedSchema | Set-Content -Path $schemaPath -Encoding UTF8; Write-Log -Message "Consolidated schema saved." -Level 'DEBUG' } catch { Write-Log -Message "Could not save consolidated schema file." -Level 'WARN' }

            $finalScript = Invoke-GeminiAnalysis -ModelName $script:ChosenModel -ApiKey $script:GeminiApiKey -PromptBody $selectedPromptBody -FullSqlText $sqlQueryText -ConsolidatedSchema $consolidatedSchema -MasterPlanXml $masterPlanXml -SqlServerVersion $sqlVersion
            
            if ($finalScript) {
                $finalScript = $finalScript.Trim()
                $analyzedScriptPath = Join-Path -Path $script:AnalysisPath -ChildPath "${baseName}_analyzed.sql"
                $finalScript | Out-File -FilePath $analyzedScriptPath -Encoding UTF8
                if ($OpenTunedFile.IsPresent) { try { Invoke-Item -Path $analyzedScriptPath } catch { Write-Log -Message "Failed to open the analyzed file automatically: $($_.Exception.Message)" -Level 'WARN' } }
                New-AnalysisSummary -TunedScript $finalScript -TotalStatementCount $statementNodes.Count -AnalysisPath $script:AnalysisPath
                Write-Log -Message "Tuning analysis complete for $baseName." -Level 'SUCCESS'
            } else {
                Write-Log -Message "Tuning analysis halted for $baseName due to an AI error." -Level 'ERROR'
            }
        } catch {
            Write-Log -Message "CRITICAL UNHANDLED ERROR during tuning of '$($analysisItem.BaseName)': $($_.Exception.Message). Moving to next item." -Level 'ERROR'
            Write-Log -Message "Stack Trace: $($_.ScriptStackTrace)" -Level 'DEBUG'
        }
    }
    
    Write-Log -Message "`n--- Batch Tuning Complete ---" -Level 'SUCCESS'
    Write-Log -Message "All analysis folders for this batch are located in:" -Level 'SUCCESS'
    Write-Log -Message "$batchFolderPath" -Level 'RESULT'
    if (-not $OpenTunedFile.IsPresent -or $DebugMode.IsPresent) { try { Invoke-Item -Path $batchFolderPath } catch { Write-log -Message "Could not automatically open the batch folder." -Level 'WARN' } }
}

function Start-CodeReviewAnalysis {
    Write-Log -Message "Starting new batch in Code Review mode." -Level 'INFO'

    # Select the AI prompt once for the entire batch.
    $selectedPromptBody = Select-AIPrompt -AnalysisType 'CodeReview'
    if (-not $selectedPromptBody) {
        Write-Log -Message "No valid prompt was selected. Halting batch." -Level 'ERROR'
        return
    }

    [array]$analysisInputs = Get-AnalysisInputs
    if ($null -eq $analysisInputs -or $analysisInputs.Count -eq 0) {
        Write-Log -Message "No valid inputs were found for analysis." -Level 'WARN'
        return
    }

    $sanitizedModelName = $script:ChosenModel -replace '[.-]', '_'
    $modelSpecificPath = Join-Path -Path $script:OptimusConfig.AnalysisBaseDir -ChildPath $sanitizedModelName
    if (-not (Test-Path -Path $modelSpecificPath)) { New-Item -Path $modelSpecificPath -ItemType Directory -Force | Out-Null }
    $batchTimestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $batchFolderPath = Join-Path -Path $modelSpecificPath -ChildPath $batchTimestamp
    New-Item -Path $batchFolderPath -ItemType Directory -Force | Out-Null
    Write-Log -Message "`nCreated batch analysis folder: $batchFolderPath" -Level 'SUCCESS'

    foreach ($analysisItem in $analysisInputs) {
        try {
            $baseName = $analysisItem.BaseName
            $sqlQueryText = $analysisItem.SqlText
            Write-Log -Message "`n--- Starting Code Review for: $baseName ---" -Level 'SUCCESS'
            $script:AnalysisPath = Join-Path -Path $batchFolderPath -ChildPath $baseName
            New-Item -Path $script:AnalysisPath -ItemType Directory -Force | Out-Null
            $script:LogFilePath = Join-Path -Path $script:AnalysisPath -ChildPath "ExecutionLog.txt"
            "# Optimus v$($script:CurrentVersion) Code Review Log | File: $baseName | Started: $(Get-Date)" | Out-File -FilePath $script:LogFilePath -Encoding utf8
            Write-Log -Message "Created analysis directory: '$($script:AnalysisPath)'" -Level 'INFO'

            # In Code Review mode, we don't need a DB connection. We go straight to the AI.
            $finalScript = Invoke-GeminiCodeReview -ModelName $script:ChosenModel -ApiKey $script:GeminiApiKey -PromptBody $selectedPromptBody -FullSqlText $sqlQueryText

            if ($finalScript) {
                $finalScript = $finalScript.Trim()
                $analyzedScriptPath = Join-Path -Path $script:AnalysisPath -ChildPath "${baseName}_analyzed.sql"
                $finalScript | Out-File -FilePath $analyzedScriptPath -Encoding UTF8
                if ($OpenTunedFile.IsPresent) { try { Invoke-Item -Path $analyzedScriptPath } catch { Write-Log -Message "Failed to open the analyzed file automatically: $($_.Exception.Message)" -Level 'WARN' } }
                # For code review, we can approximate statement count by line breaks or semicolons for the summary.
                $statementCount = ($sqlQueryText -split 'GO|\n|;').Count
                New-AnalysisSummary -TunedScript $finalScript -TotalStatementCount $statementCount -AnalysisPath $script:AnalysisPath
                Write-Log -Message "Code review complete for $baseName." -Level 'SUCCESS'
            } else {
                Write-Log -Message "Code review halted for $baseName due to an AI error." -Level 'ERROR'
            }
        }
        catch {
            Write-Log -Message "CRITICAL UNHANDLED ERROR during code review of '$($analysisItem.BaseName)': $($_.Exception.Message). Moving to next item." -Level 'ERROR'
            Write-Log -Message "Stack Trace: $($_.ScriptStackTrace)" -Level 'DEBUG'
        }
    }
    
    Write-Log -Message "`n--- Batch Code Review Complete ---" -Level 'SUCCESS'
    Write-Log -Message "All analysis folders for this batch are located in:" -Level 'SUCCESS'
    Write-Log -Message "$batchFolderPath" -Level 'RESULT'
    if (-not $OpenTunedFile.IsPresent -or $DebugMode.IsPresent) { try { Invoke-Item -Path $batchFolderPath } catch { Write-log -Message "Could not automatically open the batch folder." -Level 'WARN' } }
}

# --- Main Application Logic ---
function Start-Optimus {
    # Define the current version of the script in one place.
    $script:CurrentVersion = "3.4.0"

    if ($DebugMode) { Write-Log -Message "Starting Optimus v$($script:CurrentVersion) in Debug Mode." -Level 'DEBUG'}

    # On first run, or if config is reset, set up the .optimus directory and files.
    if ($ResetConfiguration) {
        if (-not (Reset-OptimusConfiguration)) {
            Write-Log -Message "Reset cancelled by user. Exiting script." -Level 'WARN'
            return
        }
    }
    if (-not (Initialize-Configuration)) { return }

    # Handle the prompt import utility mode
    if ($PSCmdlet.ParameterSetName -eq 'ImportPrompt') {
        Import-UserPrompt -FilePath $Path
        return # Exit after import
    }

    # Group prerequisite checks for analysis workflows
    $checksPassed = {
        if (-not (Test-WindowsEnvironment)) { return $false }
        if (-not (Test-PowerShellVersion)) { return $false }
        if (-not (Test-InternetConnection)) { Write-Log -Message "Exiting due to no internet connection or user choice." -Level 'ERROR'; return $false }
        Invoke-OptimusVersionCheck -CurrentVersion $script:CurrentVersion
        if (-not (Get-And-Set-ApiKey)) { Write-Log -Message "Exiting due to missing API key." -Level 'ERROR'; return $false }
        $script:ChosenModel = Get-And-Set-Model
        if (-not $script:ChosenModel) { Write-Log -Message "Exiting due to no model being selected." -Level 'ERROR'; return $false }

        # The SQLServer module is only required for Tuning mode.
        if (-not $CodeReview.IsPresent) {
             if (-not (Test-SqlServerModule)) { return $false }
        } else {
            Write-Log -Message "Code Review mode selected, skipping SQL Server module check." -Level 'DEBUG'
        }
        return $true
    }.Invoke()

    if (-not $checksPassed) { return }

    Write-Log -Message "`n--- Welcome to Optimus v$($script:CurrentVersion) ---" -Level 'SUCCESS'
    if (-not $DebugMode) { Write-Log -Message "All prerequisite checks passed." -Level 'SUCCESS' }
    
    do { # Outer loop to allow running multiple batches
        $script:AnalysisPath = $null 
        $script:LogFilePath = $null
        $analysisMode = 'Tuning' # Default to tuning

        # Determine Analysis Mode
        if ($CodeReview.IsPresent) {
            $analysisMode = 'CodeReview'
        } elseif ($PSCmdlet.ParameterSetName -eq 'Interactive') {
            Write-Log -Message "`nPlease select the type of analysis to perform:" -Level 'INFO'
            Write-Log -Message "   [1] T-SQL Performance Tuning (Default)" -Level 'PROMPT'
            Write-Log -Message "   [2] Standardized Code Review" -Level 'PROMPT'
            Write-Log -Message "   [Q] Quit" -Level 'PROMPT'
            
            while($true) {
                Write-Log -Message "   Enter your choice: " -Level 'PROMPT' -NoNewLine
                $choice = Read-Host
                Write-Host ""
                Write-Log -Message "User Input: $choice" -Level 'DEBUG'

                if ([string]::IsNullOrWhiteSpace($choice)) { $choice = '1' }

                $isValidChoice = $true # Assume the choice is valid initially
                switch($choice) {
                    '1' { $analysisMode = 'Tuning' }
                    '2' { $analysisMode = 'CodeReview' }
                    'Q' { Write-Log -Message "Exiting." -Level 'INFO'; return }
                    default {
                        Write-Log -Message "Invalid selection. Please enter 1, 2, or Q." -Level 'ERROR'
                        $isValidChoice = $false # The choice was not valid
                    }
                }
                
                if ($isValidChoice) {
                    break # This break now correctly exits the 'while' loop
                }
            }
        }

        # Call the appropriate handler based on the selected mode
        if ($analysisMode -eq 'Tuning') {
            Start-TuningAnalysis
        } else { # 'CodeReview'
            Start-CodeReviewAnalysis
        }
        
        Write-Log -Message "`nWould you like to analyze another batch of files? (Y/N): " -Level 'PROMPT' -NoNewLine
        $response = Read-Host
        Write-Host ""
        Write-Log -Message "User Input: $response" -Level 'DEBUG'

    } while ($response -match '^[Yy]$')
    Write-Log -Message "Thank you for using Optimus. Exiting." -Level 'SUCCESS'
}

# --- Script Entry Point ---
Start-Optimus