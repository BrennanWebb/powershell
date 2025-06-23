<#
.SYNOPSIS
    Optimus is a T-SQL tuning advisor that leverages the Gemini AI for performance recommendations.

.DESCRIPTION
    This script performs a holistic analysis of a T-SQL query. It can process a T-SQL script from a single file, a list of files, all .sql files
    in a specified folder, or a raw T-SQL string. It generates a master execution plan to validate syntax and identify database objects. It then 
    queries each required database to build a comprehensive schema document for all user-defined objects. All execution steps, 
    messages, and recommendations are recorded in a detailed log file for each analysis.

.PARAMETER SQLFile
    The path to one or more .sql files to be analyzed. For multiple files, provide a comma-separated list. This parameter
    cannot be used with -FolderPath or -AdhocSQL.

.PARAMETER FolderPath
    The path to a single folder. The script will analyze all .sql files found in this folder (non-recursively). 
    This parameter cannot be used with -SQLFile or -AdhocSQL.

.PARAMETER AdhocSQL
    A string containing the T-SQL query to be analyzed. This is useful for passing queries directly from other applications.
    This parameter cannot be used with -SQLFile or -FolderPath.

.PARAMETER ServerName
    An optional parameter to specify the SQL Server instance for the analysis, bypassing the interactive menu.

.PARAMETER UseActualPlan
    An optional switch to generate the 'Actual' execution plan. This WILL execute the query.
    If not present, the script defaults to 'Estimated' or will prompt in interactive mode.

.PARAMETER ResetConfiguration
    An optional switch to trigger an interactive menu that allows for resetting user configurations.
    This can be used to clear the saved API key, server list, and optionally, all past analysis reports.

.PARAMETER DebugMode
    Enables detailed diagnostic output to the console. All messages are always written to the execution log file regardless of this setting.

.EXAMPLE
    .\Optimus.ps1
    Runs the script in interactive mode, using a graphical file picker and server selection menu.

.EXAMPLE
    .\Optimus.ps1 -FolderPath "C:\My TSQL Projects\Batch1" -ServerName "PROD-DB01\SQL2022"
    Runs a fully automated analysis on all .sql files in the folder against the specified server.

.EXAMPLE
    .\Optimus.ps1 -AdhocSQL "SELECT * FROM Sales.SalesOrderHeader WHERE OrderDate > '2011-01-01'" -ServerName "PROD-DB01\SQL2022"
    Runs a fully automated analysis on the provided T-SQL string.

.NOTES
    Designer: Brennan Webb
    Script Engine: Gemini
    Version: 2.4
    Created: 2025-06-21
    Modified: 2025-06-23
    Change Log:
    - v2.4: Minor wording change for AI analysis message.
    - v2.3: Added -AdhocSQL parameter and 'Adhoc' parameter set to allow for direct T-SQL string analysis. Refactored input handling.
    - v2.2: Added -ServerName parameter to allow for non-interactive server selection, enabling full automation.
    - v2.1: Modified plan selection to be non-interactive for automated runs. It now defaults to 'Estimated' plan unless -UseActualPlan is specified.
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
[CmdletBinding(DefaultParameterSetName = 'Interactive')]
param (
    [Parameter(Mandatory=$true, ParameterSetName='Files')]
    [string[]]$SQLFile,

    [Parameter(Mandatory=$true, ParameterSetName='Folder')]
    [string]$FolderPath,

    [Parameter(Mandatory=$true, ParameterSetName='Adhoc')]
    [string]$AdhocSQL,

    [Parameter(Mandatory=$false)]
    [string]$ServerName,

    [Parameter(Mandatory=$false)]
    [switch]$UseActualPlan,
    
    [Parameter(Mandatory=$false)]
    [switch]$ResetConfiguration,

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
    
    $inputObjects = [System.Collections.Generic.List[object]]::new()

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

        # Combine all found plan fragments into a single master XML document
        $masterPlanXml = "<MasterShowPlan>" + ($planFragments -join '') + "</MasterShowPlan>"
        
        try {
            [xml]$masterPlanXml | Out-Null
            Write-Log -Message "Successfully generated and validated master execution plan for all statements." -Level 'SUCCESS'
            return $masterPlanXml
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
        [string]$ModelName,
        [securestring]$ApiKey, 
        [string]$FullSqlText, 
        [string]$ConsolidatedSchema, 
        [string]$MasterPlanXml,
        [string]$SqlServerVersion
    )
    Write-Log -Message "Entering Function: Invoke-GeminiAnalysis" -Level 'DEBUG'
    Write-Log -Message "Sending full script to Gemini for analysis..." -Level 'INFO'

    $plainTextApiKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ApiKey))
    $uri = "https://generativelanguage.googleapis.com/v1beta/models/$($ModelName):generateContent?key=$plainTextApiKey"

    # --- Start of Modified Prompt ---
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


(Add more numbered recommendations as needed for the same statement)
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
    # --- End of Modified Prompt ---

    $promptPath = Join-Path -Path $script:AnalysisPath -ChildPath "_FinalAIPrompt.txt"
    try { $prompt | Set-Content -Path $promptPath -Encoding UTF8; Write-Log -Message "Final AI prompt saved for review at: $promptPath" -Level 'DEBUG' } catch { Write-Log -Message "Could not save final AI prompt file." -Level 'WARN' }

    $bodyObject = @{ contents = @( @{ parts = @( @{ text = $prompt } ) } ) }
    $bodyJson = $bodyObject | ConvertTo-Json -Depth 10

    try {
        $response = Invoke-RestMethod -Uri $uri -Method Post -Body $bodyJson -ContentType 'application/json' -ErrorAction Stop
        $rawAiResponse = $response.candidates[0].content.parts[0].text
        Write-Log -Message "Successfully received raw response from Gemini API." -Level 'DEBUG'
        
        $cleanedScript = $rawAiResponse -replace '(?i)^```sql\s*','' -replace '```\s*$',''
        Write-Log -Message "Cleaned response received from AI." -Level 'DEBUG'
        
        Write-Log -Message "AI analysis complete." -Level 'SUCCESS'
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
        # Updated Regex to find the start of our new comment block format
        $recommendationBlockRegex = '(?s)\/\*\s*--- Optimus Analysis ---(.*?)\*\/';
        $recommendationBlocks = [regex]::Matches($TunedScript, $recommendationBlockRegex)
        
        # Regex to count individual recommendations within a block
        $individualRecommendationRegex = '\[\d+\]\s*Recommendation'
        $totalRecommendations = 0
        foreach ($block in $recommendationBlocks) {
            $totalRecommendations += ([regex]::Matches($block.Value, $individualRecommendationRegex)).Count
        }

        $summaryContent += "Total Statements Analyzed: $TotalStatementCount`n"
        $summaryContent += "Statements with Recommendations: $($recommendationBlocks.Count)`n"
        $summaryContent += "Total Individual Recommendations: $totalRecommendations`n`n"

        if ($recommendationBlocks.Count -gt 0) {
            $summaryContent += "--- Summary of Findings ---`n"
            $problemRegex = 'Problem:(.*?)(?=\s*-\s*Recommended Code:|\s*$)';
            
            $findingIndex = 1
            foreach ($block in $recommendationBlocks) {
                $problemMatches = [regex]::Matches($block.Value, $problemRegex)
                foreach ($problem in $problemMatches) {
                    $problemText = $problem.Groups[1].Value.Trim()
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

# --- Main Application Logic ---
function Start-Optimus {
    # Define the current version of the script in one place.
    $script:CurrentVersion = "2.4"

    if ($DebugMode) { Write-Log -Message "Starting Optimus v$($script:CurrentVersion) in Debug Mode." -Level 'DEBUG'}

    # Group prerequisite checks
    $checksPassed = {
        if (-not (Test-WindowsEnvironment)) { return $false }
        if (-not (Test-PowerShellVersion)) { return $false }
        if ($ResetConfiguration) {
            if (-not (Reset-OptimusConfiguration)) {
                Write-Log -Message "Reset cancelled by user. Exiting script." -Level 'WARN'
                return $false
            }
        }
        if (-not (Initialize-Configuration)) { return $false }
        if (-not (Test-InternetConnection)) {
            Write-Log -Message "Exiting due to no internet connection or user choice." -Level 'ERROR'
            return $false
        }
        
        Invoke-OptimusVersionCheck -CurrentVersion $script:CurrentVersion

        if (-not (Test-SqlServerModule)) { return $false }
        if (-not (Get-And-Set-ApiKey)) {
            Write-Log -Message "Exiting due to missing API key." -Level 'ERROR'
            return $false
        }
        
        $script:ChosenModel = Get-And-Set-Model
        if (-not $script:ChosenModel) {
            Write-Log -Message "Exiting due to no model being selected." -Level 'ERROR'
            return $false
        }
        return $true
    }.Invoke()

    if (-not $checksPassed) { return }

    Write-Log -Message "`n--- Welcome to Optimus v$($script:CurrentVersion) ---" -Level 'SUCCESS'
    if (-not $DebugMode) { Write-Log -Message "All prerequisite checks passed." -Level 'SUCCESS' }
    
    do { # Outer loop to allow running multiple batches
        $script:AnalysisPath = $null 
        $script:LogFilePath = $null

        # Server Selection Logic
        $selectedServer = $null
        # If the ServerName parameter is provided, use it.
        if (-not [string]::IsNullOrWhiteSpace($ServerName)) {
            Write-Log -Message "ServerName parameter provided, attempting to connect to '$ServerName'." -Level 'INFO'
            if (Test-SqlServerConnection -ServerInstance $ServerName) {
                $selectedServer = $ServerName
                # Logic to save the validated server to the config file for future use
                try {
                    [array]$servers = Get-Content -Path $script:OptimusConfig.ServerFile | ConvertFrom-Json
                    if ($selectedServer -notin $servers) {
                        $servers += $selectedServer
                        ($servers | Sort-Object -Unique) | ConvertTo-Json -Depth 5 | Set-Content -Path $script:OptimusConfig.ServerFile
                        Write-Log -Message "'$selectedServer' has been validated and saved to the configuration." -Level 'DEBUG'
                    }
                } catch {
                    Write-Log -Message "Could not save the provided server name to the configuration file." -Level 'WARN'
                }
            } else {
                # If the connection fails, stop the script.
                Write-Log -Message "Connection test to the server '$ServerName' failed. Please check the server name and permissions." -Level 'ERROR'
                return # Exit the function/script
            }
        } else {
            # Otherwise, fall back to the interactive menu.
            $selectedServer = Select-SqlServer
        }

        # Final check to ensure a server was successfully selected.
        if (-not $selectedServer) {
            Write-Log -Message "No valid SQL Server was selected or provided. Halting analysis for this batch." -Level 'WARN'
            break 
        }
        
        # Get the inputs for analysis (from files, folder, or ad-hoc string)
        [array]$analysisInputs = Get-AnalysisInputs
        if ($null -eq $analysisInputs -or $analysisInputs.Count -eq 0) {
            Write-Log -Message "No valid inputs were found for analysis." -Level 'WARN'
            continue
        }

        # Select Plan Type once for the entire batch.
        $useActualPlanSwitch = $UseActualPlan.IsPresent

        # In interactive mode, if the plan type isn't specified, we must ask the user.
        # In non-interactive modes ('Files', 'Folder', 'Adhoc'), it defaults to Estimated unless -UseActualPlan is specified.
        if ($PSCmdlet.ParameterSetName -eq 'Interactive' -and -not $UseActualPlan.IsPresent) {
            Write-Log -Message "`nWhich execution plan would you like to generate for this batch?" -Level 'INFO'
            Write-Log -Message "   [1] Estimated (Default - Recommended, does not run the query)" -Level 'PROMPT'
            Write-Log -Message "   [2] Actual (Executes the query, use with caution on all files)" -Level 'PROMPT'
            Write-Log -Message "   Enter your choice: " -Level 'PROMPT' -NoNewLine
            $choice = Read-Host
            Write-Host ""
            Write-Log -Message "User Input: $choice" -Level 'DEBUG'
            if ($choice -eq '2') {
                Write-Log -Message "Proceeding with 'Actual Execution Plan'. This will EXECUTE every SQL script in the batch." -Level 'WARN'
                $useActualPlanSwitch = $true
            } else {
                Write-Log -Message "Defaulting to 'Estimated Execution Plan' for this batch." -Level 'INFO'
            }
        }

        # Create the model-specific and batch parent folders
        $sanitizedModelName = $script:ChosenModel -replace '[.-]', '_'
        $modelSpecificPath = Join-Path -Path $script:OptimusConfig.AnalysisBaseDir -ChildPath $sanitizedModelName

        # Ensure the model-specific parent directory exists
        if (-not (Test-Path -Path $modelSpecificPath)) {
            New-Item -Path $modelSpecificPath -ItemType Directory -Force | Out-Null
        }

        $batchTimestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $batchFolderPath = Join-Path -Path $modelSpecificPath -ChildPath $batchTimestamp
        New-Item -Path $batchFolderPath -ItemType Directory -Force | Out-Null
        Write-Log -Message "`nCreated batch analysis folder: $batchFolderPath" -Level 'SUCCESS'

        # Loop through each input object
        foreach ($input in $analysisInputs) {
            try {
                $baseName = $input.BaseName
                $sqlQueryText = $input.SqlText

                Write-Log -Message "`n--- Starting Analysis for: $baseName ---" -Level 'SUCCESS'

                # Create a sub-folder for this specific file's analysis
                $script:AnalysisPath = Join-Path -Path $batchFolderPath -ChildPath $baseName
                New-Item -Path $script:AnalysisPath -ItemType Directory -Force | Out-Null

                # Set up the log file path for this specific analysis
                $script:LogFilePath = Join-Path -Path $script:AnalysisPath -ChildPath "ExecutionLog.txt"
                "# Optimus v$($script:CurrentVersion) Execution Log | File: $baseName | Started: $(Get-Date)" | Out-File -FilePath $script:LogFilePath -Encoding utf8
                
                Write-Log -Message "Created analysis directory: '$($script:AnalysisPath)'" -Level 'INFO'
                $sqlVersion = Get-SqlServerVersion -ServerInstance $selectedServer
                Write-Log -Message "Detected SQL Server Version: $sqlVersion" -Level 'DEBUG'
                
                # 1. Get Master Plan. Use a default context but the plan will have fully-qualified names.
                $initialDbContext = ([regex]::Match($sqlQueryText, '(?im)^\s*USE\s+\[?([\w\d_]+)\]?')).Groups[1].Value
                if ([string]::IsNullOrWhiteSpace($initialDbContext)) { $initialDbContext = 'master' }
                $masterPlanXml = Get-MasterExecutionPlan -ServerInstance $selectedServer -DatabaseContext $initialDbContext -FullQueryText $sqlQueryText -IsActualPlan:$useActualPlanSwitch
                if (-not $masterPlanXml) { 
                    Write-Log -Message "Could not generate a master plan for $baseName. Skipping to next item." -Level 'ERROR'
                    continue 
                }
                
                $planPath = Join-Path -Path $script:AnalysisPath -ChildPath "_MasterPlan.xml"
                try { $masterPlanXml | Set-Content -Path $planPath -Encoding UTF8; Write-Log -Message "Master execution plan saved." -Level 'DEBUG' } catch { Write-Log -Message "Could not save master plan file." -Level 'WARN' }

                # 2. Parse unique object names from the plan
                [xml]$masterPlan = $masterPlanXml
                $ns = New-Object System.Xml.XmlNamespaceManager($masterPlan.NameTable)
                $ns.AddNamespace("sql", "http://schemas.microsoft.com/sqlserver/2004/07/showplan")
                [string[]]$uniqueObjectNames = @(Get-ObjectsFromPlan -MasterPlan $masterPlan -NamespaceManager $ns)
                $statementNodes = $masterPlan.SelectNodes("//sql:StmtSimple", $ns)
                
                # 3. Build the consolidated schema document using the robust iterative method
                $consolidatedSchema = ""
                if ($null -ne $uniqueObjectNames -and $uniqueObjectNames.Count -gt 0) {
                    Write-Log -Message "Starting schema collection for all objects..." -Level 'INFO'
                    $objectsByDb = $uniqueObjectNames | Group-Object { ($_ -split '\.')[0] }

                    foreach ($dbGroup in $objectsByDb) {
                        $dbName = $dbGroup.Name
                        # Skip the hidden mssqlsystemresource database as it cannot be queried directly
                        if ($dbName -eq 'mssqlsystemresource') {
                            Write-Log -Message "Skipping schema collection for internal database: 'mssqlsystemresource'." -Level 'DEBUG'
                            continue
                        }
                        
                        Write-Log -Message "Querying database '$dbName'..." -Level 'INFO'
                        foreach ($objName in $dbGroup.Group) {
                            $parts = $objName.Split('.')
                            $consolidatedSchema += Get-ObjectSchema -ServerInstance $selectedServer -DatabaseName $parts[0] -SchemaName $parts[1] -ObjectName $parts[2]
                        }
                    }
                } else {
                    Write-Log -Message "No user database objects were found in the execution plan for $baseName. Halting analysis for this item." -Level 'WARN'
                    continue
                }

                if ([string]::IsNullOrWhiteSpace($consolidatedSchema)) {
                     Write-Log -Message "Schema collection resulted in an empty document. This could be due to permissions or missing objects. Halting analysis for this item." -Level 'WARN'
                     continue
                }

                $schemaPath = Join-Path -Path $script:AnalysisPath -ChildPath "_ConsolidatedSchema.txt"
                try { $consolidatedSchema | Set-Content -Path $schemaPath -Encoding UTF8; Write-Log -Message "Consolidated schema saved." -Level 'DEBUG' } catch { Write-Log -Message "Could not save consolidated schema file." -Level 'WARN' }

                # 4. Make single "Omnibus" call to AI
                $finalScript = Invoke-GeminiAnalysis -ModelName $script:ChosenModel -ApiKey $script:GeminiApiKey -FullSqlText $sqlQueryText -ConsolidatedSchema $consolidatedSchema -MasterPlanXml $masterPlanXml -SqlServerVersion $sqlVersion
                
                # 5. Process and save the final result
                if ($finalScript) {
                    $finalScript = $finalScript.Trim()
                    $tunedScriptPath = Join-Path -Path $script:AnalysisPath -ChildPath "${baseName}_tuned.sql"
                    $finalScript | Out-File -FilePath $tunedScriptPath -Encoding UTF8
                    New-AnalysisSummary -TunedScript $finalScript -TotalStatementCount $statementNodes.Count -AnalysisPath $script:AnalysisPath
                    Write-Log -Message "Analysis complete for $baseName." -Level 'SUCCESS'
                } else {
                    Write-Log -Message "Analysis halted for $baseName due to an error or empty response from the AI." -Level 'ERROR'
                }
            }
            catch {
                Write-Log -Message "CRITICAL UNHANDLED ERROR during analysis of '$($input.BaseName)': $($_.Exception.Message). Moving to next item." -Level 'ERROR'
                Write-Log -Message "Stack Trace: $($_.ScriptStackTrace)" -Level 'DEBUG'
            }
        } # End foreach loop

        Write-Log -Message "`n--- Batch Analysis Complete ---" -Level 'SUCCESS'
        Write-Log -Message "All analysis folders for this batch are located in:" -Level 'SUCCESS'
        Write-Log -Message "$batchFolderPath" -Level 'RESULT'

        try {
            Invoke-Item -Path $batchFolderPath
            Write-Log -Message "Opening batch folder in File Explorer." -Level 'DEBUG'
        } catch {
            Write-Log -Message "Could not automatically open the batch folder. Please navigate to the path above manually." -Level 'WARN'
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