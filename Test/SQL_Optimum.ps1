<#
.SYNOPSIS
    SQL_Optimum is a tuning advisor that leverages the Gemini API to analyze T-SQL queries and provide performance recommendations.

.DESCRIPTION
    This script is self-contained and automatically manages its own dependencies by downloading the required Microsoft SQL Server libraries. 
    It facilitates SQL performance tuning by collecting a T-SQL query, its estimated execution plan, and relevant object schemas. 
    It sends this information to the Google Gemini API for analysis.

.PARAMETER DebugMode
    A switch that, when present, enables verbose diagnostic logging to the console.

.NOTES
    Author:     Powershell Developer (Gemini) & brennan.webb
    Version:    12.0
    Created:    15-June-2025
    Modified:   16-June-2025
    Requires:   PowerShell 5.1 or higher. An active internet connection on first run to download dependencies.
#>
[CmdletBinding()]
param (
    [string]$ApiKey,
    [string]$ServerName,
    [string]$SqlFilePath,
    [switch]$DebugMode
)

Set-StrictMode -Version Latest

$script:DebugModeActive = $PSBoundParameters.ContainsKey('DebugMode')
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

#region Helper Functions
function Write-Host-Cyan { param([string]$Message) Write-Host -Object $Message -ForegroundColor Cyan }
function Write-Host-Green { param([string]$Message) Write-Host -Object $Message -ForegroundColor Green }
function Write-Host-Yellow { param([string]$Message) Write-Host -Object $Message -ForegroundColor Yellow }
function Write-Host-Red { param([string]$Message) Write-Error -Message $Message }
function Write-Host-White { param([string]$Message) Write-Host -Object $Message -ForegroundColor White }
function Write-DebugLog {
    param([string]$Message)
    if ($script:DebugModeActive) { Write-Host "[DEBUG] $Message" -ForegroundColor Gray }
}
#endregion

#region Prerequisite and Configuration Setup
function Ensure-Module {
    param([string]$ModuleName)
    Write-Host-Cyan "Checking for required '$ModuleName' module..."
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host-Yellow "'$ModuleName' module not found. Attempting to install..."
        try {
            Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
            Write-Host-Green "'$ModuleName' module installed successfully."
        } catch {
            Write-Host-Red "FATAL: Failed to install '$ModuleName' module. Error: $($_.Exception.Message)"
            throw "Halting script execution due to missing dependency: $ModuleName."
        }
    } else {
        Write-Host-Green "'$ModuleName' module is already installed."
    }
    Import-Module -Name $ModuleName -DisableNameChecking -Force > $null
}

function Initialize-ConfigurationDirectory {
    Write-Host-Cyan "Initializing configuration directory..."
    $configPath = Join-Path -Path $env:USERPROFILE -ChildPath ".sql_optimum"
    if (-not (Test-Path -Path $configPath)) {
        try {
            New-Item -Path $configPath -ItemType Directory -ErrorAction Stop | Out-Null
            Write-Host-Green "Configuration directory created at '$configPath'."
        } catch {
            Write-Host-Red "FATAL: Could not create configuration directory at '$configPath'."
            throw "Halting script execution due to file system error."
        }
    } else {
        Write-Host-Green "Configuration directory already exists."
    }
    return $configPath
}

function Ensure-ParserLibraries {
    param(
        [string]$ConfigDir
    )
    Write-Host-Cyan "Checking for required SQL Parser libraries..."
    $libDir = Join-Path -Path $ConfigDir -ChildPath 'lib'
    
    $requiredDlls = @{
        'Microsoft.SqlServer.Management.SqlParser.dll' = $false;
        'Microsoft.SqlServer.Management.Sdk.Sfc.dll'   = $false;
    }

    if (-not (Test-Path $libDir)) {
        New-Item -Path $libDir -ItemType Directory | Out-Null
    }

    $requiredDlls.Keys | ForEach-Object {
        if (Test-Path (Join-Path -Path $libDir -ChildPath $_)) {
            $requiredDlls[$_] = $true
        }
    }

    if ($requiredDlls.ContainsValue($false)) {
        Write-Host-Yellow "One or more required SQL Parser libraries not found. Attempting to download..."
        try {
            $packagesToDownload = @{
                'Microsoft.SqlServer.Management.SqlParser' = '172.0.0';
                'Microsoft.SqlServer.Management.Sdk.Sfc'   = '161.47021.0'
            }
            
            foreach ($pkgName in $packagesToDownload.Keys) {
                $dllName = "$($pkgName).dll"
                if (-not $requiredDlls[$dllName]) {
                    $pkgVersion = $packagesToDownload[$pkgName]
                    $downloadUrl = "https://www.nuget.org/api/v2/package/$pkgName/$pkgVersion"
                    $outFile = Join-Path $libDir "$($pkgName).$($pkgVersion).nupkg"
                    Write-DebugLog "Downloading package '$pkgName' with Invoke-WebRequest from $downloadUrl"
                    Invoke-WebRequest -Uri $downloadUrl -OutFile $outFile -UseBasicParsing
                }
            }

            $nupkgFiles = Get-ChildItem -Path $libDir -Filter "*.nupkg"
            foreach ($nupkgFile in $nupkgFiles) {
                $zipPath = $nupkgFile.FullName -replace '\.nupkg$', '.zip'
                Rename-Item -Path $nupkgFile.FullName -NewName $zipPath -Force
                $tempExtractPath = Join-Path $libDir "temp_extract_$($nupkgFile.BaseName)"
                Expand-Archive -Path $zipPath -DestinationPath $tempExtractPath -Force
                $sourceLibDir = (Get-ChildItem -Path $tempExtractPath -Recurse -Filter '*.dll' | Select-Object -First 1).Directory
                if ($sourceLibDir) {
                    Copy-Item -Path "$($sourceLibDir.FullName)\*.dll" -Destination $libDir -Force
                }
                Remove-Item -Path $zipPath -Force
                Remove-Item -Path $tempExtractPath -Recurse -Force
            }
            Write-Host-Green "Libraries installed successfully."
        } catch {
            Write-Host-Red "FATAL: Failed to download and configure libraries. Error: $($_.Exception.Message)"
            throw "Halting execution due to dependency failure."
        }
    } else {
        Write-Host-Green "SQL Parser libraries are already present."
    }
    return $libDir
}
#endregion

#region User Input Functions
function Get-GeminiApiKey {
    param([string]$ConfigPath)
    $apiKeyPath = Join-Path -Path $ConfigPath -ChildPath "api.config"
    if (Test-Path -Path $apiKeyPath) {
        Write-Host-Cyan "Loading saved API key..."
        try { return (Get-Content -Path $apiKeyPath | ConvertTo-SecureString) }
        catch { Write-Host-Yellow "Could not read saved API key. It may be corrupted." }
    }
    Write-Host-White "Please enter your Google Gemini API Key. It will be stored securely for future use."
    $secureString = Read-Host -Prompt "API Key" -AsSecureString
    try {
        $secureString | ConvertFrom-SecureString | Set-Content -Path $apiKeyPath
        Write-Host-Green "API Key saved successfully."
    } catch {
        Write-Host-Yellow "Warning: Could not save API key to disk."
    }
    return $secureString
}

function Select-SqlServerInstance {
    param([string]$ConfigPath)
    $serverListPath = Join-Path -Path $ConfigPath -ChildPath "servers.json"
    $serverList = @()
    if (Test-Path -Path $serverListPath) {
        try { $serverList = @(Get-Content -Path $serverListPath -Raw | ConvertFrom-Json) }
        catch { Write-Host-Yellow "Warning: Could not parse server list. Starting empty." }
    }
    while ($true) {
        Write-Host-Cyan "`nPlease select a SQL Server instance:"
        for ($i = 0; $i -lt $serverList.Count; $i++) { Write-Host-White "  [$($i+1)] $($serverList[$i])" }
        Write-Host-White "  [A] Add a new server"
        Write-Host-White "  [Q] Quit"
        $choice = Read-Host -Prompt "Your choice"
        if ($choice -ieq 'q') { return $null }
        if ($choice -ieq 'a') {
            $newServer = Read-Host -Prompt "Enter new server name"
            Write-Host-Cyan "Testing connection to '$newServer'..."
            try {
                Invoke-Sqlcmd -ServerInstance $newServer -Query "SELECT @@SERVERNAME" -TrustServerCertificate -ConnectionTimeout 5 -ErrorAction Stop | Out-Null
                Write-Host-Green "Connection successful."
                if ($serverList -notcontains $newServer) {
                    $serverList += $newServer
                    $serverList | ConvertTo-Json | Set-Content -Path $serverListPath
                    Write-Host-Green "'$newServer' added to saved list."
                }
                return $newServer
            } catch { Write-Host-Red "Error: Could not connect to '$newServer'." }
        } elseif ($choice -match "^\d+$" -and [int]$choice -ge 1 -and [int]$choice -le $serverList.Count) {
            $selectedServer = $serverList[[int]$choice - 1]
            Write-Host-Green "Selected server: '$selectedServer'"
            return $selectedServer
        } else { Write-Host-Yellow "Invalid choice. Please try again." }
    }
}

function Get-SqlFile {
    try { Add-Type -AssemblyName System.Windows.Forms }
    catch {
        Write-Host-Red "Could not load .NET 'System.Windows.Forms' for file picker."
        return $null
    }
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Title = "Select the SQL file to analyze"
    $fileDialog.Filter = "SQL Files (*.sql)|*.sql"
    $fileDialog.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $filePath = $fileDialog.FileName
        try {
            $fileContent = Get-Content -Path $filePath -Raw -ErrorAction Stop
            Write-Host-Green "Successfully read file: $filePath"
            return [PSCustomObject]@{ FilePath = $filePath; Content = $fileContent }
        } catch {
            Write-Host-Red "Error reading file '$filePath'."
            return $null
        }
    } else {
        Write-Host-Yellow "File selection was cancelled."
        return $null
    }
}
#endregion

#region SQL Analysis Functions
function Parse-SqlScript {
    param(
        [string]$SqlContent,
        [string]$LibPath # Path to the directory containing our DLLs
    )
    Write-Host-Cyan "Parsing SQL script..."
    try {
        # Load the required dependency DLL first.
        $sfcDllPath = Join-Path -Path $LibPath -ChildPath 'Microsoft.SqlServer.Management.Sdk.Sfc.dll'
        if (Test-Path $sfcDllPath) {
             Write-DebugLog "Loading dependency DLL: $sfcDllPath"
             Add-Type -Path $sfcDllPath -ErrorAction Stop
        } else {
            throw "Sfc.dll dependency not found at $sfcDllPath."
        }
        
        # Now load the main parser DLL.
        $parserDllPath = Join-Path -Path $LibPath -ChildPath 'Microsoft.SqlServer.Management.SqlParser.dll'
        Write-DebugLog "Loading main parser DLL: $parserDllPath"
        Add-Type -Path $parserDllPath -ErrorAction Stop
    } catch {
        Write-Host-Red "FATAL: Could not load the locally stored parser libraries from $LibPath. Error: $($_.Exception.Message)"
        throw "Halting execution due to critical library failure."
    }

    $sqlVersion = [Microsoft.SqlServer.Management.SqlParser.Parser.SqlVersion]::Sql150
    $parser = [Microsoft.SqlServer.Management.SqlParser.Parser.Parser]::new($sqlVersion)
    
    $parseResult = $parser.Parse($SqlContent)
    if ($parseResult.Errors.Count -gt 0) {
        Write-Host-Yellow "Warning: The script contains syntax errors, results may be incomplete."
        foreach ($err in $parseResult.Errors) {
            Write-Host-Yellow "  - $($err.Message) (Line: $($err.Line), Column: $($err.Column))"
        }
    }
    $databaseName = $null
    $identifiedObjects = [System.Collections.Generic.List[string]]::new()
    $objectKeywords = @('FROM', 'JOIN', 'UPDATE', 'INTO', 'MERGE')
    $useStatements = $parseResult.Script.Batches.Statements | Where-Object { $_ -is [Microsoft.SqlServer.Management.SqlParser.Model.UseStatement] }
    if ($useStatements) { $databaseName = $useStatements[-1].DatabaseName.Value }
    $tokens = $parseResult.Script.Tokens
    for ($i = 0; $i -lt $tokens.Count; $i++) {
        $currentToken = $tokens[$i]
        if ($objectKeywords -contains $currentToken.Text.ToUpperInvariant()) {
            for ($j = $i + 1; $j -lt $tokens.Count; $j++) {
                $nextToken = $tokens[$j]
                if ($nextToken.Type -ne 'Whitespace') {
                    if ($nextToken.Type -eq 'Identifier') { $identifiedObjects.Add($nextToken.Text) }
                    $i = $j; break 
                }
            }
        }
    }
    $uniqueObjects = $identifiedObjects | Select-Object -Unique
    Write-Host-Green "SQL parsing complete."
    return [PSCustomObject]@{
        DatabaseName      = $databaseName
        IdentifiedObjects = $uniqueObjects
        Errors            = $parseResult.Errors
    }
}
#endregion

# --- Main Script Body ---
# Clear-Host # Commeneted out for testing
Write-Host-Cyan "--- Welcome to SQL_Optimum v12.0 ---"
Write-DebugLog "Debug mode is ON."

# Step 1: Initialize environment and dependencies
Ensure-Module -ModuleName 'SqlServer'
$configDir = Initialize-ConfigurationDirectory
$libPath = Ensure-ParserLibraries -ConfigDir $configDir
$secureApiKey = Get-GeminiApiKey -ConfigPath $configDir
if (-not $secureApiKey) {
    Write-Host-Red "No API Key provided. Exiting."
    exit 1
}

# Step 2: Main application loop
do {
    # Clear-Host # Commeneted out for testing
    Write-Host-Cyan "`n--- Starting New Analysis ---"
    $selectedServer = Select-SqlServerInstance -ConfigPath $configDir
    if (-not $selectedServer) { $continueAnalysis = 'n'; continue }
    $sqlFile = Get-SqlFile
    if (-not $sqlFile) {
        Write-Host-Yellow "Analysis cancelled due to no file selection."
    }
    else {
        $parsedSql = Parse-SqlScript -SqlContent $sqlFile.Content -LibPath $libPath
        if (-not $parsedSql.DatabaseName) {
            Write-Host-Yellow "No 'USE [DatabaseName]' statement was found in the script."
            $parsedSql.DatabaseName = Read-Host -Prompt "Please specify the database context for this query"
        }
        Write-Host-Green "Analysis proceeding with database '$($parsedSql.DatabaseName)'."
        if ($parsedSql.IdentifiedObjects) {
             Write-Host-Cyan "Found $($parsedSql.IdentifiedObjects.Count) potential objects to analyze: $($parsedSql.IdentifiedObjects -join ', ')"
        }
        Write-Host-Cyan "`nNext step will be to collect schema information and execution plan from '$selectedServer'."
    }
    Write-Host-Cyan "`n----------------------------------------------------"
    $response = Read-Host -Prompt "Would you like to analyze another SQL file? (Y/N)"
    $continueAnalysis = if ($response -and $response.Trim().ToLower() -eq 'y') { 'y' } else { 'n' }
} while ($continueAnalysis -eq 'y')

Write-Host-Cyan "`nExiting SQL_Optimum. Goodbye!"