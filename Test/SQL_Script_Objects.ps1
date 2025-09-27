<#
.SYNOPSIS
    Scripts SQL Server database objects to .sql files, organized into folders by type.

.DESCRIPTION
    This script connects to a specified SQL Server instance and database to generate schema-only scripts for various database objects.
    It mimics the functionality of the SSMS "Generate Scripts" wizard.
    
    The script first determines which object types exist in the database, then creates subdirectories only for those types within a main timestamped folder in the user's temp directory. This prevents the creation of empty folders.
    
    For Tables, the script is configured to include Indexes and Triggers within the table's script file.
    
    The script requires the 'SqlServer' PowerShell module. It will check if the module is installed and offer to install it if it is not found.

.PARAMETER SqlServerInstance
    The name of the SQL Server instance to connect to (e.g., 'localhost\SQLEXPRESS' or 'MyProdServer').

.PARAMETER DatabaseName
    The name of the database whose objects you want to script.

.PARAMETER DebugMode
    A switch parameter that, if present, enables verbose debugging messages throughout the script's execution.

.EXAMPLE
    PS C:\> .\Generate-SqlScripts.ps1 -SqlServerInstance "localhost\SQLEXPRESS" -DatabaseName "AdventureWorks2019"
    This command will script all supported objects from the 'AdventureWorks2019' database on the 'localhost\SQLEXPRESS' instance.

.EXAMPLE
    PS C:\> .\Generate-SqlScripts.ps1
    When run without parameters, the script will interactively prompt the user for the SQL Server Instance and Database Name.

.NOTES
    Designer: Brennan Webb & Gemini
    Script Engine: Gemini
    Version: 1.2.1
    Created: 2025-07-18
    Modified: 2025-08-11
    Change Log:
        1.2.0 (2025-08-11): Reverted from experimental versions. Base includes simple assembly loading and orphaned user check.
        1.2.1 (2025-08-11): Corrected type checking from 'Role' to 'DatabaseRole' for all database role objects.
---
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, HelpMessage="Enter the SQL Server instance name.")]
    [string]$SqlServerInstance,

    [Parameter(Mandatory=$false, HelpMessage="Enter the database name.")]
    [string]$DatabaseName,

    [Parameter(Mandatory=$false, HelpMessage="Enable verbose debug logging.")]
    [switch]$DebugMode
)

#region Helper Functions

# Helper function for colored console output.
function Write-HostInColor {
    param(
        [string]$Message,
        [string]$ForegroundColor
    )
    Write-Host -Object $Message -ForegroundColor $ForegroundColor
}

# Helper function to check for and prompt installation of the SqlServer module.
function Test-SqlServerModule {
    if (-not (Get-Module -ListAvailable -Name SqlServer)) {
        Write-HostInColor -Message "The required 'SqlServer' module was not found." -ForegroundColor Yellow
        $choice = Read-Host "Would you like to try and install it now? (Y/N)"
        if ($choice -eq 'y') {
            try {
                Write-HostInColor -Message "Installing 'SqlServer' module from the PowerShell Gallery..." -ForegroundColor Cyan
                Install-Module -Name SqlServer -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Write-HostInColor -Message "Module installed successfully. Please re-run the script." -ForegroundColor Green
            }
            catch {
                Write-HostInColor -Message "Failed to install the 'SqlServer' module. Please install it manually and try again." -ForegroundColor Red
                Write-HostInColor -Message "You can install it by running: Install-Module -Name SqlServer" -ForegroundColor Red
            }
        }
        else {
            Write-HostInColor -Message "Script cannot continue without the 'SqlServer' module." -ForegroundColor Red
        }
        return $false
    }
    
    if ($DebugMode) { Import-Module SqlServer -Verbose }
    else { Import-Module SqlServer -DisableNameChecking }
    
    try {
        # Force the loading of the main SMO assembly to ensure object types are recognized
        $modulePath = (Get-Module -Name SqlServer).ModuleBase
        $smoAssemblyPath = Join-Path -Path $modulePath -ChildPath "Microsoft.SqlServer.Smo.dll"
        Add-Type -Path $smoAssemblyPath
        if ($DebugMode) { Write-HostInColor -Message "[DEBUG] Explicitly loaded SMO assembly from: $smoAssemblyPath" -ForegroundColor Yellow }
    }
    catch {
        Write-HostInColor -Message "CRITICAL: Could not manually load the SMO assembly. Script may fail." -ForegroundColor Red
    }
    
    return $true
}

#endregion Helper Functions

# --- Script Main Body ---

if (-not (Test-SqlServerModule)) {
    return
}

if ([string]::IsNullOrWhiteSpace($SqlServerInstance)) {
    $SqlServerInstance = Read-Host "Please enter the SQL Server instance name (e.g., localhost\SQLEXPRESS)"
}
if ([string]::IsNullOrWhiteSpace($DatabaseName)) {
    $DatabaseName = Read-Host "Please enter the database name"
}

$timestamp = Get-Date -Format "yyyyMMddHHmmss"
$mainOutputDir = Join-Path -Path $env:TEMP -ChildPath "SqlScripts_${DatabaseName}_${timestamp}"

if ($DebugMode) { Write-HostInColor -Message "[DEBUG] Main output directory set to: $mainOutputDir" -ForegroundColor Yellow }

$objectFolders = @{
    "Tables"               = "Tables"
    "Views"                = "Views"
    "StoredProcedures"     = "StoredProcedures"
    "UserDefinedFunctions" = "Functions"
    "Users"                = "Users"
    "DatabaseRoles"        = "Roles"
    "Schemas"              = "Schemas"
    "DatabaseTriggers"     = "DatabaseTriggers"
}

try {
    Write-HostInColor -Message "Connecting to server '$SqlServerInstance'..." -ForegroundColor Cyan
    $server = New-Object Microsoft.SqlServer.Management.Smo.Server($SqlServerInstance)
    $server.ConnectionContext.TrustServerCertificate = $true

    if (-not $server.Databases.Contains($DatabaseName)) {
        throw "Database '$DatabaseName' not found on instance '$SqlServerInstance'."
    }

    $db = $server.Databases[$DatabaseName]
    Write-HostInColor -Message "Successfully connected to database '$DatabaseName'." -ForegroundColor Green

    Write-HostInColor -Message "Gathering database objects..." -ForegroundColor Cyan
    $objectsToScript = New-Object System.Collections.ArrayList

    $db.Tables | Where-Object { -not $_.IsSystemObject } | ForEach-Object { $null = $objectsToScript.Add($_) }
    $db.Views | Where-Object { -not $_.IsSystemObject } | ForEach-Object { $null = $objectsToScript.Add($_) }
    $db.StoredProcedures | Where-Object { -not $_.IsSystemObject } | ForEach-Object { $null = $objectsToScript.Add($_) }
    $db.UserDefinedFunctions | Where-Object { -not $_.IsSystemObject } | ForEach-Object { $null = $objectsToScript.Add($_) }
    $db.Users | Where-Object { -not $_.IsSystemObject } | ForEach-Object { $null = $objectsToScript.Add($_) }
    $db.Roles | ForEach-Object { $null = $objectsToScript.Add($_) }
    $db.Schemas | Where-Object { -not ($_.Name -in ('guest', 'INFORMATION_SCHEMA', 'sys', 'db_owner', 'db_accessadmin', 'db_securityadmin', 'db_ddladmin', 'db_backupoperator', 'db_datareader', 'db_datawriter', 'db_denydatareader', 'db_denydatawriter')) } | ForEach-Object { $null = $objectsToScript.Add($_) }
    $db.Triggers | Where-Object { -not $_.IsSystemObject } | ForEach-Object { $null = $objectsToScript.Add($_) }

    $totalObjectCount = $objectsToScript.Count
    if ($totalObjectCount -eq 0) {
        Write-HostInColor -Message "No scriptable objects found in database '$DatabaseName'." -ForegroundColor Yellow
        return
    }
    
    if ($DebugMode) { Write-HostInColor -Message "[DEBUG] Found $totalObjectCount objects to script." -ForegroundColor Yellow }

    $null = New-Item -Path $mainOutputDir -ItemType Directory -ErrorAction SilentlyContinue

    if ($DebugMode) { Write-HostInColor -Message "[DEBUG] Determining required folders..." -ForegroundColor Yellow }
    $foldersToCreate = $objectsToScript | ForEach-Object {
        $obj = $_
        if (($obj -is [Microsoft.SqlServer.Management.Smo.User]) -and $obj.IsOrphaned) { return } # Exclude orphaned users from folder creation
        if ($obj -is [Microsoft.SqlServer.Management.Smo.Table])                 { $objectFolders.Tables }
        elseif ($obj -is [Microsoft.SqlServer.Management.Smo.View])              { $objectFolders.Views }
        elseif ($obj -is [Microsoft.SqlServer.Management.Smo.StoredProcedure])   { $objectFolders.StoredProcedures }
        elseif ($obj -is [Microsoft.SqlServer.Management.Smo.UserDefinedFunction]){ $objectFolders.UserDefinedFunctions }
        elseif ($obj -is [Microsoft.SqlServer.Management.Smo.User])              { $objectFolders.Users }
        elseif ($obj -is [Microsoft.SqlServer.Management.Smo.DatabaseRole])      { $objectFolders.DatabaseRoles }
        elseif ($obj -is [Microsoft.SqlServer.Management.Smo.Schema])            { $objectFolders.Schemas }
        elseif ($obj -is [Microsoft.SqlServer.Management.Smo.Trigger])           { $objectFolders.DatabaseTriggers }
    } | Select-Object -Unique

    Write-HostInColor -Message "Creating required directories..." -ForegroundColor Cyan
    foreach ($folder in $foldersToCreate) {
        if (-not ([string]::IsNullOrWhiteSpace($folder))) {
            if ($DebugMode) { Write-HostInColor -Message "[DEBUG] Creating folder: $folder" -ForegroundColor Yellow }
            $null = New-Item -Path (Join-Path -Path $mainOutputDir -ChildPath $folder) -ItemType Directory
        }
    }

    $scripter = New-Object Microsoft.SqlServer.Management.Smo.Scripter($server)
    $scripter.Options.ScriptSchema = $true
    $scripter.Options.ScriptData = $false
    $scripter.Options.ToFileOnly = $true
    $scripter.Options.Encoding = [System.Text.Encoding]::Default
    $scripter.Options.DriAll = $true

    Write-HostInColor -Message "Starting script generation for $totalObjectCount objects..." -ForegroundColor Cyan

    for ($i = 0; $i -lt $totalObjectCount; $i++) {
        $obj = $objectsToScript[$i]
        
        try {
            # Pre-check for orphaned users, which cannot be scripted reliably.
            if (($obj -is [Microsoft.SqlServer.Management.Smo.User]) -and $obj.IsOrphaned) {
                Write-HostInColor -Message "Skipping orphaned user '$($obj.Name)'." -ForegroundColor Yellow
                continue
            }

            $targetFolder = ""
            if ($obj -is [Microsoft.SqlServer.Management.Smo.Table]) { $targetFolder = $objectFolders.Tables }
            elseif ($obj -is [Microsoft.SqlServer.Management.Smo.View]) { $targetFolder = $objectFolders.Views }
            elseif ($obj -is [Microsoft.SqlServer.Management.Smo.StoredProcedure]) { $targetFolder = $objectFolders.StoredProcedures }
            elseif ($obj -is [Microsoft.SqlServer.Management.Smo.UserDefinedFunction]) { $targetFolder = $objectFolders.UserDefinedFunctions }
            elseif ($obj -is [Microsoft.SqlServer.Management.Smo.User]) { $targetFolder = $objectFolders.Users }
            elseif ($obj -is [Microsoft.SqlServer.Management.Smo.DatabaseRole]) { $targetFolder = $objectFolders.DatabaseRoles }
            elseif ($obj -is [Microsoft.SqlServer.Management.Smo.Schema]) { $targetFolder = $objectFolders.Schemas }
            elseif ($obj -is [Microsoft.SqlServer.Management.Smo.Trigger]) { $targetFolder = $objectFolders.DatabaseTriggers }
            else { continue }
            
            $objName = $obj.Name
            $objSchema = $obj.PSObject.Properties.Match('Schema') | Select-Object -First 1
            $fileName = if ($objSchema) { "$($obj.Schema).$objName.sql" } else { "$objName.sql" }
            
            $filePath = Join-Path -Path (Join-Path -Path $mainOutputDir -ChildPath $targetFolder) -ChildPath $fileName
            $scripter.Options.FileName = $filePath

            if ($obj -is [Microsoft.SqlServer.Management.Smo.Table]) {
                $scripter.Options.Indexes = $true
                $scripter.Options.Triggers = $true
            }
            else {
                $scripter.Options.Indexes = $false
                $scripter.Options.Triggers = $false
            }
            
            $progressParams = @{
                Activity        = "Scripting Objects from '$DatabaseName'"
                Status          = "Processing [$($obj.GetType().Name)]: $fileName ($($i + 1) of $totalObjectCount)"
                PercentComplete = (($i + 1) / $totalObjectCount) * 100
            }
            Write-Progress @progressParams

            if ($DebugMode) { Write-HostInColor -Message "[DEBUG] Scripting '$fileName' to '$filePath'" -ForegroundColor Yellow }

            $scripter.Script(@($obj))
        }
        catch {
            Write-HostInColor -Message "Could not script object '$($obj.Name)'. Error: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    Write-Progress -Activity "Scripting Objects from '$DatabaseName'" -Completed

    Write-HostInColor -Message "--------------------------------------------------------" -ForegroundColor Cyan
    Write-HostInColor -Message "Script generation completed successfully!" -ForegroundColor Green
    Write-HostInColor -Message "Output located at: $mainOutputDir" -ForegroundColor Green
    
    Invoke-Item -Path $mainOutputDir
}
catch {
    Write-HostInColor -Message "An unexpected error occurred:" -ForegroundColor Red
    Write-HostInColor -Message $_.Exception.Message -ForegroundColor Red
}
finally {
    if ($server) {
        $server.ConnectionContext.Disconnect()
    }
}