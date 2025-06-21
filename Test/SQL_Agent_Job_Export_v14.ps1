<#
.SYNOPSIS
    Exports SQL Agent job scripts to a specified directory within a Git repository, creates a new branch, and commits the scripts.

.DESCRIPTION
    This script efficiently exports SQL Agent jobs from a specified SQL Server instance. It performs a reliable connection test before executing the main workflow.
    It uses a direct SQL query to filter jobs by category on the server for improved performance. The script is designed to be run either interactively, prompting for necessary information, or non-interactively in automated pipelines.
    It handles Git operations to create a new branch, move the exported files, commit them, and push the branch to the remote repository.

.PARAMETER ServerInstance
    The name of the target SQL Server instance. This parameter is mandatory. If run interactively, the user will be prompted for it.

.PARAMETER TargetDirectory
    The full path to the directory inside your Git repository where the SQL scripts should be saved. This path must exist within a valid Git repo. This parameter is mandatory.

.PARAMETER BranchName
    The name of the new branch to create for the commit. If not provided, a name will be generated automatically (e.g., 'SERVER_Job_Export_yyyyMMdd'), and the user will be prompted with a 20-second timeout to accept or override it.

.PARAMETER JobCategoryFilter
    The name of a specific job category to export. If this parameter is omitted, all jobs from all categories will be exported. The user will be prompted with a 20-second timeout to provide a value if one is not supplied.

.EXAMPLE
    PS C:\> .\SQL_Agent_Job_Export.ps1 -ServerInstance "PRODDB01\SQL2019" -TargetDirectory "C:\Git\MyRepo\SqlJobs"

    Description:
    Connects to the "PRODDB01\SQL2019" instance, exports all SQL Agent jobs to a temporary location, then moves them to "C:\Git\MyRepo\SqlJobs". It creates a new branch (e.g., 'PRODDB01_SQL2019_Job_Export_20250612'), commits the files, and pushes the new branch to the remote.

.EXAMPLE
    PS C:\> .\SQL_Agent_Job_Export.ps1 -ServerInstance "TestDB01" -TargetDirectory "C:\Git\MyRepo\SqlJobs" -JobCategoryFilter "Database Maintenance" -BranchName "feature/update-maintenance-jobs"

    Description:
    Exports only the jobs in the "Database Maintenance" category from "TestDB01" and commits them to a new branch named "feature/update-maintenance-jobs".

.NOTES
    Version:     14.2
    Author:      Bereket W., Brennan W., Gemini
    Last Modified: 2025-06-12
    Requires:    
    1. PowerShell 5.1 or later.
    2. The 'SqlServer' PowerShell module must be installed or installable from the PSGallery.
    3. The Git command-line client (git.exe) must be installed and in your system's PATH.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Enter the SQL Server instance name.")]
    [string]$ServerInstance,

    [Parameter(Mandatory = $true, HelpMessage = "Enter the target directory for the SQL scripts (must be inside a Git repo).")]
    [ValidateScript({
        $path = $_
        $isGitRepo = $false
        while ($path -ne $null -and $path -ne "") {
            if (Test-Path -Path (Join-Path $path ".git")) {
                $isGitRepo = $true
                break
            }
            $path = Split-Path -Parent $path
        }
        if (-not $isGitRepo) {
            throw "The specified TargetDirectory '$_' is not inside a Git repository. A '.git' folder could not be found in any parent directory."
        }
        return $true
    })]
    [string]$TargetDirectory,

    [Parameter(HelpMessage = "The name of the new branch to create for the commit.")]
    [string]$BranchName,

    [Parameter(HelpMessage = "The name of a specific job category to export. If omitted, all jobs will be exported.")]
    [string]$JobCategoryFilter
)

#region Helper Functions
function Read-HostWithTimeout {
    param(
        [string]$Prompt,
        [int]$TimeoutSeconds
    )
    Write-Host -NoNewline $Prompt -ForegroundColor White
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $inputBuffer = New-Object -TypeName System.Text.StringBuilder

    while ($stopwatch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            if ($key.Key -eq 'Enter') {
                Write-Host
                return $inputBuffer.ToString()
            }
            elseif ($key.Key -eq 'Backspace') {
                if ($inputBuffer.Length -gt 0) {
                    $inputBuffer.Remove($inputBuffer.Length - 1, 1) | Out-Null
                    Write-Host -NoNewline "`b `b"
                }
            }
            else {
                $inputBuffer.Append($key.KeyChar) | Out-Null
                Write-Host -NoNewline $key.KeyChar -ForegroundColor White
            }
        }
        Start-Sleep -Milliseconds 100
    }
    Write-Host
    return $null
}
#endregion

#region Interactive Parameter Prompts for Optional Params
if ([string]::IsNullOrWhiteSpace($BranchName)) {
    # Sanitize server name for use in a branch name.
    $sanitizedServerName = $ServerInstance -replace '[\\.\s]', '_'
    $generatedBranchName = "{0}_Job_Export_{1}" -f $sanitizedServerName, (Get-Date -Format 'yyyyMMdd')
    $promptMessage = "Enter a branch name or press Enter to accept the default: '$generatedBranchName'.`n(The default will be used in 20 seconds): "
    $userInput = Read-HostWithTimeout -Prompt $promptMessage -TimeoutSeconds 20

    if ($null -eq $userInput) {
        Write-Host "`nTimeout reached. Using default branch name." -ForegroundColor Yellow
        $BranchName = $generatedBranchName
    }
    elseif ([string]::IsNullOrWhiteSpace($userInput)) {
        $BranchName = $generatedBranchName
    }
    else {
        $BranchName = $userInput
    }
}

if ([string]::IsNullOrWhiteSpace($JobCategoryFilter)) {
    $promptMessage = "Enter job category to filter by (optional, press Enter to skip).`n(This will be skipped in 20 seconds): "
    $userInput = Read-HostWithTimeout -Prompt $promptMessage -TimeoutSeconds 20

    if ($null -eq $userInput) {
        # Timeout condition
        Write-Host "`nTimeout reached. Exporting all jobs." -ForegroundColor Yellow
        # $JobCategoryFilter remains empty, which is the desired outcome for exporting all jobs.
    }
    elseif ([string]::IsNullOrWhiteSpace($userInput)) {
        # User pressed Enter to skip, so no action is taken.
        # $JobCategoryFilter remains empty, which is the desired outcome for exporting all jobs.
    }
    else {
        # User provided a category name.
        $JobCategoryFilter = $userInput
    }
}
#endregion

# --- Initialize script-level variables ---
$initialBranch = $null
$gitRoot = $null
# Create a unique temporary directory for the export
$tempDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ([System.Guid]::NewGuid().ToString())


# --- Main script execution block with error handling ---
try {
    Write-Host "Starting SQL Agent Job Export process." -ForegroundColor Cyan
    Write-Host "Using Branch Name: '$BranchName'" -ForegroundColor Cyan
    if (-not [string]::IsNullOrWhiteSpace($JobCategoryFilter)) {
        Write-Host "Using Job Category Filter: '$JobCategoryFilter'" -ForegroundColor Cyan
    }

    # --- PRE-FLIGHT CHECKS ---
    Write-Host "Performing pre-flight checks..." -ForegroundColor Cyan
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
        throw "Git command-line client not found. Please ensure Git is installed and in your system's PATH."
    }
    if (-not (Get-Module -Name SqlServer -ListAvailable)) {
        Write-Host "SqlServer module not found. Attempting to install from PSGallery..." -ForegroundColor Yellow
        if ($PSCmdlet.ShouldProcess("SqlServer Module", "Install from PSGallery")) {
            Install-Module -Name SqlServer -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
        }
        else {
            throw "SqlServer module is required. Aborting script as per user request."
        }
    }
    Import-Module -Name SqlServer

    # --- Pre-execution connection test ---
    Write-Host "Verifying connection to '$ServerInstance'..." -ForegroundColor Cyan
    $conn = New-Object System.Data.SqlClient.SqlConnection
    # Use a short timeout for a quick and reliable test. Assumes integrated (Windows) security.
    $conn.ConnectionString = "Server=$ServerInstance;Integrated Security=True;Connection Timeout=5;TrustServerCertificate=True"
    try {
        $conn.Open()
        Write-Host "Connection successful." -ForegroundColor Green
    }
    catch {
        # Throw a terminating error if connection fails
        throw "Could not connect to server '$ServerInstance'. Please check the name and network connectivity. Details: $($_.Exception.GetBaseException().Message)"
    }
    finally {
        $conn.Close()
    }

    # --- STEP 1: SQL EXPORT TO TEMPORARY DIRECTORY ---
    Write-Host "Preparing to export SQL jobs..." -ForegroundColor Cyan
    New-Item -Path $TempDir -ItemType Directory | Out-Null

    $server = New-Object Microsoft.SqlServer.Management.Smo.Server($ServerInstance)
    $server.ConnectionContext.TrustServerCertificate = $true

    $jobs = @()
    if (-not [string]::IsNullOrWhiteSpace($JobCategoryFilter)) {
        Write-Host "Retrieving jobs from category '$JobCategoryFilter'..." -ForegroundColor Cyan
        $escapedCategory = $JobCategoryFilter.Replace("'", "''")
        $query = "SELECT j.name FROM msdb.dbo.sysjobs j JOIN msdb.dbo.syscategories c ON j.category_id = c.category_id WHERE c.name = '$escapedCategory'"

        $invokeSqlParams = @{
            ServerInstance = $ServerInstance
            Query = $query
            ErrorAction = 'Stop'
            TrustServerCertificate = $true
        }
        $jobNames = Invoke-Sqlcmd @invokeSqlParams
        foreach ($jobName in $jobNames) {
            $jobs += $server.JobServer.Jobs[$jobName.name]
        }
    }
    else {
        Write-Host "Retrieving all jobs from server... (This may take a moment)" -ForegroundColor Cyan
        $jobs = $server.JobServer.Jobs
    }

    if ($jobs.Count -eq 0) {
        $warningMessage = "No SQL jobs found on '$ServerInstance'"
        if (-not [string]::IsNullOrWhiteSpace($JobCategoryFilter)) {
            $warningMessage += " with the category filter '$JobCategoryFilter'."
        }
        $warningMessage += " Halting process."
        Write-Warning -Message $warningMessage
        # Exit gracefully if no jobs are found
        return
    }

    Write-Host "Found $($jobs.Count) jobs. Beginning export to temporary directory..." -ForegroundColor Cyan
    $utf8WithoutBom = New-Object System.Text.UTF8Encoding($False)
    $totalJobs = $jobs.Count
    $currentJobIndex = 0

    foreach ($job in $jobs) {
        $currentJobIndex++
        $percentComplete = ($currentJobIndex / $totalJobs) * 100
        $statusMessage = "Processing job $currentJobIndex of ${totalJobs}: $($job.Name)"
        Write-Progress -Activity "Exporting SQL Agent Jobs from $ServerInstance" -Status $statusMessage -PercentComplete $percentComplete

        # Replace characters that are invalid in file names
        $safeJobName = $job.Name -replace '[\\/:"*?<>|]', '_'
        $fileName = Join-Path -Path $tempDir -ChildPath "$($safeJobName).sql"
        $jobScriptContent = $job.Script()
        [System.IO.File]::WriteAllLines($fileName, $jobScriptContent, $utf8WithoutBom)
    }
    Write-Progress -Activity "Exporting SQL Agent Jobs from $ServerInstance" -Completed
    Write-Host "Temporary export complete." -ForegroundColor Green

    # --- STEP 2: GIT OPERATIONS ---
    # Find the git repository root from the target directory
    $currentPath = $TargetDirectory
    while ($currentPath -ne $null -and $currentPath -ne "") {
        if (Test-Path -Path (Join-Path $currentPath ".git")) { $gitRoot = $currentPath; break }
        $currentPath = Split-Path -Parent $currentPath
    }
    # This check is redundant due to ValidateScript, but kept as a safeguard.
    if ([string]::IsNullOrEmpty($gitRoot)) { throw "Could not find a Git repository root from '$TargetDirectory'." }

    Write-Host "Git repository root detected at: $gitRoot" -ForegroundColor Cyan
    Write-Host "Starting Git operations..." -ForegroundColor Cyan

    Push-Location -Path $gitRoot

    # Store the current branch name for cleanup
    $initialBranch = git rev-parse --abbrev-ref HEAD
    if ($LASTEXITCODE -ne 0) { throw "Failed to get current branch name. Is '$gitRoot' a valid Git repository?" }

    # Checkout master/main, pull, then create the new branch
    git checkout master
    if ($LASTEXITCODE -ne 0) { throw "Failed to checkout 'master' branch." }
    git pull
    if ($LASTEXITCODE -ne 0) { throw "Failed to pull latest changes for 'master'." }
    git checkout -b $BranchName
    if ($LASTEXITCODE -ne 0) { throw "Failed to create new Git branch '$BranchName'." }

    Write-Host "Moving exported files into the repository at '$TargetDirectory'..." -ForegroundColor Cyan
    if (-not (Test-Path -Path $TargetDirectory)) {
        New-Item -Path $TargetDirectory -ItemType Directory | Out-Null
    }
    Move-Item -Path "$tempDir\*" -Destination $TargetDirectory -Force

    # Stage, Commit, and Push
    Write-Host "Staging, committing, and pushing changes..." -ForegroundColor Cyan
    git add .
    if ($LASTEXITCODE -ne 0) { throw "Failed to stage changes with 'git add'." }
    git commit -m "Exported SQL Agent Jobs from $ServerInstance"
    if ($LASTEXITCODE -ne 0) { throw "Failed to commit changes. There may be no changes to commit." }
    git push --set-upstream origin $BranchName
    if ($LASTEXITCODE -ne 0) { throw "Failed to push branch to the remote repository." }

    Write-Host "Successfully pushed branch '$BranchName' to the remote repository." -ForegroundColor Green
}
catch {
    # Catch any terminating error from the 'try' block
    Write-Error -Message "A critical error occurred: $($_.Exception.Message)" -ErrorAction Continue
    Write-Progress -Activity "Exporting SQL Agent Jobs from $ServerInstance" -Completed -ErrorAction SilentlyContinue
}
finally {
    # --- CLEANUP ---
    Write-Host "Performing cleanup..." -ForegroundColor Cyan
    if ($gitRoot) {
        if ($initialBranch -and (git rev-parse --abbrev-ref HEAD) -ne $initialBranch) {
            Write-Host "Returning to original branch '$initialBranch'..." -ForegroundColor Cyan
            git checkout $initialBranch
        }
        Pop-Location
    }
    if (Test-Path -Path $tempDir) {
        Write-Host "Removing temporary directory..." -ForegroundColor Cyan
        Remove-Item -Path $tempDir -Recurse -Force
    }
    Write-Host "Script finished." -ForegroundColor Green
}