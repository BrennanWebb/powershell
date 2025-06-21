<#
.SYNOPSIS
    Exports SQL Agent job scripts to a temporary folder, then briefly interacts with a Git repo to create a branch, move the files, commit, and push.

.DESCRIPTION
    This script is highly optimized to efficiently export SQL Agent jobs. It performs a fast and reliable
    connection test before starting the main workflow. It uses a direct SQL query to filter jobs
    by category on the server. It automatically trusts the server's SSL certificate.
    Optional prompts will time out after 20 seconds for automated execution.

.PARAMETER ServerInstance
    The name of the target SQL Server instance. Will be prompted for if not provided.

.PARAMETER TargetDirectory
    The full path to the directory inside your Git repo where the SQL scripts should be saved. Will be prompted for if not provided.

.PARAMETER BranchName
    (Optional) The name of the new branch to create for the commit. If not provided, a name will be generated
    and the user will be prompted with a 20-second timeout to accept or override.

.PARAMETER JobCategoryFilter
    (Optional) The name of a specific job category to export. If this parameter is omitted, all jobs will be exported.

.EXAMPLE
    PS C:\> .\Export-SqlJobs-V13.1.ps1 -ServerInstance "PRODDB01" -TargetDirectory "C:\Git\jobs"

    Description:
    The script will first perform a reliable check to see if it can connect to "PRODDB01". If successful,
    it will proceed. If not, it will stop with a connection error.

.NOTES
    Version:     13.1
    Author:      Bereket W., Brennan W., and Gemini
    Last Modified: 2025-06-12
    Requires:    1. Windows PowerShell 5.1 or later.
                 2. The 'SqlServer' PowerShell module.
                 3. The Git command-line client (git.exe) must be installed and in your system's PATH.
#>
[CmdletBinding()]
param (
    [string]$ServerInstance,
    [string]$TargetDirectory,
    [string]$BranchName,
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

function Export-SqlJobScript {
    [CmdletBinding()]
    param(
        [string]$ServerInstance,
        [string]$TargetDirectory,
        [string]$BranchName,
        [string]$JobCategoryFilter
    )

    #region Interactive Parameter Prompts
    if ([string]::IsNullOrWhiteSpace($ServerInstance)) { $ServerInstance = Read-Host -Prompt "Enter the SQL Server instance name" }
    if ([string]::IsNullOrWhiteSpace($TargetDirectory)) { $TargetDirectory = Read-Host -Prompt "Enter the target directory for the SQL scripts (must be inside a Git repo)" }

    if ([string]::IsNullOrWhiteSpace($BranchName)) {
        $sanitizedServerName = $ServerInstance -replace '[\\.]', '_'
        $generatedBranchName = "{0}_Job_Export_{1}" -f $sanitizedServerName, (Get-Date -Format 'yyyyMMdd')
        $promptMessage = "Proceed with branch name = '$generatedBranchName'? (Press Enter to accept, or type a new name). `nThe default will be used in 20 seconds: "
        $userInput = Read-HostWithTimeout -Prompt $promptMessage -TimeoutSeconds 20
        
        if ($null -eq $userInput) {
            Write-Host "Timeout reached. Using default branch name." -ForegroundColor Yellow
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
        $promptMessage = "Enter job category to filter by (optional). `nThis will be skipped in 20 seconds: "
        $userInput = Read-HostWithTimeout -Prompt $promptMessage -TimeoutSeconds 20
        if ($null -ne $userInput) {
            $JobCategoryFilter = $userInput
        } else {
             Write-Host "Timeout reached. Exporting all jobs." -ForegroundColor Yellow
        }
    }
    #endregion

    # --- Pre-execution connection test ---
    try {
        Write-Host "Verifying connection to '$ServerInstance'..." -ForegroundColor Cyan
        $conn = New-Object System.Data.SqlClient.SqlConnection
        # Use a short timeout for a quick and reliable test. Assumes integrated (Windows) security.
        $conn.ConnectionString = "Server=$ServerInstance;Integrated Security=True;Connection Timeout=5;TrustServerCertificate=True"
        $conn.Open()
        $conn.Close()
        Write-Host "Connection successful." -ForegroundColor Green
    }
    catch {
        Write-Host -Object "Error: Could not connect to server '$ServerInstance'. Please check the name and network connectivity." -ForegroundColor Red
        Write-Host -Object "Details: $($_.Exception.GetBaseException().Message)" -ForegroundColor Red
        return # Exit the function
    }
    
    $initialBranch = $null
    $gitRoot = $null
    $tempDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ([System.Guid]::NewGuid().ToString())

    try {
        Write-Host "Using branch name: '$BranchName'" -ForegroundColor Cyan
        if (-not [string]::IsNullOrWhiteSpace($JobCategoryFilter)){
            Write-Host "Using job category filter: '$JobCategoryFilter'" -ForegroundColor Cyan
        }

        # --- PRE-FLIGHT CHECKS ---
        Write-Host "Performing pre-flight checks..." -ForegroundColor Cyan
        if (-not (Get-Command git -ErrorAction SilentlyContinue)) { throw "Git command-line client not found. Please ensure Git is installed and in your system's PATH." }
        if (-not (Get-Module -Name SqlServer -ListAvailable)) {
            Write-Host "SqlServer module not found. Attempting to install from PSGallery..." -ForegroundColor Yellow
            Install-Module -Name SqlServer -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
        }
        Import-Module -Name SqlServer
        
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
            
            try {
                $invokeSqlParams = @{
                    ServerInstance = $ServerInstance; Query = $query;
                    ErrorAction = 'Stop'; TrustServerCertificate = $true
                }
                $jobNames = Invoke-Sqlcmd @invokeSqlParams
                foreach ($jobName in $jobNames) {
                    $jobs += $server.JobServer.Jobs[$jobName.name]
                }
            }
            catch {
                throw "Failed to query for jobs. Error: $($_.Exception.Message)"
            }
        }
        else {
            Write-Host "Retrieving all jobs from server... (This may take a moment on servers with many jobs)" -ForegroundColor Cyan
            $jobs = $server.JobServer.Jobs
        }

        if ($jobs.Count -eq 0) {
            Write-Host "No SQL jobs found on '$ServerInstance'" -ForegroundColor Yellow
            if (-not [string]::IsNullOrWhiteSpace($JobCategoryFilter)) {
                Write-Host " (with the category filter '$JobCategoryFilter'). Halting process." -ForegroundColor Yellow
            } else {
                 Write-Host ". Halting process." -ForegroundColor Yellow
            }
            return
        }
        
        Write-Host "Found $($jobs.Count) jobs. Beginning export..." -ForegroundColor Cyan
        $utf8WithoutBom = New-Object System.Text.UTF8Encoding($False)
        $totalJobs = $jobs.Count
        $currentJobIndex = 0

        foreach ($job in $jobs) {
            $currentJobIndex++
            $percentComplete = ($currentJobIndex / $totalJobs) * 100
            $statusMessage = "Processing job $currentJobIndex of ${totalJobs}: $($job.Name)"
            Write-Progress -Activity "Exporting SQL Agent Jobs from $ServerInstance" -Status $statusMessage -PercentComplete $percentComplete

            $safeJobName = $job.Name -replace '[\\/:"*?<>|]', '_'
            $fileName = Join-Path -Path $tempDir -ChildPath "$($safeJobName).sql"
            $jobScriptContent = $job.Script()
            [System.IO.File]::WriteAllLines($fileName, $jobScriptContent, $utf8WithoutBom)
        }
        Write-Progress -Activity "Exporting SQL Agent Jobs from $ServerInstance" -Completed
        Write-Host "Temporary export complete." -ForegroundColor Green
        
        # --- STEP 2: GIT OPERATIONS ---
        $currentPath = $TargetDirectory
        while ($currentPath -ne $null -and $currentPath -ne "") {
            if (Test-Path -Path (Join-Path $currentPath ".git")) { $gitRoot = $currentPath; break }
            $currentPath = Split-Path -Parent $currentPath
        }
        if ([string]::IsNullOrEmpty($gitRoot)) { throw "Could not find a Git repository root (.git folder) in any parent directory of '$TargetDirectory'." }
        Write-Host "Git repository root detected at: $gitRoot" -ForegroundColor Cyan

        Write-Host "Starting fast Git operations..." -ForegroundColor Cyan
        
        Push-Location -Path $gitRoot
        
        $initialBranch = git rev-parse --abbrev-ref HEAD
        
        git checkout master
        if ($LASTEXITCODE -ne 0) { throw "Failed to checkout 'master' branch." }
        git pull
        if ($LASTEXITCODE -ne 0) { throw "Failed to pull latest changes for 'master'." }
        git checkout -b $BranchName
        if ($LASTEXITCODE -ne 0) { throw "Failed to create new Git branch '$BranchName'." }
        
        Write-Host "Moving exported files into the repository at '$TargetDirectory'..." -ForegroundColor Cyan
        if (-not (Test-Path -Path $TargetDirectory)) { New-Item -Path $TargetDirectory -ItemType Directory | Out-Null }
        Move-Item -Path "$tempDir\*" -Destination $TargetDirectory -Force
        
        Write-Host "Staging, committing, and pushing changes..." -ForegroundColor Cyan
        git add .
        if ($LASTEXITCODE -ne 0) { throw "Failed to stage changes with 'git add'." }
        git commit -m "$BranchName"
        if ($LASTEXITCODE -ne 0) { throw "Failed to commit changes." }
        git push --set-upstream origin $BranchName
        if ($LASTEXITCODE -ne 0) { throw "Failed to push branch to the remote repository." }
        
        Write-Host "Successfully pushed branch '$BranchName' to the remote repository." -ForegroundColor Green
    }
    catch {
        Write-Host -Object "An error occurred: $($_.Exception.Message)" -ForegroundColor Red
        Write-Progress -Activity "Exporting SQL Agent Jobs from $ServerInstance" -Completed -ErrorAction SilentlyContinue
    }
    finally {
        # --- CLEANUP ---
        Write-Host "Cleaning up..." -ForegroundColor Cyan
        if ($gitRoot) {
            if ($initialBranch) {
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
}

Export-SqlJobScript @PSBoundParameters