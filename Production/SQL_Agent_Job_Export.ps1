<#
.SYNOPSIS
    Exports SQL Agent job scripts to a temporary folder, then briefly interacts with a Git repo to create a branch, move the files, commit, and push.

.DESCRIPTION
    This script is highly optimized to efficiently export SQL Agent jobs. It performs a fast and reliable
    connection test before starting the main workflow. It uses a direct SQL query to filter jobs
    by category on the server. It automatically trusts the server's SSL certificate.
    Optional prompts will time out after 20 seconds for automated execution. All console output is
    centralized through a logging function, and a -DebugMode switch is available for verbose output.

.PARAMETER ServerInstance
    The name of the target SQL Server instance. Will be prompted for if not provided.

.PARAMETER TargetDirectory
    The full path to the directory inside your Git repo where the SQL scripts should be saved. Will be prompted for if not provided.

.PARAMETER BranchName
    (Optional) The name of the new branch to create for the commit. If not provided, a name will be generated
    and the user will be prompted with a 20-second timeout to accept or override.

.PARAMETER JobCategoryFilter
    (Optional) The name of a specific job category to export. If this parameter is omitted, all jobs will be exported.

.PARAMETER DebugMode
    (Optional) Enables detailed diagnostic output to the console for troubleshooting.

.EXAMPLE
    PS C:\> .\Export-SqlJobs-V14.0.ps1 -ServerInstance "PRODDB01" -TargetDirectory "C:\Git\jobs" -DebugMode

    Description:
    The script will first perform a reliable check to see if it can connect to "PRODDB01". If successful,
    it will proceed, showing verbose debugging messages along the way. If not, it will stop with a connection error.

.NOTES
    Version:     14.0
    Author:      Bereket W., Brennan W., and Gemini
    Last Modified: 2025-07-01
    Requires:    1. Windows PowerShell 5.1 or later.
                 2. The 'SqlServer' PowerShell module.
                 3. The Git command-line client (git.exe) must be installed and in your system's PATH.
    Change Log:
    - v14.0: Major refactor. Introduced a centralized Write-Log function and a -DebugMode switch to align with
             the Optimus.ps1 development standard. Replaced all Write-Host calls.
#>
[CmdletBinding()]
param (
    [string]$ServerInstance,
    [string]$TargetDirectory,
    [string]$BranchName,
    [string]$JobCategoryFilter,
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

#region Helper Functions
function Read-HostWithTimeout {
    param(
        [string]$Prompt,
        [int]$TimeoutSeconds
    )
    Write-Log -Message $Prompt -Level 'PROMPT' -NoNewLine
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $inputBuffer = New-Object -TypeName System.Text.StringBuilder

    while ($stopwatch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true) 
            if ($key.Key -eq 'Enter') {
                Write-Host # Move to the next line after user presses Enter
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
    Write-Host # Move to the next line after timeout
    return $null
}
#endregion

function Export-SqlJobScript {
    [CmdletBinding()]
    param(
        [string]$ServerInstance,
        [string]$TargetDirectory,
        [string]$BranchName,
        [string]$JobCategoryFilter,
        [switch]$DebugMode
    )

    if ($DebugMode) { Write-Log -Message "Running in Debug Mode." -Level 'DEBUG' }
    
    #region Interactive Parameter Prompts
    if ([string]::IsNullOrWhiteSpace($ServerInstance)) { $ServerInstance = Read-Host -Prompt "Enter the SQL Server instance name" }
    if ([string]::IsNullOrWhiteSpace($TargetDirectory)) { $TargetDirectory = Read-Host -Prompt "Enter the target directory for the SQL scripts (must be inside a Git repo)" }

    if ([string]::IsNullOrWhiteSpace($BranchName)) {
        $sanitizedServerName = $ServerInstance -replace '[\\.]', '_'
        $generatedBranchName = "{0}_Job_Export_{1}" -f $sanitizedServerName, (Get-Date -Format 'yyyyMMdd')
        $promptMessage = "Proceed with branch name = '$generatedBranchName'? (Press Enter to accept, or type a new name). `nThe default will be used in 20 seconds: "
        $userInput = Read-HostWithTimeout -Prompt $promptMessage -TimeoutSeconds 20
        
        if ($null -eq $userInput) {
            Write-Log -Message "Timeout reached. Using default branch name." -Level 'WARN'
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
             Write-Log -Message "Timeout reached. Exporting all jobs." -Level 'WARN'
        }
    }
    #endregion

    # --- Pre-execution connection test ---
    try {
        Write-Log -Message "Verifying connection to '$ServerInstance'..." -Level 'INFO'
        $conn = New-Object System.Data.SqlClient.SqlConnection
        # Use a short timeout for a quick and reliable test. Assumes integrated (Windows) security.
        $conn.ConnectionString = "Server=$ServerInstance;Integrated Security=True;Connection Timeout=5;TrustServerCertificate=True"
        $conn.Open()
        Write-Log -Message "Connection test successful, closing test connection." -Level 'DEBUG'
        $conn.Close()
        Write-Log -Message "Connection successful." -Level 'SUCCESS'
    }
    catch {
        Write-Log -Message "Error: Could not connect to server '$ServerInstance'. Please check the name and network connectivity." -Level 'ERROR'
        Write-Log -Message "Details: $($_.Exception.GetBaseException().Message)" -Level 'ERROR'
        return # Exit the function
    }
    
    $initialBranch = $null
    $gitRoot = $null
    $tempDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ([System.Guid]::NewGuid().ToString())

    try {
        Write-Log -Message "Using branch name: '$BranchName'" -Level 'INFO'
        if (-not [string]::IsNullOrWhiteSpace($JobCategoryFilter)){
            Write-Log -Message "Using job category filter: '$JobCategoryFilter'" -Level 'INFO'
        }

        # --- PRE-FLIGHT CHECKS ---
        Write-Log -Message "Performing pre-flight checks..." -Level 'INFO'
        if (-not (Get-Command git -ErrorAction SilentlyContinue)) { throw "Git command-line client not found. Please ensure Git is installed and in your system's PATH." }
        if (-not (Get-Module -Name SqlServer -ListAvailable)) {
            Write-Log -Message "SqlServer module not found. Attempting to install from PSGallery..." -Level 'WARN'
            Install-Module -Name SqlServer -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
        }
        Import-Module -Name SqlServer
        Write-Log -Message "Pre-flight checks passed." -Level 'DEBUG'
        
        # --- STEP 1: SQL EXPORT TO TEMPORARY DIRECTORY ---
        Write-Log -Message "Preparing to export SQL jobs..." -Level 'INFO'
        New-Item -Path $TempDir -ItemType Directory | Out-Null
        Write-Log -Message "Created temporary directory: $tempDir" -Level 'DEBUG'
        
        $server = New-Object Microsoft.SqlServer.Management.Smo.Server($ServerInstance)
        $server.ConnectionContext.TrustServerCertificate = $true
        
        $jobs = @() 
        if (-not [string]::IsNullOrWhiteSpace($JobCategoryFilter)) {
            Write-Log -Message "Retrieving jobs from category '$JobCategoryFilter'..." -Level 'INFO'
            $escapedCategory = $JobCategoryFilter.Replace("'", "''")
            $query = "SELECT j.name FROM msdb.dbo.sysjobs j JOIN msdb.dbo.syscategories c ON j.category_id = c.category_id WHERE c.name = '$escapedCategory'"
            Write-Log -Message "Executing job filter query: $query" -Level 'DEBUG'
            
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
            Write-Log -Message "Retrieving all jobs from server... (This may take a moment on servers with many jobs)" -Level 'INFO'
            $jobs = $server.JobServer.Jobs
        }

        if ($jobs.Count -eq 0) {
            $warningMsg = "No SQL jobs found on '$ServerInstance'"
            if (-not [string]::IsNullOrWhiteSpace($JobCategoryFilter)) {
                $warningMsg += " (with the category filter '$JobCategoryFilter'). Halting process."
            } else {
                 $warningMsg += ". Halting process."
            }
            Write-Log -Message $warningMsg -Level 'WARN'
            return
        }
        
        Write-Log -Message "Found $($jobs.Count) jobs. Beginning export..." -Level 'INFO'
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
            Write-Log -Message "Scripting job '$($job.Name)' to '$fileName'" -Level 'DEBUG'
            $jobScriptContent = $job.Script()
            [System.IO.File]::WriteAllLines($fileName, $jobScriptContent, $utf8WithoutBom)
        }
        Write-Progress -Activity "Exporting SQL Agent Jobs from $ServerInstance" -Completed
        Write-Log -Message "Temporary export complete." -Level 'SUCCESS'
        
        # --- STEP 2: GIT OPERATIONS ---
        $currentPath = $TargetDirectory
        while ($currentPath -ne $null -and $currentPath -ne "") {
            if (Test-Path -Path (Join-Path $currentPath ".git")) { $gitRoot = $currentPath; break }
            $currentPath = Split-Path -Parent $currentPath
        }
        if ([string]::IsNullOrEmpty($gitRoot)) { throw "Could not find a Git repository root (.git folder) in any parent directory of '$TargetDirectory'." }
        Write-Log -Message "Git repository root detected at: $gitRoot" -Level 'INFO'

        Write-Log -Message "Starting fast Git operations..." -Level 'INFO'
        
        Push-Location -Path $gitRoot
        Write-Log -Message "Changed directory to Git root." -Level 'DEBUG'
        
        $initialBranch = git rev-parse --abbrev-ref HEAD
        Write-Log -Message "Original Git branch is '$initialBranch'." -Level 'DEBUG'
        
        git checkout master
        if ($LASTEXITCODE -ne 0) { throw "Failed to checkout 'master' branch." }
        git pull
        if ($LASTEXITCODE -ne 0) { throw "Failed to pull latest changes for 'master'." }
        git checkout -b $BranchName
        if ($LASTEXITCODE -ne 0) { throw "Failed to create new Git branch '$BranchName'." }
        
        Write-Log -Message "Moving exported files into the repository at '$TargetDirectory'..." -Level 'INFO'
        if (-not (Test-Path -Path $TargetDirectory)) { New-Item -Path $TargetDirectory -ItemType Directory | Out-Null }
        Move-Item -Path "$tempDir\*" -Destination $TargetDirectory -Force
        
        Write-Log -Message "Staging, committing, and pushing changes..." -Level 'INFO'
        git add .
        if ($LASTEXITCODE -ne 0) { throw "Failed to stage changes with 'git add'." }
        git commit -m "$BranchName"
        if ($LASTEXITCODE -ne 0) { throw "Failed to commit changes." }
        git push --set-upstream origin $BranchName
        if ($LASTEXITCODE -ne 0) { throw "Failed to push branch to the remote repository." }
        
        Write-Log -Message "Successfully pushed branch '$BranchName' to the remote repository." -Level 'SUCCESS'
    }
    catch {
        Write-Log -Message "An error occurred: $($_.Exception.Message)" -Level 'ERROR'
        Write-Progress -Activity "Exporting SQL Agent Jobs from $ServerInstance" -Completed -ErrorAction SilentlyContinue
    }
    finally {
        # --- CLEANUP ---
        Write-Log -Message "Cleaning up..." -Level 'INFO'
        if ($gitRoot) {
            if ($initialBranch) {
                Write-Log -Message "Returning to original branch '$initialBranch'..." -Level 'INFO'
                git checkout $initialBranch
            }
            Pop-Location
            Write-Log -Message "Returned to original directory." -Level 'DEBUG'
        }
        if (Test-Path -Path $tempDir) {
            Write-Log -Message "Removing temporary directory..." -Level 'INFO'
            Remove-Item -Path $tempDir -Recurse -Force
        }
        Write-Log -Message "Script finished." -Level 'SUCCESS'
    }
}

Export-SqlJobScript @PSBoundParameters