# Test-SmoLoad.ps1
# This script is for diagnostic purposes only.

function Write-HostInColor {
    param([string]$Message, [string]$ForegroundColor)
    Write-Host -Object $Message -ForegroundColor $ForegroundColor
}

Write-HostInColor "--- PowerShell Environment Diagnostics ---" "Cyan"
$PSVersionTable
Write-Host "Is 64-bit Process: $([System.Environment]::Is64BitProcess)"
Write-Host "CLR Version: $([System.Runtime.InteropServices.RuntimeEnvironment]::GetSystemVersion())"
Write-HostInColor "-----------------------------------------" "Cyan"

Write-HostInColor "[INFO] Checking for 'SqlServer' module..." "White"
$module = Get-Module -Name SqlServer -ListAvailable | Select-Object -First 1
if (-not $module) {
    Write-HostInColor "[ERROR] 'SqlServer' module not found. Please install it and try again." "Red"
    return
}

$modulePath = $module.ModuleBase
Write-HostInColor "[INFO] Found 'SqlServer' module version $($module.Version) at: $modulePath" "White"
$smoAssemblyPath = Join-Path -Path $modulePath -ChildPath "Microsoft.SqlServer.Smo.dll"

if (-not (Test-Path -Path $smoAssemblyPath)) {
    Write-HostInColor "[ERROR] Could not find 'Microsoft.SqlServer.Smo.dll' at path: $smoAssemblyPath" "Red"
    return
}

Write-HostInColor "[INFO] Attempting to load SMO assembly: $smoAssemblyPath" "White"
$assembly = $null
try {
    # The -PassThru parameter will return the loaded assembly object on success.
    $assembly = Add-Type -Path $smoAssemblyPath -ErrorAction Stop -PassThru
    Write-HostInColor "[SUCCESS] 'Add-Type' command completed without a terminating error." "Green"
    Write-Host "Assembly Details: $($assembly.FullName)"
}
catch {
    Write-HostInColor "[FATAL] 'Add-Type' command failed with a terminating error." "Red"
    Write-HostInColor "Error: $($_.Exception.Message)" "Red"
    if ($_.Exception.InnerException) {
        Write-HostInColor "Inner Exception: $($_.Exception.InnerException.Message)" "Red"
    }
    return
}

Write-HostInColor "-----------------------------------------" "Cyan"
Write-HostInColor "[INFO] Verification Step: Checking if 'Role' type is available..." "White"

try {
    $typeName = 'Microsoft.SqlServer.Management.Smo.Role'
    $type = [System.Management.Automation.PSTypeName]$typeName
    
    if ($type.Type) {
        Write-HostInColor "[SUCCESS] The type '$typeName' was found successfully after loading." "Green"
    }
    else {
        Write-HostInColor "[FAILURE] The type '$typeName' IS NOT AVAILABLE, even though Add-Type seemed to succeed." "Red"
        
        # Additional Diagnostics
        Write-HostInColor "[INFO] Checking all loaded assemblies for 'Microsoft.SqlServer.Smo'..." "White"
        $loadedSmo = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.FullName -like 'Microsoft.SqlServer.Smo*' }
        
        if ($loadedSmo) {
            Write-HostInColor "Found the following SMO assemblies loaded in the AppDomain:" "Yellow"
            $loadedSmo | ForEach-Object { Write-Host " - $($_.FullName) from ($($_.Location))" }
        } else {
            Write-HostInColor "No assembly matching 'Microsoft.SqlServer.Smo*' was found in the current AppDomain." "Red"
        }
    }
}
catch {
    Write-HostInColor "[FAILURE] An error occurred while trying to verify the type '$typeName'." "Red"
    Write-HostInColor "Error: $($_.Exception.Message)" "Red"
}

Write-HostInColor "--- Diagnostics Complete ---" "Cyan"