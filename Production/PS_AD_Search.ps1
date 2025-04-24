# PS_AD_Search - Combined AD User and Group Search Tool
# Version: 1.0.9
# Description: Interactive script to search Active Directory for user group memberships or group members across multiple domains. Supports CLI parameters and help display. Now includes nested group resolution and labels.

# Define command-line parameters
param (
    [Alias("h", "help")][switch]$Help,
    [string]$UserName,
    [string]$GroupName
)

# Global variable to track auto-export preference
$global:AutoExport = $false

# Display help content if -Help is passed
if ($Help) {
    Write-Host @"
PS_AD_Search - Combined AD User and Group Search Tool
Version: 1.0.9

USAGE:
    powershell.exe -File .\PS_AD_Search.ps1
    Optional parameters:
        -UserName <username>  : Run a user group membership search
        -GroupName <groupname>: Run a group membership lookup
        -h or -help           : Display this help information

DESCRIPTION:
    Search Active Directory across multiple domains for:
        - The groups a user is a member of
        - The members of a group

REQUIREMENTS:
    - Requires the ImportExcel module
    - Appropriate permissions for AD access across domains

"@ -ForegroundColor Cyan
    exit
}

# Prompt for auto-export preference if no parameters are passed
if (-not $UserName -and -not $GroupName) {
    $exportChoice = Read-Host "Would you like to auto-export all results to Excel? (Y/N)"
    if ($exportChoice -match '^Y$') { $global:AutoExport = $true }
}

# Ensure ImportExcel module is installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Installing ImportExcel module..." -ForegroundColor Cyan
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}

# Define the domains to search in
$domains = @("SQSENIOR", "SQIS-CORP")

# Header printing utility function
function Write-ArtsyHeader {
    param ([string]$title ,[string]$ForegroundColor = "Cyan")
    Write-Host "===========================================" -ForegroundColor $ForegroundColor
    Write-Host "===       AD Search: $title       ===" -ForegroundColor $ForegroundColor
    Write-Host "===========================================" -ForegroundColor $ForegroundColor
}

# Function to search for all groups a user is a member of
function UserSearch {
    Write-ArtsyHeader -title "User Search" -foregroundcolor green

    if (-not $UserName) {
        $UserName = Read-Host "Enter the AD user name to search"
    }

    if ([string]::IsNullOrWhiteSpace($UserName)) {
        Write-Host "User name is required." -ForegroundColor Red
        return
    }

    $userGroups = @()
    foreach ($domain in $domains) {
        try {
            $user = Get-ADUser -Identity $UserName -Server $domain -Properties memberOf
            $userGroups = $user.memberOf
            if ($userGroups.Count -gt 0) {
                Write-Host "User '$UserName' found in domain '$domain' with $($userGroups.Count) group(s)." -ForegroundColor Green
                break
            } else {
                Write-Host "User '$UserName' exists in domain '$domain' but is not a member of any groups." -ForegroundColor Yellow
            }
        } catch {
            Write-Host "User '$UserName' not found in domain '$domain'." -ForegroundColor Yellow
        }
    }

    if ($userGroups.Count -eq 0) {
        Write-Host "User '$UserName' could not be found in any of the domains." -ForegroundColor Red
        return
    }

    $results = @()
    $counter = 0
    $total = $userGroups.Count
    foreach ($group in $userGroups) {
        $counter++
        Write-Progress -Activity "Resolving user groups in domain '$domain'" -Status "$counter of $total" -PercentComplete (($counter / $total) * 100)
        try {
            $resolved = Get-ADGroup -Identity $group -Server $domain -Properties Name
            $entry = [PSCustomObject]@{
                GroupName      = $resolved.Name
                GroupDomain    = $domain
            }
            $results += $entry
        } catch {
            $results += [PSCustomObject]@{
                GroupName      = $group
                GroupDomain    = $domain
            }
        }
    }

    Write-Progress -Activity "Complete" -Completed
    Write-Host "Search complete." -ForegroundColor Green

    if ($global:AutoExport -or (Read-Host "Would you like to export the results to Excel? (Y/N)" -eq 'Y')) {
        Write-Host "Exporting Results." -ForegroundColor Cyan
        $sheetName = "$domain-$UserName" -replace '[\\/]','-'
        $results |
            Select-Object GroupName, GroupDomain |
            Export-Excel -AutoSize -Show -WorksheetName $sheetName
    }
}

# Function to search for members of a group, with FSP and nested group resolution
function GroupSearch {
    Write-ArtsyHeader -title "Group Search" -foregroundcolor cyan
    if (-not $GroupName) {
        $GroupName = Read-Host "Enter the AD group name to search"
    }
    if ([string]::IsNullOrWhiteSpace($GroupName)) {
        Write-Host "Group name is required." -ForegroundColor Red
        return
    }

    $originalDomains = $domains.Clone()

    if ($GroupName -match '^(?<domain>[^\\/]+)[\\/](?<name>.+)$') {
        $defaultDomain = $matches.domain
        $GroupName = $matches.name
        $domains = @($defaultDomain) + ($domains | Where-Object { $_ -ne $defaultDomain })
    }

    $groupMembers = @()
    foreach ($domain in $domains) {
        try {
            $group = Get-ADGroup -Identity $GroupName -Server $domain -Properties member
            $groupMembers = $group.member
            if ($groupMembers.Count -gt 0) {
                Write-Host "Group '$GroupName' found in domain '$domain' with $($groupMembers.Count) member(s)." -ForegroundColor Green

                if ($groupMembers.Count -gt 50) {
                    $proceed = Read-Host "This group has more than 50 members. Loading them may take some time. Would you like to continue? (Y/N)"
                    if ($proceed -notmatch '^Y$') {
                        $domains = $originalDomains
                        return
                    }
                }
                $resolvedDomain=$domain
                break
            } else {
                Write-Host "Group '$GroupName' exists in domain '$domain' but has no members." -ForegroundColor Yellow
            }
        } catch {
            Write-Host "Group '$GroupName' not found in domain '$domain'." -ForegroundColor Yellow
        }
    }

    if ($groupMembers.Count -eq 0) {
        Write-Host "Group '$GroupName' could not be found in any of the domains." -ForegroundColor Red
        $domains = $originalDomains
        return
    }

    $results = @()
    $counter = 0
    $total = $groupMembers.Count
    foreach ($member in $groupMembers) {
        $counter++
        Write-Progress -Activity "Resolving group members in domain '$resolvedDomain'" -Status "$counter of $total" -PercentComplete (($counter / $total) * 100)
        try {
            $resolved = Get-ADObject -Identity $member -Server $resolvedDomain -Properties SamAccountName, Name, ObjectClass, ObjectSID
            $entry = [PSCustomObject]@{
                SID            = $member
                Name           = $resolved.Name
                SamAccountName = $resolved.SamAccountName
                ObjectClass    = $resolved.ObjectClass
                Domain         = $resolvedDomain
            }
            # Check if this is a nested group
            if ($resolved.ObjectClass -eq 'group') {
                $entry.Name += " (Nested Group)"
            }
            $results += $entry
        } catch {
            $results += [PSCustomObject]@{
                SID            = $member
                Name           = $member
                SamAccountName = ''
                ObjectClass    = 'foreignSecurityPrincipal'
                Domain         = $resolvedDomain
            }
        }
    }

    $counter = 0
    $FSPs = $results | Where-Object { $_.ObjectClass -eq 'foreignSecurityPrincipal' -or $_.Name -match '^S-1-5-21-.*' }
    [int]$totalFSP = ($FSPs | Measure-Object).Count
    Write-Host "Number of FSP's to resolve:" $totalFSP -ForegroundColor Cyan
    if ($totalFSP -gt 0) {
        foreach ($entry in $results) {
            if ($entry.ObjectClass -eq 'foreignSecurityPrincipal' -or $entry.Name -match '^S-1-5-21-.*') {
                $counter++
                Write-Progress -Activity "Resolving FSPs across domains" -Status "$counter of $totalFSP" -PercentComplete (($counter / $totalFSP) * 100)
                $rawSid = ($entry.Name -split ',')[0]
                $resolved = $null
                foreach ($domain in ($domains | Where-Object { $_ -ne $resolvedDomain })) {
                    try {
                        $resolved = Get-ADObject -Filter "ObjectSID -eq '$($rawSid)'" -Server $domain -Properties SamAccountName, Name, ObjectClass
                        if ($resolved) {
                            $entry.Name           = $resolved.Name
                            if ($resolved.ObjectClass -eq 'group') { $entry.Name += " (Nested Group)" }
                            $entry.SamAccountName = $resolved.SamAccountName
                            $entry.ObjectClass    = $resolved.ObjectClass
                            $entry.Domain         = $domain
                            break
                        }
                    } catch {}
                }
            }
        }
    } else {
        Write-Host "No FSPs to resolve." -ForegroundColor Green
    }

    Write-Progress -Activity "Complete" -Completed
    Write-Host "Search complete." -ForegroundColor Green
    if ($global:AutoExport -or (Read-Host "Would you like to export the results to Excel? (Y/N)" -eq 'Y')) {
        Write-Host "Exporting Results." -ForegroundColor Cyan
        $sheetName = "$resolvedDomain-$GroupName" -replace '[\\/]','-'
        $results |
            Select-Object Name, SamAccountName, ObjectClass, Domain |
            Export-Excel -AutoSize -Show -WorksheetName $sheetName
    }
    $domains = $originalDomains
}

# Handle execution via CLI parameter or prompt interactively
if ($UserName) {
    UserSearch
    return
} elseif ($GroupName) {
    GroupSearch
    return
}

# Interactive main loop for user choice
while ($true) {
    Write-ArtsyHeader -title "Welcome     " -foregroundcolor Green
    Write-Host "Choose Search Type" -ForegroundColor Cyan
    Write-Host "[1] User Search"
    Write-Host "[2] Group Search"
    $choice = Read-Host "Enter choice (1 or 2)"
    switch ($choice) {
        '1' { UserSearch }
        '2' { GroupSearch }
        default { Write-Host "Invalid choice. Please enter 1 or 2." -ForegroundColor Red; continue }
    }
    $again = Read-Host "Would you like to run another search? (Y to continue, Enter to exit)"
    if ($again -notmatch '^Y$') { break }
}
