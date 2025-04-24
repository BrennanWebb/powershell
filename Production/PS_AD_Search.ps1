# PS_AD_Search - Combined AD User and Group Search Tool
# Version: 1.0.1
# Description: Interactive script to search Active Directory for user group memberships or group members across multiple domains.

# Ensure ImportExcel module is installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Installing ImportExcel module..." -ForegroundColor Cyan
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}

# Define domain array
$domains = @("SQSENIOR", "SQIS-CORP")

function UserSearch {
    # Prompt user for username
    $userName = Read-Host "Enter the username to search (e.g. jdoe)"

    # Initialize list for all group memberships
    $allGroups = @()

    foreach ($domain in $domains) {
        Write-Host "Searching in domain: $domain" -ForegroundColor Cyan
        try {
            $user = Get-ADUser -Identity $userName -Server $domain -Properties MemberOf
            if ($user) {
                $groups = $user.MemberOf | ForEach-Object {
                    try {
                        $group = Get-ADGroup -Identity $_ -Server $domain -Properties Members
                        $memberCount = ($group.Members).Count
                        [PSCustomObject]@{
                            Name = $group.Name
                            SamAccountName = $group.SamAccountName
                            GroupScope = $group.GroupScope
                            DistinguishedName = $group.DistinguishedName
                            MemberCount = $memberCount
                        }
                    } catch {
                        # If the group cannot be resolved in this domain
                        [PSCustomObject]@{
                            Name = $_
                            SamAccountName = "N/A"
                            GroupScope = "Unknown"
                            DistinguishedName = $_
                            MemberCount = "N/A"
                        }
                    }
                }

                $allGroups += $groups | ForEach-Object {
                    $_ | Add-Member -NotePropertyName Domain -NotePropertyValue $domain -PassThru
                }
            }
        } catch {
            Write-Host "User '$userName' not found in domain '$domain'." -ForegroundColor Yellow
        }
    }

    if ($allGroups.Count -eq 0) {
        Write-Host "No groups found for user '$userName' in any of the specified domains." -ForegroundColor Red
        return
    }

    # Export to Excel
    $sheetName = "$userName-Groups" -replace '[\/]','-'
    $allGroups |
        Select-Object Name, SamAccountName, GroupScope, Domain, DistinguishedName, MemberCount |
        Export-Excel -AutoSize -Show -WorksheetName $sheetName
}

function GroupSearch {
    # Prompt user for group name
    $groupName = Read-Host "Enter the AD group name to search"

    # If groupName includes a domain prefix, prioritize that domain
    if ($groupName -match '^(?<domain>[^\\/]+)[\\/](?<name>.+)$') {
        $defaultDomain = $matches.domain
        $groupName = $matches.name
        $domains = @($defaultDomain) + ($domains | Where-Object { $_ -ne $defaultDomain })
    }

    # Initialize variable to hold the group members
    $groupMembers = @()

    foreach ($domain in $domains) {
        try {
            $groupMembers = Get-ADGroup -Identity $groupName -Server $domain -Properties member |
                Select-Object -ExpandProperty member
            if ($groupMembers.Count -gt 0) {
                Write-Host "Group '$groupName' found in domain '$domain' with $($groupMembers.Count) member(s)." -ForegroundColor Green
                $resolvedDomain=$domain
                break
            } else {
                Write-Host "Group '$groupName' exists in domain '$domain' but has no members." -ForegroundColor Yellow
            }
        } catch {
            Write-Host "Group '$groupName' not found in domain '$domain'." -ForegroundColor Yellow
        }
    }

    if ($groupMembers.Count -eq 0) {
        Write-Host "Group '$groupName' could not be found in any of the domains." -ForegroundColor Red
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
            $results += [PSCustomObject]@{
                SID            = $member
                Name           = $resolved.Name
                SamAccountName = $resolved.SamAccountName
                ObjectClass    = $resolved.ObjectClass
                Domain         = $resolvedDomain
            }
        }
        catch {
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
                            $entry.SamAccountName = $resolved.SamAccountName
                            $entry.ObjectClass    = $resolved.ObjectClass
                            $entry.Domain         = $domain
                            break
                        }
                    } catch {
                    }
                }
            }
        }
    } else {
        Write-Host "No FSPs to resolve." -ForegroundColor Green
    }

    Write-Progress -Activity "Complete" -Completed

    $sheetName = "$resolvedDomain-$groupName" -replace '[\\/]','-'

    $results |
        Select-Object Name, SamAccountName, ObjectClass, Domain |
        Export-Excel -AutoSize -Show -WorksheetName $sheetName
}

# Main prompt loop
while ($true) {
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
