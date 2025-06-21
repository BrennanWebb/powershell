<#
.SYNOPSIS
    A multi-tool for searching Active Directory users and groups across all discoverable domains, with results exported to Excel.

.DESCRIPTION
    This script automatically discovers all discoverable domains and provides two primary search functions:
    1. User Search: Finds users by AD Login, Name, or SID (wildcards supported for names). If multiple users are found, it prompts for selection.
       It then lists the chosen user's group memberships, asking once to perform a "deep scan" on all non-home domains.
    2. Group Search: Finds groups by name (wildcards supported). If multiple groups are found, it prompts for selection before analyzing the group's members.

    All results are exported directly to an Excel file that opens automatically. The script will loop, allowing for multiple searches.

.NOTES
    Author: Brennan Webb via Gemini AI
    Version: 9.3
    Dependencies: This script requires the 'ImportExcel' module. It will attempt to install it if it is not found.
#>

[CmdletBinding()]
param ()

#region Prerequisite: Ensure ImportExcel Module is installed
try {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "The 'ImportExcel' module is required. Attempting to install..." -ForegroundColor Cyan
        Install-Module -Name ImportExcel -Force -Scope CurrentUser -ErrorAction Stop
        Write-Host "ImportExcel module installed successfully." -ForegroundColor Green
    }
    Import-Module ImportExcel
}
catch {
    Write-Error "Failed to install the ImportExcel module. Error: $($_.Exception.Message)"
    return
}
#endregion

#region Domain Discovery Function
function Get-DiscoveredDomains {
    [CmdletBinding()]
    param()
    $localForest = Get-ADForest
    $localDomains = $localForest.Domains
    $trusts = $null
    if (Get-Command Get-ADForestTrust -ErrorAction SilentlyContinue) {
        $trusts = Get-ADForestTrust -Identity $localForest.Name -ErrorAction SilentlyContinue
    }
    if (-not $trusts) {
        $trusts = Get-ADTrust -Filter * -ErrorAction SilentlyContinue
    }
    $trustedTargets = $trusts | Select-Object -ExpandProperty Target -Unique
    $allDiscoveredDomainNames = $localDomains + $trustedTargets | Sort-Object -Unique
    return $allDiscoveredDomainNames
}
#endregion

#region Function: User Membership Search
function Start-UserSearch {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$SearchInput,
        [Parameter(Mandatory)]
        [string[]]$SearchDomains
    )
    
    $primaryUserObject = $null
    $homeDomain = $null

    # --- Path A: User provides a SID ---
    if ($SearchInput -match '^S-1-[\d-]+$') {
        Write-Host "SID format detected. Searching for user by SID..." -ForegroundColor Cyan
        try {
            # -Identity on Get-ADUser will search the global catalog for a SID, no need to loop domains.
            $primaryUserObject = Get-ADUser -Identity $SearchInput -Properties MemberOf, SID -ErrorAction Stop
            # Determine the home domain from the found object's distinguished name
            $homeDomain = ($primaryUserObject.DistinguishedName -split '(?<!\\),DC=')[1..99] -join '.'
        }
        catch {
            Write-Error "Could not find a user with the specified SID: $SearchInput"
            return
        }
    }
    # --- Path B: User provides a Name or AD Login ---
    else {
        Write-Host "Searching for users matching '$SearchInput' across all domains..." -ForegroundColor Cyan
        $foundUsers = [System.Collections.ArrayList]@()
        foreach ($Domain in $SearchDomains) {
            try {
                $filter = "(SamAccountName -like ""$SearchInput"") -or (Name -like ""$SearchInput"") -or (GivenName -like ""$SearchInput"") -or (Surname -like ""$SearchInput"")"
                $usersInDomain = Get-ADUser -Filter $filter -Server $Domain -Properties Name, SamAccountName, GivenName, Surname -ErrorAction Stop
                if ($usersInDomain) {
                    foreach ($user in $usersInDomain) {
                        [void]$foundUsers.Add([PSCustomObject]@{ UserObject = $user; Domain = $Domain })
                    }
                }
            }
            catch {}
        }

        $selectedUserContainer = $null
        switch ($foundUsers.Count) {
            0 { Write-Warning "No users found matching '$SearchInput' in any of the specified domains."; return }
            1 {
                Write-Host "One user found. Auto-selecting." -ForegroundColor Green
                $selectedUserContainer = $foundUsers[0]
            }
            default {
                Write-Host "--------------------------------------------------" -ForegroundColor Cyan
                Write-Host "Multiple users found. Please select one to continue:" -ForegroundColor Yellow
                for ($i = 0; $i -lt $foundUsers.Count; $i++) {
                    $userEntry = $foundUsers[$i]
                    $number = $i + 1
                    Write-Host ('{0,3}. {1,-30} ({2,-20}) - {3}' -f $number, $userEntry.UserObject.Name, $userEntry.UserObject.SamAccountName, $userEntry.Domain) -ForegroundColor White
                }
                Write-Host "--------------------------------------------------" -ForegroundColor Cyan
                while (-not $selectedUserContainer) {
                    $choice = Read-Host "Please enter the number of the user to analyze (or press Enter to cancel)"
                    if ([string]::IsNullOrWhiteSpace($choice)) { return }
                    if (($choice -match '^\d+$') -and ([int]$choice -ge 1) -and ([int]$choice -le $foundUsers.Count)) {
                        $selectedUserContainer = $foundUsers[[int]$choice - 1]
                    } else { Write-Warning "Invalid selection. Please enter a number between 1 and $($foundUsers.Count)." }
                }
            }
        }
        
        try {
            # Get the full user object to ensure we have all properties for the analysis phase
            $primaryUserObject = Get-ADUser -Identity $selectedUserContainer.UserObject -Server $selectedUserContainer.Domain -Properties MemberOf, SID -ErrorAction Stop
            $homeDomain = $selectedUserContainer.Domain
        } catch { Write-Error "Could not retrieve full details for the selected user. Error: $($_.Exception.Message)"; return }
    }


    # --- Analysis Phase (This part is now common to both search paths) ---
    Write-Host "SUCCESS: Analyzing user '$($primaryUserObject.Name)' from domain '$homeDomain'." -ForegroundColor Green
    
    $performDeepScan = $false
    $otherDomains = $SearchDomains | Where-Object { $_ -ne $homeDomain }
    if ($otherDomains) {
        Write-Host
        $response = Read-Host "Scan all other domains for memberships? (Press ENTER for Yes / N for No)"
        if ($response.ToUpper() -ne 'N') { $performDeepScan = $true }
    }

    $allFoundGroups = [System.Collections.ArrayList]@()
    $domainsToProcess = if ($performDeepScan) { $SearchDomains } else { @($homeDomain) }
    Write-Host "--------------------------------------------------" -ForegroundColor Cyan
    foreach ($Domain in $domainsToProcess) {
        Write-Host "Processing domain '$Domain'..." -ForegroundColor Cyan
        if ($Domain -eq $homeDomain) {
            Write-Host "Retrieving direct group memberships from home domain..." -ForegroundColor Cyan
            foreach ($groupDN in $primaryUserObject.MemberOf) {
                try {
                    $group = Get-ADGroup -Identity $groupDN -Server $Domain -Properties Members, GroupScope, objectSid -ErrorAction Stop
                    [void]$allFoundGroups.Add([PSCustomObject]@{ Name = $group.Name; GroupScope = $group.GroupScope; Domain = $Domain; 'Member Count' = $group.Members.Count; SID = $group.objectSid.Value })
                } catch { Write-Warning "Could not retrieve details for group '$groupDN' in '$Domain'." }
            }
        } else {
            Write-Host "Performing deep scan on '$Domain' based on user request..." -ForegroundColor Cyan
            try {
                $domainInfo = Get-ADDomain -Server $Domain -ErrorAction Stop
                $fspDN = "CN=$($primaryUserObject.SID.Value),CN=ForeignSecurityPrincipals,$($domainInfo.DistinguishedName)"
                $fspGroups = Get-ADGroup -LDAPFilter "(member=$fspDN)" -Server $Domain -Properties Members, GroupScope, objectSid -ErrorAction Stop
                if($fspGroups) {
                     Write-Host "Found $($fspGroups.Count) group(s) with FSP membership in '$Domain'." -ForegroundColor Green
                     foreach ($group in $fspGroups) { [void]$allFoundGroups.Add([PSCustomObject]@{ Name = $group.Name; GroupScope = $group.GroupScope; Domain = $Domain; 'Member Count' = $group.Members.Count; SID = $group.objectSid.Value }) }
                } else { Write-Warning "No FSP memberships found for user in '$Domain'." }
            } catch { Write-Error "An error occurred during the deep scan of '$Domain'. Error: $($_.Exception.Message)" }
        }
    }

    if ($allFoundGroups.Count -gt 0) {
        Write-Host "--------------------------------------------------" -ForegroundColor Cyan
        Write-Host "Exporting consolidated group memberships for '$($primaryUserObject.Name)' to Excel..." -ForegroundColor Green
        $sheetName = $primaryUserObject.SamAccountName -replace '[\\/*?:\[\]]'
        $sheetName = $sheetName.Substring(0, [System.Math]::Min(31, $sheetName.Length))
        $allFoundGroups | Sort-Object Domain, Name | Export-Excel -AutoSize -Show -WorksheetName $sheetName
    } else { Write-Warning "No group memberships were found for '$($primaryUserObject.Name)' in the processed domains." }
}
#endregion

#region Function: Group Search
function Start-GroupSearch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$GroupName,
        [Parameter(Mandatory)]
        [string[]]$SearchDomains
    )

    Write-Host "Searching for groups matching '$GroupName' across all domains..." -ForegroundColor Cyan
    $foundGroups = [System.Collections.ArrayList]@()
    foreach($Domain in $SearchDomains) {
        try {
            $groupsInDomain = Get-ADGroup -Filter "Name -like '$GroupName'" -Server $Domain -Properties SamAccountName -ErrorAction Stop
            if ($groupsInDomain) {
                foreach($group in $groupsInDomain) {
                    [void]$foundGroups.Add([PSCustomObject]@{ GroupObject = $group; Domain = $Domain })
                }
            }
        } catch {}
    }
    
    $selectedGroupContainer = $null
    switch ($foundGroups.Count) {
        0 { Write-Warning "No groups found matching '$GroupName' in any of the specified domains."; return }
        1 {
            Write-Host "One group found. Auto-selecting." -ForegroundColor Green
            $selectedGroupContainer = $foundGroups[0]
        }
        default {
            Write-Host "--------------------------------------------------" -ForegroundColor Cyan
            Write-Host "Multiple groups found. Please select one to continue:" -ForegroundColor Yellow
            for ($i = 0; $i -lt $foundGroups.Count; $i++) {
                $groupEntry = $foundGroups[$i]
                $number = $i + 1
                Write-Host ('{0,3}. {1,-40} - {2}' -f $number, $groupEntry.GroupObject.Name, $groupEntry.Domain) -ForegroundColor White
            }
            Write-Host "--------------------------------------------------" -ForegroundColor Cyan
            while (-not $selectedGroupContainer) {
                $choice = Read-Host "Please enter the number of the group to analyze (or press Enter to cancel)"
                if ([string]::IsNullOrWhiteSpace($choice)) { return }
                if (($choice -match '^\d+$') -and ([int]$choice -ge 1) -and ([int]$choice -le $foundGroups.Count)) {
                    $selectedGroupContainer = $foundGroups[[int]$choice - 1]
                } else { Write-Warning "Invalid selection. Please enter a number between 1 and $($foundGroups.Count)." }
            }
        }
    }

    $selectedGroup = $selectedGroupContainer.GroupObject
    $selectedGroupDomain = $selectedGroupContainer.Domain
    Write-Host "SUCCESS: Analyzing group '$($selectedGroup.Name)' from domain '$selectedGroupDomain'." -ForegroundColor Green
    
    $results = [System.Collections.ArrayList]@()
    function Get-AllMemberDNs {
        param([string]$GroupDN, [string]$Server, [System.Collections.ArrayList]$ProcessedGroups)
        if ($ProcessedGroups -contains $GroupDN) { return @() }
        [void]$ProcessedGroups.Add($GroupDN)
        $dns = New-Object System.Collections.ArrayList
        try {
            $group = Get-ADGroup -Identity $GroupDN -Server $Server -Properties Members -ErrorAction Stop
            foreach ($memberDN in $group.Members) {
                [void]$dns.Add($memberDN)
                $memberObject = Get-ADObject -Identity $memberDN -Server $Server -Properties ObjectClass -ErrorAction SilentlyContinue
                if ($memberObject.ObjectClass -eq 'group') {
                    $nestedDns = Get-AllMemberDNs -GroupDN $memberDN -Server $Server -ProcessedGroups $ProcessedGroups
                    $dns.AddRange($nestedDns)
                }
            }
        } catch { Write-Warning "Could not process group '$GroupDN' to get member DNs. Error: $($_.Exception.Message)" }
        return $dns
    }

    $allMemberDns = Get-AllMemberDNs -GroupDN $selectedGroup.DistinguishedName -Server $selectedGroupDomain -ProcessedGroups ([System.Collections.ArrayList]@())
    
    Write-Host "Processing unique members..." -ForegroundColor Cyan
    $uniqueMemberDns = $allMemberDns | Sort-Object -Unique
    foreach ($dn in $uniqueMemberDns) {
        $isFsp = $dn -like '*CN=S-1-5-*,CN=ForeignSecurityPrincipals,*'
        [void]$results.Add([PSCustomObject]@{ DistinguishedName = $dn; Name = if ($isFsp) { ($dn -split ',')[0] } else { '' }; SamAccountName = ''; ObjectClass = if ($isFsp) { 'foreignSecurityPrincipal' } else { '' }; Domain = ($dn -split '(?<!\\),DC=')[1..99] -join '.'; EmailAddress = ''; SID = ''; IsResolved = $false })
    }

    $unresolved = $results | Where-Object { -not $_.IsResolved }
    $totalToResolve = $unresolved.Count
    if ($totalToResolve -gt 0) {
        Write-Host "Found $totalToResolve unique FSPs and cross-domain members to resolve. This may take some time..." -ForegroundColor Cyan
        $counter = 0
        foreach ($entry in $unresolved) {
            $counter++
            $status = 'Resolving object {0} of {1}: {2}' -f $counter, $totalToResolve, $entry.Name
            Write-Progress -Activity "Resolving Foreign & Cross-Domain Principals" -Status $status -PercentComplete (($counter / $totalToResolve) * 100)
            foreach($searchDomain in $SearchDomains) {
                try {
                    $resolvedObject = $null
                    $properties = @('SamAccountName', 'mail', 'objectSid')
                    if ($entry.ObjectClass -eq 'foreignSecurityPrincipal') {
                        $sid = ($entry.Name -split '=')[1]
                        $resolvedObject = Get-ADObject -Filter "ObjectSID -eq '$sid'" -Server $searchDomain -Properties $properties -ErrorAction Stop
                    } else {
                        $resolvedObject = Get-ADObject -Identity $entry.DistinguishedName -Server $searchDomain -Properties $properties -ErrorAction Stop
                    }
                    if ($resolvedObject) {
                        $entry.Name = $resolvedObject.Name; $entry.SamAccountName = $resolvedObject.SamAccountName; $entry.ObjectClass = $resolvedObject.ObjectClass
                        $entry.Domain = ($resolvedObject.DistinguishedName -split '(?<!\\),DC=')[1..99] -join '.'; $entry.EmailAddress = $resolvedObject.mail; $entry.SID = $resolvedObject.objectSid.Value; $entry.IsResolved = $true
                        break 
                    }
                } catch {}
            }
        }
        Write-Progress -Activity "Resolving Foreign & Cross-Domain Principals" -Completed
    }

    $finalOutput = foreach ($entry in $results) { [PSCustomObject]@{ Name = $entry.Name; SamAccountName = $entry.SamAccountName; ObjectClass = $entry.ObjectClass; Domain = $entry.Domain; EmailAddress = $entry.EmailAddress; SID = $entry.SID } }
    
    if ($finalOutput.Count -gt 0) {
        Write-Host "Exporting all unique group members to Excel..." -ForegroundColor Green
        $sheetName = $selectedGroup.SamAccountName -replace '[\\/*?:\[\]]'
        $sheetName = $sheetName.Substring(0, [System.Math]::Min(31, $sheetName.Length))
        $finalOutput | Sort-Object Domain, Name | Export-Excel -AutoSize -Show -WorksheetName $sheetName
    } else { Write-Warning "Search complete, but no members were found or could be resolved in the group(s)." }
}
#endregion

#region Main Script Body - Interactive Menu
try {
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        throw "The Active Directory module is not installed. Please install the RSAT tools for ActiveDirectory."
    }
    Write-Host "Discovering trusted domains..." -ForegroundColor Cyan
    $Domains = Get-DiscoveredDomains
    if (-not $Domains) {
        throw "No domains were discovered. Cannot proceed. Please ensure this script is run on a domain-joined computer."
    }
    while ($true) {
        Clear-Host
        Write-Host "--------------------------------------------------" -ForegroundColor Cyan
        Write-Host "Active Directory Multi-Domain Search Tool" -ForegroundColor Cyan
        Write-Host "Domains being searched:" -ForegroundColor Yellow
        for ($i = 0; $i -lt $Domains.Count; $i++) {
            $domain = $Domains[$i]
            $number = $i + 1
            Write-Host "    $number. $domain" -ForegroundColor Yellow
        }
        Write-Host
        Write-Host "1: Search for a User (and list their group memberships)" -ForegroundColor White
        Write-Host "2: Search for a Group (and list its members)" -ForegroundColor White
        $choice = Read-Host "Please select an option (1 or 2)"
        switch($choice) {
            '1' {
                $userInput = Read-Host "Enter the User's AD Login, Name, or SID to search for (wildcards '*' supported for names)"
                if(-not [string]::IsNullOrWhiteSpace($userInput)) { Start-UserSearch -SearchInput $userInput -SearchDomains $Domains } 
                else { Write-Warning "Input cannot be empty." }
            }
            '2' {
                $userInput = Read-Host "Enter the Group Name to search for (wildcards '*' are supported)"
                if(-not [string]::IsNullOrWhiteSpace($userInput)) { Start-GroupSearch -GroupName $userInput -SearchDomains $Domains }
                else { Write-Warning "Input cannot be empty." }
            }
            default {
                Write-Error "Invalid option."
            }
        }
        Write-Host
        $again = Read-Host "Would you like to run another search? (Y to continue, Enter to exit)"
        if ($again -notmatch '^Y$') {
            break 
        }
    }
}
catch {
    Write-Error "A critical error occurred: $($_.Exception.Message)"
}
#endregion