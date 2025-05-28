# Ensure ImportExcel module is installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Installing ImportExcel module..." -ForegroundColor Cyan
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}

# Define domain array in priority order
$domains = @("SQSENIOR","SQIS-CORP")

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

# Loop through each domain to find the group and set group members
foreach ($domain in $domains) {
    try {
        # Attempt to get the group from the current domain
        $groupMembers = Get-ADGroup -Identity $groupName -Server $domain -Properties member |
            Select-Object -ExpandProperty member
        if ($groupMembers.Count -gt 0) {
            Write-Host "Group '$groupName' found in domain '$domain' with $($groupMembers.Count) member(s)." -ForegroundColor Green
            $resolvedDomain=$domain
            break # Exit loop once group is found
        } else {
            Write-Host "Group '$groupName' exists in domain '$domain' but has no members." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Group '$groupName' not found in domain '$domain'." -ForegroundColor Yellow
    }
}

# If group members are still empty, exit the script
if ($groupMembers.Count -eq 0) {
    Write-Host "Group '$groupName' could not be found in any of the domains." -ForegroundColor Red
    return
}

# First pass: Try to resolve in primary domain (the first domain in $domains)
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
        # Assume FSP or unresolved
        $results += [PSCustomObject]@{
            SID            = $member
            Name           = $member
            SamAccountName = ''
            ObjectClass    = 'foreignSecurityPrincipal'
            Domain         = $resolvedDomain
        }
    }
}

# Second pass: Try to resolve FSPs using remaining domains
$counter = 0
$FSPs = $results | Where-Object { $_.ObjectClass -eq 'foreignSecurityPrincipal' -or $_.Name -match '^S-1-5-21-.*' }
[int]$totalFSP = ($FSPs | Measure-Object).Count
Write-Host "Number of FSP's to resolve:" $totalFSP -ForegroundColor Cyan

if ($totalFSP -gt 0) {
    foreach ($entry in $results) {
        if ($entry.ObjectClass -eq 'foreignSecurityPrincipal'-or $entry.Name -match '^S-1-5-21-.*') {
            $counter++
            Write-Progress -Activity "Resolving FSPs across domains" -Status "$counter of $totalFSP" -PercentComplete (($counter / $totalFSP) * 100)
            
            $rawSid = ($entry.Name -split ',')[0]  # Extract SID
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
                }
                catch {
                    # Try next domain
                }
            }
        }
    }
} else {
    Write-Host "No FSPs to resolve." -ForegroundColor Green
}  

Write-Progress -Activity "Complete" -Completed

# Clean sheet name (Excel doesn't allow / or \ in sheet names)
$sheetName = "$resolvedDomain-$groupName" -replace '[\\\/]', '-'

# Export to Excel with custom worksheet name
$results |
    Select-Object Name, SamAccountName, ObjectClass, Domain |
    Export-Excel -AutoSize -Show -WorksheetName $sheetName
