# Ensure ImportExcel module is installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Installing ImportExcel module..." -ForegroundColor Cyan
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}

# Define domain array
$domains = @("SQSENIOR", "SQIS-CORP")

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
                    Get-ADGroup -Identity $_ -Server $domain | Select-Object Name, SamAccountName, GroupScope, DistinguishedName
                } catch {
                    # If the group cannot be resolved in this domain
                    [PSCustomObject]@{
                        Name = $_
                        SamAccountName = "N/A"
                        GroupScope = "Unknown"
                        DistinguishedName = $_
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
$sheetName = "$userName-Groups" -replace '[\/]', '-'
$allGroups |
    Select-Object Name, SamAccountName, GroupScope, Domain, DistinguishedName |
    Export-Excel -AutoSize -Show -WorksheetName $sheetName