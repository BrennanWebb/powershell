# Ensure ImportExcel module is installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Installing ImportExcel module..." -ForegroundColor Cyan
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}

Add-Type -AssemblyName System.Windows.Forms

$form = New-Object System.Windows.Forms.Form
$form.Text = "AD Group Search"
$form.Size = New-Object System.Drawing.Size(400, 250)
$form.StartPosition = "CenterScreen"

$label = New-Object System.Windows.Forms.Label
$label.Text = "Enter the AD group name:"
$label.AutoSize = $true
$label.Location = New-Object System.Drawing.Point(10, 20)
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Size = New-Object System.Drawing.Size(360, 20)
$textBox.Location = New-Object System.Drawing.Point(10, 45)
$form.Controls.Add($textBox)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = ""
$statusLabel.AutoSize = $true
$statusLabel.Location = New-Object System.Drawing.Point(10, 75)
$form.Controls.Add($statusLabel)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 100)
$progressBar.Size = New-Object System.Drawing.Size(360, 25)
$progressBar.Style = 'Continuous'
$form.Controls.Add($progressBar)

$button = New-Object System.Windows.Forms.Button
$button.Text = "Search"
$button.Location = New-Object System.Drawing.Point(10, 140)

$button.Add_Click({
    $button.Enabled = $false
    $textBox.Enabled = $false
    $statusLabel.Text = "Please wait..."
    $form.Refresh()

    $groupName = $textBox.Text.Trim()
    if (-not $groupName) {
        [System.Windows.Forms.MessageBox]::Show("No group name provided. Exiting.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $form.Close()
        return
    }

    $domains = @("SQSENIOR","SQIS-CORP")
    if ($groupName -match '^(?<domain>[^\\/]+)[\\/](?<name>.+)$') {
        $defaultDomain = $matches.domain
        $groupName = $matches.name
        $domains = @($defaultDomain) + ($domains | Where-Object { $_ -ne $defaultDomain })
    }

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
        $form.Close()
        return
    }

    $results = @()
    $counter = 0
    $total = $groupMembers.Count
    $progressBar.Maximum = $total
    $progressBar.Value = 0
    $statusLabel.Text = "Resolving group members..."

    foreach ($member in $groupMembers) {
        $counter++
        $statusLabel.Text = "Resolving group members ($counter of $total)..."
        $progressBar.Value = [Math]::Min($counter, $progressBar.Maximum)
        $form.Refresh()

        try {
            $resolved = Get-ADObject -Identity $member -Server $resolvedDomain -Properties SamAccountName, Name, ObjectClass, ObjectSID
            $results += [PSCustomObject]@{
                SID            = $member
                Name           = $resolved.Name
                SamAccountName = $resolved.SamAccountName
                ObjectClass    = $resolved.ObjectClass
                Domain         = $resolvedDomain
            }
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

    $FSPs = $results | Where-Object { $_.ObjectClass -eq 'foreignSecurityPrincipal' -or $_.Name -match '^S-1-5-21-.*' }
    [int]$totalFSP = ($FSPs | Measure-Object).Count
    Write-Host "Number of FSP's to resolve:" $totalFSP -ForegroundColor Cyan

    if ($totalFSP -gt 0) {
        $counter = 0
        $progressBar.Maximum = $totalFSP
        $progressBar.Value = 0
        $statusLabel.Text = "Resolving FSPs..."

        foreach ($entry in $results) {
            if ($entry.ObjectClass -eq 'foreignSecurityPrincipal' -or $entry.Name -match '^S-1-5-21-.*') {
                $counter++
                $statusLabel.Text = "Resolving FSPs ($counter of $totalFSP)..."
                $progressBar.Value = [Math]::Min($counter, $progressBar.Maximum)
                $form.Refresh()

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
                    } catch {}
                }
            }
        }
    } else {
        Write-Host "No FSPs to resolve." -ForegroundColor Green
    }

    $sheetName = "$resolvedDomain-$groupName" -replace '[\\\/]', '-'

    $statusLabel.Text = "Exporting to Excel..."
    $form.Refresh()

    $results |
        Select-Object Name, SamAccountName, ObjectClass, Domain |
        Export-Excel -AutoSize -Show -WorksheetName $sheetName

    $statusLabel.Text = "Complete. Excel exported!"
    Start-Sleep -Seconds 2
    $form.Close()
})

$form.Controls.Add($button)
$form.AcceptButton = $button
$form.Topmost = $true
$form.ShowDialog()