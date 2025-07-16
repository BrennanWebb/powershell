<#
.SYNOPSIS
    A graphical user interface for Optimus, the T-SQL tuning advisor.

.DESCRIPTION
    This script provides a Windows Form interface to run Optimus analysis. It gathers all necessary
    parameters and launches the Optimus.ps1 engine in a new console window.

.NOTES
    Designer: Brennan Webb
    GUI Engine: PowerShell + Windows Forms
    Version: 3.0
    Created: 2025-07-03
    Modified: 2025-07-07
    Change Log:
    - v3.0: Reorganized the 'Options' menu based on user feedback. Moved 'Set AI Model' to the top level and 'Set/Reset API Key' under the 'Reset Configuration' sub-menu.
    - v2.9: Reorganized the menus. Moved the "Prerequisites" options into a sub-menu under "Options" for a cleaner layout.
    - v2.8: Added a 'Prerequisites' menu to handle setting the API Key and AI Model directly within the GUI.
    - v2.7: Implemented a GUI prompt for the API key if it's not found, removing the need to run the console script for initial setup.
    Powershell Version: 5.1+
#>

# --- Minimize Host Console Window ---
try {
    # Check if we are in a host with a visible window before proceeding.
    $process = Get-Process -Id $PID
    if ($process.MainWindowHandle -ne [IntPtr]::Zero) {
        # Define the Win32 API method only if it hasn't been defined in this session.
        if (-not ([System.Management.Automation.PSTypeName]'Win32.Win32').Type) {
            $cSharpCode = @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@
            Add-Type -TypeDefinition $cSharpCode -Namespace 'Win32'
        }
        # 2 = SW_SHOWMINIMIZED
        [Win32.Win32]::ShowWindow($process.MainWindowHandle, 2) | Out-Null
    }
}
catch {
    # This is a non-critical feature. If it fails for any reason, we will write a warning.
    Write-Warning "Could not minimize the PowerShell console window."
}


# --- Load Windows Forms Assembly ---
Add-Type -AssemblyName System.Windows.Forms

#region Reused Core Functions

# Centralized Logging - MODIFIED FOR GUI
function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [ValidateSet('DEBUG', 'INFO', 'SUCCESS', 'WARN', 'ERROR', 'PROMPT', 'RESULT')]
        [string]$Level = 'INFO'
    )
    if ($null -ne $script:LogTextBox) {
        # Define colors for the RichTextBox
        $colorMap = @{
            'INFO'    = 'Cyan'
            'SUCCESS' = 'Green'
            'WARN'    = 'Yellow'
            'ERROR'   = 'Red'
            'PROMPT'  = 'White'
            'DEBUG'   = 'Gray'
            'RESULT'  = 'White'
        }
        $logColor = if ($colorMap.ContainsKey($Level)) { $colorMap[$Level] } else { 'White' }

        $timestamp = Get-Date -Format 'HH:mm:ss'
        $logMessage = "$timestamp [$Level] - $Message" + [Environment]::NewLine

        # Append text with color
        $script:LogTextBox.SelectionStart = $script:LogTextBox.TextLength
        $script:LogTextBox.SelectionLength = 0
        $script:LogTextBox.SelectionColor = $logColor
        $script:LogTextBox.AppendText($logMessage)
        $script:LogTextBox.SelectionColor = $script:LogTextBox.ForeColor # Reset color
        $script:LogTextBox.ScrollToCaret()

    } else {
        # Fallback to console if GUI isn't ready
        Write-Host "$timestamp [$Level] - $Message"
    }
}


# Configuration Management
function Initialize-Configuration {
    Write-Log -Message "Initializing Optimus configuration..." -Level 'DEBUG'
    try {
        $userProfile = $env:USERPROFILE
        $script:configDir = Join-Path -Path $userProfile -ChildPath ".optimus" # Make configDir script-scoped
        $serverFile = Join-Path -Path $script:configDir -ChildPath "servers.json"
        $modelFile = Join-Path -Path $script:configDir -ChildPath "model.config"

        if (-not (Test-Path -Path $script:configDir)) { New-Item -Path $script:configDir -ItemType Directory -Force | Out-Null }
        if (-not (Test-Path -Path $serverFile)) { Set-Content -Path $serverFile -Value "[]" | Out-Null }
        if (-not (Test-Path -Path $modelFile)) { Set-Content -Path $modelFile -Value "gemini-1.5-flash-latest" | Out-Null } # Default model

        $script:OptimusConfig = @{
            ServerFile = $serverFile
            ModelFile  = $modelFile
        }
        Write-Log -Message "Configuration initialized successfully." -Level 'INFO'
        return $true
    }
    catch { Write-Log -Message "Could not initialize configuration: $($_.Exception.Message)" -Level 'ERROR'; return $false }
}

# --- GUI State and Helper Functions ---
function Invoke-ApiKeyPrompt {
    Write-Log -Message "Creating API Key prompt..." -Level 'DEBUG'
    $InputForm = New-Object System.Windows.Forms.Form
    $InputForm.Text = "API Key Required"
    $InputForm.Size = New-Object System.Drawing.Size(450, 180)
    $InputForm.StartPosition = 'CenterParent'
    $InputForm.FormBorderStyle = 'FixedDialog'

    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = "To use Optimus, you need a Gemini API key.`nPlease enter your key below:"
    $Label.Location = New-Object System.Drawing.Point(10, 10)
    $Label.Size = New-Object System.Drawing.Size(420, 40)
    $InputForm.Controls.Add($Label)

    $TextBox = New-Object System.Windows.Forms.TextBox
    $TextBox.Location = New-Object System.Drawing.Point(10, 60)
    $TextBox.Size = New-Object System.Drawing.Size(410, 25)
    $InputForm.Controls.Add($TextBox)

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Text = "Save Key"
    $OKButton.Location = New-Object System.Drawing.Point(320, 100)
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $InputForm.Controls.Add($OKButton)
    $InputForm.AcceptButton = $OKButton

    if ($InputForm.ShowDialog() -eq 'OK') {
        $apiKey = $TextBox.Text.Trim()
        if (-not [string]::IsNullOrWhiteSpace($apiKey)) {
            try {
                $apiKeyFile = Join-Path -Path $script:configDir -ChildPath "api.config"
                $secureKey = ConvertTo-SecureString -String $apiKey -AsPlainText -Force
                $secureKey | ConvertFrom-SecureString | Set-Content -Path $apiKeyFile
                Write-Log -Message "API Key has been saved securely." -Level 'SUCCESS'
                return $true
            }
            catch {
                Write-Log -Message "Failed to save API Key: $($_.Exception.Message)" -Level 'ERROR'
                [System.Windows.Forms.MessageBox]::Show("Could not save the API key. Please check permissions.", "Error", "OK", "Error")
                return $false
            }
        }
    }
    # User cancelled
    return $false
}

function Invoke-ModelSelectionPrompt {
    Write-Log -Message "Creating AI Model selection prompt..." -Level 'DEBUG'
    $ModelForm = New-Object System.Windows.Forms.Form
    $ModelForm.Text = "Set AI Model"
    $ModelForm.Size = New-Object System.Drawing.Size(420, 240)
    $ModelForm.StartPosition = 'CenterParent'
    $ModelForm.FormBorderStyle = 'FixedDialog'

    $Label = New-Object System.Windows.Forms.Label; $Label.Text = "Select the Gemini model for all future analyses:"; $Label.Location = New-Object System.Drawing.Point(15, 15); $Label.AutoSize = $true
    
    $Group = New-Object System.Windows.Forms.GroupBox; $Group.Location = New-Object System.Drawing.Point(15, 40); $Group.Size = New-Object System.Drawing.Size(370, 110)

    $Radio1 = New-Object System.Windows.Forms.RadioButton; $Radio1.Text = "Gemini 1.5 Flash (Fastest, general use)"; $Radio1.Tag = "gemini-1.5-flash-latest"; $Radio1.Location = New-Object System.Drawing.Point(15, 20); $Radio1.AutoSize = $true
    $Radio2 = New-Object System.Windows.Forms.RadioButton; $Radio2.Text = "Gemini 2.5 Flash (Next-gen speed)"; $Radio2.Tag = "gemini-2.5-flash"; $Radio2.Location = New-Object System.Drawing.Point(15, 45); $Radio2.AutoSize = $true
    $Radio3 = New-Object System.Windows.Forms.RadioButton; $Radio3.Text = "Gemini 2.5 Pro (Most powerful)"; $Radio3.Tag = "gemini-2.5-pro"; $Radio3.Location = New-Object System.Drawing.Point(15, 70); $Radio3.AutoSize = $true
    $Group.Controls.AddRange(@($Radio1, $Radio2, $Radio3))

    # Pre-select the current model
    $currentModel = Get-Content -Path $script:OptimusConfig.ModelFile -ErrorAction SilentlyContinue
    foreach($radio in @($Radio1, $Radio2, $Radio3)){ if($radio.Tag -eq $currentModel){ $radio.Checked = $true } }

    $OKButton = New-Object System.Windows.Forms.Button; $OKButton.Text = "Save"; $OKButton.Location = New-Object System.Drawing.Point(300, 165); $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $ModelForm.Controls.AddRange(@($Label, $Group, $OKButton)); $ModelForm.AcceptButton = $OKButton

    if ($ModelForm.ShowDialog() -eq 'OK') {
        $selectedModel = $null
        foreach($radio in @($Radio1, $Radio2, $Radio3)){ if($radio.Checked){ $selectedModel = $radio.Tag } }
        if ($selectedModel) {
            try {
                Set-Content -Path $script:OptimusConfig.ModelFile -Value $selectedModel
                Write-Log -Message "AI Model set to '$selectedModel'." -Level 'SUCCESS'
            } catch { Write-Log -Message "Failed to save model selection: $($_.Exception.Message)" -Level 'ERROR' }
        }
    }
}

function Update-RunButtonState {
    Write-Log -Message "Checking for API Key..." -Level 'DEBUG'
    $apiKeyFile = Join-Path -Path $script:configDir -ChildPath "api.config"
    if (Test-Path -Path $apiKeyFile -PathType Leaf) {
        $script:Button_Run.Enabled = $true
        Write-Log -Message "API Key found. Analysis is enabled." -Level 'DEBUG'
    }
    else {
        $script:Button_Run.Enabled = $false
        Write-Log -Message "API Key not found. 'Run Analysis' is disabled." -Level 'WARN'
        if (Invoke-ApiKeyPrompt) {
            # If key was set successfully, re-run this check to enable the button.
            Update-RunButtonState
        } else {
            Write-Log -Message "API key setup was cancelled. Analysis remains disabled." -Level 'INFO'
        }
    }
}

function Invoke-GuiReset {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Action
    )

    $title = "Confirm Reset"
    $buttons = [System.Windows.Forms.MessageBoxButtons]::YesNo
    $icon = [System.Windows.Forms.MessageBoxIcon]::Warning
    $message = ""
    $resetFunc = $null

    switch($Action) {
        "Settings" {
            $message = "Are you sure you want to delete the API key, server list, and model selection?"
            $resetFunc = {
                $configFiles = @( "servers.json", "api.config", "lastpath.config", "model.config" )
                $filesToDelete = $configFiles | ForEach-Object { Join-Path -Path $script:configDir -ChildPath $_ }
                foreach ($file in $filesToDelete) {
                    if (Test-Path -Path $file) { Remove-Item -Path $file -Force }
                }
                Write-Log -Message "Core settings have been reset." -Level 'SUCCESS'
                Set-Content -Path (Join-Path -Path $script:configDir -ChildPath "servers.json") -Value "[]"
                $script:ComboBox_Server.Items.Clear()
            }
        }
        "Reports" {
            $message = "Are you sure you want to delete ALL past analysis reports?"
            $resetFunc = {
                $analysisDir = Join-Path -Path $script:configDir -ChildPath "Analysis"
                if(Test-Path $analysisDir){ Remove-Item -Path $analysisDir -Recurse -Force }
            }
        }
        "Full" {
            $message = "WARNING: This will delete ALL configuration AND all saved analysis reports."
            $icon = [System.Windows.Forms.MessageBoxIcon]::Stop
            $resetFunc = {
                Remove-Item -Path $script:configDir -Recurse -Force
                Write-Log -Message "Full reset complete. Restart the application to re-initialize." -Level 'WARN'
                $script:ComboBox_Server.Items.Clear()
            }
        }
    }

    if ([System.Windows.Forms.MessageBox]::Show($message, $title, $buttons, $icon) -eq 'Yes') {
        try {
            & $resetFunc
            Update-RunButtonState
        }
        catch { Write-Log -Message "Reset operation failed: $($_.Exception.Message)" -Level 'ERROR' }
    } else {
        Write-Log -Message "Reset operation cancelled by user." -Level 'INFO'
    }
}

# SQL Server Functions
function Test-SqlServerConnection {
    param([string]$ServerInstance)
    Write-Log -Message "Testing connection to '$ServerInstance'..." -Level 'INFO'
    try {
        if (-not (Get-Module -Name SqlServer)) { Import-Module SqlServer -ErrorAction Stop }
        Invoke-Sqlcmd -ServerInstance $ServerInstance -Query "SELECT 1" -QueryTimeout 5 -TrustServerCertificate -ErrorAction Stop | Out-Null
        Write-Log -Message "Connection successful!" -Level 'SUCCESS'
        return $true
    }
    catch { Write-Log -Message "Failed to connect to '$ServerInstance': $($_.Exception.Message)" -Level 'ERROR'; return $false }
}
#endregion

#region GUI Construction

# --- Help Window Function ---
function Show-HelpWindow {
    param([string]$Title, [string]$Content)
    $HelpForm = New-Object System.Windows.Forms.Form; $HelpForm.Text = $Title; $HelpForm.Size = New-Object System.Drawing.Size(700, 500); $HelpForm.StartPosition = 'CenterParent'; $HelpForm.MinimizeBox = $false
    $HelpTextBox = New-Object System.Windows.Forms.TextBox; $HelpTextBox.Dock = 'Fill'; $HelpTextBox.Multiline = $true; $HelpTextBox.Scrollbars = 'Vertical'; $HelpTextBox.ReadOnly = $true; $HelpTextBox.Font = New-Object System.Drawing.Font("Consolas", 9); $HelpTextBox.Text = $Content
    $HelpForm.Controls.Add($HelpTextBox); [void]$HelpForm.ShowDialog()
}

# --- Main Form ---
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Optimus - T-SQL Tuning Advisor v3.0" # Updated Version
$Form.Size = New-Object System.Drawing.Size(600, 580)
$Form.StartPosition = 'CenterScreen'
$Form.FormBorderStyle = 'FixedSingle'
$Form.MaximizeBox = $false

# --- Menu Bar ---
$MenuStrip = New-Object System.Windows.Forms.MenuStrip
$OptionsMenu = New-Object System.Windows.Forms.ToolStripMenuItem("&Options")
$HelpMenu = New-Object System.Windows.Forms.ToolStripMenuItem("&Help")

# Define individual menu items
$SetModelItem = New-Object System.Windows.Forms.ToolStripMenuItem("Set AI Model...")
$SetApiKeyItem = New-Object System.Windows.Forms.ToolStripMenuItem("Set/Reset API Key...")

# Create the "Reset Configuration" sub-menu and add the API key item to it
$ResetMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Reset Configuration")
$ResetSettingsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Reset Settings Only...")
$ResetReportsItem = New-Object System.Windows.Forms.ToolStripMenuItem("Remove All Analysis Reports...")
$ResetFullItem = New-Object System.Windows.Forms.ToolStripMenuItem("Full Reset (Settings & Reports)...")
$ResetMenu.DropDownItems.AddRange(@(
    $SetApiKeyItem,
    (New-Object System.Windows.Forms.ToolStripSeparator),
    $ResetSettingsItem,
    $ResetReportsItem,
    (New-Object System.Windows.Forms.ToolStripSeparator),
    $ResetFullItem
))

# Populate the main "Options" menu
$OptionsMenu.DropDownItems.AddRange(@(
    $SetModelItem,
    (New-Object System.Windows.Forms.ToolStripSeparator),
    $ResetMenu
))

# Help Menu Items
$HelpOptimusItem = New-Object System.Windows.Forms.ToolStripMenuItem("View Optimus Engine Help...")
$HelpGuiItem = New-Object System.Windows.Forms.ToolStripMenuItem("View GUI Help...")
$HelpMenu.DropDownItems.AddRange(@($HelpOptimusItem, $HelpGuiItem))

# Add top-level menus to the strip
$MenuStrip.Items.AddRange(@($OptionsMenu, $HelpMenu))
$Form.Controls.Add($MenuStrip); $Form.MainMenuStrip = $MenuStrip

# --- Server Controls ---
$Label_Server = New-Object System.Windows.Forms.Label; $Label_Server.Text = "SQL Server Instance:"; $Label_Server.Location = New-Object System.Drawing.Point(20, 50); $Label_Server.AutoSize = $true
$ComboBox_Server = New-Object System.Windows.Forms.ComboBox; $ComboBox_Server.Location = New-Object System.Drawing.Point(150, 45); $ComboBox_Server.Size = New-Object System.Drawing.Size(280, 25); $ComboBox_Server.DropDownStyle = 'DropDownList'
$script:ComboBox_Server = $ComboBox_Server
$Button_AddServer = New-Object System.Windows.Forms.Button; $Button_AddServer.Text = "Add New..."; $Button_AddServer.Location = New-Object System.Drawing.Point(440, 44); $Button_AddServer.Size = New-Object System.Drawing.Size(120, 25)
$Form.Controls.AddRange(@($Label_Server, $ComboBox_Server, $Button_AddServer))

# --- Input Source Controls ---
$GroupBox_Input = New-Object System.Windows.Forms.GroupBox; $GroupBox_Input.Text = "Input Source"; $GroupBox_Input.Location = New-Object System.Drawing.Point(20, 85); $GroupBox_Input.Size = New-Object System.Drawing.Size(540, 120)
$Radio_Files = New-Object System.Windows.Forms.RadioButton; $Radio_Files.Text = "Select .SQL File(s)"; $Radio_Files.Location = New-Object System.Drawing.Point(15, 25); $Radio_Files.AutoSize = $true; $Radio_Files.Checked = $true
$Radio_Folder = New-Object System.Windows.Forms.RadioButton; $Radio_Folder.Text = "Select Folder"; $Radio_Folder.Location = New-Object System.Drawing.Point(180, 25); $Radio_Folder.AutoSize = $true
$Radio_Adhoc = New-Object System.Windows.Forms.RadioButton; $Radio_Adhoc.Text = "Ad-hoc T-SQL"; $Radio_Adhoc.Location = New-Object System.Drawing.Point(320, 25); $Radio_Adhoc.AutoSize = $true
$TextBox_Input = New-Object System.Windows.Forms.TextBox; $TextBox_Input.Location = New-Object System.Drawing.Point(15, 55); $TextBox_Input.Size = New-Object System.Drawing.Size(395, 25)
$Button_Browse = New-Object System.Windows.Forms.Button; $Button_Browse.Text = "Browse..."; $Button_Browse.Location = New-Object System.Drawing.Point(420, 54); $Button_Browse.Size = New-Object System.Drawing.Size(100, 25)
$GroupBox_Input.Controls.AddRange(@($Radio_Files, $Radio_Folder, $Radio_Adhoc, $TextBox_Input, $Button_Browse)); $Form.Controls.Add($GroupBox_Input)

# --- Settings Controls ---
$GroupBox_Settings = New-Object System.Windows.Forms.GroupBox; $GroupBox_Settings.Text = "Settings"; $GroupBox_Settings.Location = New-Object System.Drawing.Point(20, 215); $GroupBox_Settings.Size = New-Object System.Drawing.Size(540, 90)
$CheckBox_ActualPlan = New-Object System.Windows.Forms.CheckBox; $CheckBox_ActualPlan.Text = "Use 'Actual' Execution Plan (will execute query)"; $CheckBox_ActualPlan.Location = New-Object System.Drawing.Point(15, 25); $CheckBox_ActualPlan.AutoSize = $true
$CheckBox_OpenTunedFile = New-Object System.Windows.Forms.CheckBox; $CheckBox_OpenTunedFile.Text = "Open tuned file(s) upon completion"; $CheckBox_OpenTunedFile.Location = New-Object System.Drawing.Point(15, 55); $CheckBox_OpenTunedFile.AutoSize = $true; $CheckBox_OpenTunedFile.Checked = $true
$CheckBox_DebugMode = New-Object System.Windows.Forms.CheckBox; $CheckBox_DebugMode.Text = "Run in Debug Mode"; $CheckBox_DebugMode.Location = New-Object System.Drawing.Point(360, 25); $CheckBox_DebugMode.AutoSize = $true
$GroupBox_Settings.Controls.AddRange(@($CheckBox_ActualPlan, $CheckBox_OpenTunedFile, $CheckBox_DebugMode)); $Form.Controls.Add($GroupBox_Settings)

# --- Status Box ---
$Label_Status = New-Object System.Windows.Forms.Label; $Label_Status.Text = "GUI Status Log:"; $Label_Status.Location = New-Object System.Drawing.Point(20, 315); $Label_Status.AutoSize = $true
$TextBox_Status = New-Object System.Windows.Forms.RichTextBox; $TextBox_Status.Location = New-Object System.Drawing.Point(20, 335); $TextBox_Status.Size = New-Object System.Drawing.Size(540, 150); $TextBox_Status.Scrollbars = 'Vertical'; $TextBox_Status.ReadOnly = $true; $TextBox_Status.Font = New-Object System.Drawing.Font("Consolas", 8); $TextBox_Status.BackColor = "Black"
$script:LogTextBox = $TextBox_Status; $Form.Controls.AddRange(@($Label_Status, $TextBox_Status))

# --- Action Buttons ---
$Button_Run = New-Object System.Windows.Forms.Button; $Button_Run.Text = "Run Analysis"; $Button_Run.Location = New-Object System.Drawing.Point(450, 495); $Button_Run.Size = New-Object System.Drawing.Size(110, 30)
$script:Button_Run = $Button_Run # Make button script-scoped
$Form.Controls.Add($Button_Run)

#endregion

#region GUI Event Handlers

# --- Form Load Event ---
$Form.Add_Load({
    Write-Log -Message "Welcome to Optimus GUI." -Level 'INFO'
    if (Initialize-Configuration) {
        try {
            [array]$servers = Get-Content -Path $script:OptimusConfig.ServerFile | ConvertFrom-Json
            if ($servers.Count -gt 0) {
                $ComboBox_Server.Items.AddRange($servers)
                $ComboBox_Server.SelectedIndex = 0 | Out-Null
                Write-Log -Message "Loaded $($servers.Count) server(s) from configuration." -Level 'SUCCESS'
            } else { Write-Log -Message "Server list is empty. Use 'Add New...' to add a server." -Level 'WARN' }
        } catch { Write-Log -Message "Could not load servers from $($script:OptimusConfig.ServerFile)." -Level 'ERROR' }
    }
    # Check for API key and set button state
    Update-RunButtonState
})

# --- Menu Item Click Events ---
$SetModelItem.Add_Click({ Invoke-ModelSelectionPrompt })
$SetApiKeyItem.Add_Click({ Update-RunButtonState }) # Re-running the update function will trigger the prompt if the key is missing
$ResetSettingsItem.Add_Click({ Invoke-GuiReset -Action "Settings" })
$ResetReportsItem.Add_Click({ Invoke-GuiReset -Action "Reports" })
$ResetFullItem.Add_Click({ Invoke-GuiReset -Action "Full" })
$HelpOptimusItem.Add_Click({
    $scriptPath = Join-Path $PSScriptRoot "Optimus.ps1"
    if (Test-Path $scriptPath) {
        $helpContent = Get-Help -Full $scriptPath | Out-String
        Show-HelpWindow -Title "Optimus.ps1 Engine Help" -Content $helpContent
    } else {
        [System.Windows.Forms.MessageBox]::Show("Could not find Optimus.ps1 in the same directory.", "Error", "OK", "Error")
    }
})
$HelpGuiItem.Add_Click({
    $scriptPath = $PSCommandPath
    $helpContent = Get-Help -Full $scriptPath | Out-String
    Show-HelpWindow -Title "Optimus GUI Help" -Content $helpContent
})


# --- Button Click Events ---
$Button_AddServer.Add_Click({
    $InputBox = New-Object System.Windows.Forms.Form; $InputBox.Text = "Add New SQL Server"; $InputBox.Size = New-Object System.Drawing.Size(400, 130); $InputBox.StartPosition = 'CenterParent'; $InputBox.FormBorderStyle = 'FixedDialog'
    $Label = New-Object System.Windows.Forms.Label; $Label.Text = "Enter Server Name (e.g., HOST\INSTANCE):"; $Label.Location = New-Object System.Drawing.Point(10, 10); $Label.AutoSize = $true
    $TextBox = New-Object System.Windows.Forms.TextBox; $TextBox.Location = New-Object System.Drawing.Point(10, 35); $TextBox.Size = New-Object System.Drawing.Size(360, 25)
    $OKButton = New-Object System.Windows.Forms.Button; $OKButton.Text = "OK"; $OKButton.Location = New-Object System.Drawing.Point(280, 65); $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $InputBox.Controls.AddRange(@($Label, $TextBox, $OKButton)); $InputBox.AcceptButton = $OKButton
    if ($InputBox.ShowDialog() -eq 'OK') {
        $newServer = $TextBox.Text.Trim()
        if (-not [string]::IsNullOrWhiteSpace($newServer)) {
            if (Test-SqlServerConnection -ServerInstance $newServer) {
                if ($ComboBox_Server.Items -notcontains $newServer) {
                    $ComboBox_Server.Items.Add($newServer); ($ComboBox_Server.Items | ForEach-Object { $_ }) | ConvertTo-Json -Depth 5 | Set-Content -Path $script:OptimusConfig.ServerFile
                    Write-Log -Message "Server '$newServer' added and saved." -Level 'SUCCESS'
                }
                $ComboBox_Server.SelectedItem = $newServer
            }
        } else { Write-Log -Message "Server name cannot be empty." -Level 'WARN' }
    }
})

$Button_Browse.Add_Click({
    if ($Radio_Files.Checked) {
        $FileDialog = New-Object System.Windows.Forms.OpenFileDialog; $FileDialog.Title = "Select SQL File(s)"; $FileDialog.Filter = "SQL Files (*.sql)|*.sql|All files (*.*)|*.*"; $FileDialog.Multiselect = $true
        if ($FileDialog.ShowDialog() -eq 'OK') {
            # Format the file list as a comma-separated list of quoted strings for the command line
            $fileList = $FileDialog.FileNames | ForEach-Object { "'$_'" }
            $TextBox_Input.Text = $fileList -join ", "
        }
    } elseif ($Radio_Folder.Checked) {
        $FolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog; $FolderDialog.Description = "Select a folder containing .sql files"
        if ($FolderDialog.ShowDialog() -eq 'OK') { $TextBox_Input.Text = $FolderDialog.SelectedPath }
    }
})

$Button_Run.Add_Click({
    Write-Log -Message "--- Preparing for Analysis ---" -Level 'INFO'
    $hasError = $false

    # 1. Validation Checks
    if ([string]::IsNullOrWhiteSpace($ComboBox_Server.SelectedItem)) { Write-Log -Message "Validation Failed: Please select a SQL Server." -Level 'ERROR'; $hasError = $true }
    if ([string]::IsNullOrWhiteSpace($TextBox_Input.Text) -and -not $Radio_Adhoc.Checked) { Write-Log -Message "Validation Failed: Please provide an input source." -Level 'ERROR'; $hasError = $true }
    if ($Radio_Adhoc.Checked -and [string]::IsNullOrWhiteSpace($TextBox_Input.Text)) { Write-Log -Message "Validation Failed: Ad-hoc T-SQL input cannot be empty." -Level 'ERROR'; $hasError = $true }

    $optimusEnginePath = Join-Path $PSScriptRoot "Optimus.ps1"
    if (-not (Test-Path -Path $optimusEnginePath)) { Write-Log -Message "Validation Failed: Optimus.ps1 not found in the same directory." -Level 'ERROR'; $hasError = $true }

    if ($hasError) { Write-Log -Message "Please correct the errors before running." -Level 'WARN'; return }

    # 2. Build the command line arguments
    $commandParts = [System.Collections.Generic.List[string]]::new()
    $commandParts.Add("-ServerName '$($ComboBox_Server.SelectedItem)'")

    if ($Radio_Files.Checked) { $commandParts.Add("-SQLFile $($TextBox_Input.Text)") }
    elseif ($Radio_Folder.Checked) { $commandParts.Add("-FolderPath '$($TextBox_Input.Text)'") }
    else { $commandParts.Add("-AdhocSQL '$($TextBox_Input.Text -replace "'", "''")'") }

    if ($CheckBox_ActualPlan.Checked) { $commandParts.Add("-UseActualPlan") }
    if ($CheckBox_OpenTunedFile.Checked) { $commandParts.Add("-OpenTunedFile") }
    if ($CheckBox_DebugMode.Checked) { $commandParts.Add("-DebugMode") }

    # 3. Construct and Execute
    $fullCommand = "& `"$optimusEnginePath`" $($commandParts -join ' ')"
    Write-Log -Message "Launching Optimus engine..." -Level 'SUCCESS'
    Write-Log -Message "Command: powershell.exe -NoExit -Command `"$fullCommand`"" -Level 'DEBUG'

    try {
        $startProcessArgs = @{ FilePath = 'powershell.exe'; ArgumentList = "-NoExit", "-Command", $fullCommand; ErrorAction = 'Stop' }
        if (-not $CheckBox_DebugMode.Checked) {
            $startProcessArgs['WindowStyle'] = 'Minimized'
            Write-Log -Message "Engine will start in a minimized window." -Level 'INFO'
        } else { Write-Log -Message "Engine will start in a normal window for debugging." -Level 'INFO' }
        Start-Process @startProcessArgs
        Write-Log -Message "Engine started successfully." -Level 'SUCCESS'
    }
    catch { Write-Log -Message "Failed to launch Optimus engine: $($_.Exception.Message)" -Level 'ERROR' }
})

# --- Radio Button Change Event ---
$radio_event_handler = {
    if ($this.Checked) {
        $TextBox_Input.Clear()
        if ($Radio_Adhoc.Checked) {
            $TextBox_Input.Multiline = $true; $TextBox_Input.Scrollbars = 'Vertical'; $TextBox_Input.Size = New-Object System.Drawing.Size(505, 50); $Button_Browse.Enabled = $false
            Write-Log -Message "Switched to Ad-hoc input mode." -Level 'DEBUG'
        } else {
            $TextBox_Input.Multiline = $false; $TextBox_Input.Scrollbars = 'None'; $TextBox_Input.Size = New-Object System.Drawing.Size(395, 25); $Button_Browse.Enabled = $true
            Write-Log -Message "Switched to File/Folder input mode." -Level 'DEBUG'
        }
    }
}
$Radio_Files.Add_CheckedChanged($radio_event_handler)
$Radio_Folder.Add_CheckedChanged($radio_event_handler)
$Radio_Adhoc.Add_CheckedChanged($radio_event_handler)

#endregion

# --- Show the Form ---
try {
    [void]$Form.ShowDialog()
}
catch {
    Write-Warning "An error occurred while running the GUI."
    Write-Warning $_.Exception.Message
}