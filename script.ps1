<# 
    Rebuild-StartMenuShortcuts.ps1
    Windows 11 / Windows 10

    Goal:
      Discover classic Win32 apps still installed and (re)create Start menu shortcuts.
      - Leverages App Paths + Uninstall registry, with safe parsing.
      - Heuristic scan of common folders (optional, capped).
      - Skips uninstallers/updaters/helpers with simple rules.
      - Avoids duplicates if a shortcut already exists.
      - Interactive picker (Out-GridView if present, else console), or Auto mode.

    Notes:
      - UWP/Store apps are NOT handled (they don’t use .lnk the same way).
      - Use -AllUsers to write under ProgramData (requires admin).
#>

[CmdletBinding()]
param(
    [switch]$AllUsers,        # Write to ProgramData\...\Programs (needs admin)
    [switch]$Interactive,     # Let you pick items
    [switch]$CheckboxUI,      # Use a checkbox grid with rename support
    [switch]$Auto,            # Create for all candidates without prompt
    [switch]$Preview,         # Show what would be done, exit
    [int]$ScanLimit = 5000,   # Max EXEs to inspect during folder scan
    [string[]]$IncludeRoots,  # Optional extra roots to scan (e.g., "D:\Apps","D:\Games")
    [string]$Subfolder = 'Recovered' # Subfolder under Programs for new shortcuts
)

# ------------------------- Helpers -------------------------

function Test-Admin {
    try {
        $id = [Security.Principal.WindowsIdentity]::GetCurrent()
        $p = New-Object Security.Principal.WindowsPrincipal($id)
        return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { return $false }
}

function Normalize-Name {
    param([Parameter(Mandatory)][string]$PathOrName)
    if (Test-Path $PathOrName) {
        return ([System.IO.Path]::GetFileNameWithoutExtension($PathOrName))
    } else {
        return $PathOrName
    }
}

function Get-ProgramsFolder {
    param([switch]$AllUsers)
    if ($AllUsers) {
        return Join-Path $env:ProgramData "Microsoft\Windows\Start Menu\Programs"
    } else {
        return Join-Path $env:AppData    "Microsoft\Windows\Start Menu\Programs"
    }
}

function Get-ExistingShortcutNames {
    param([string]$ProgramsFolder)
    if (-not (Test-Path $ProgramsFolder)) { return @() }
    Get-ChildItem $ProgramsFolder -Recurse -Filter *.lnk -ErrorAction SilentlyContinue |
        ForEach-Object { $_.BaseName.ToLower() } |
        Select-Object -Unique
}

# Parse messy DisplayIcon values: quotes, ,index, extra junk after .exe, or wrong file (like .ico)
function Resolve-ExeFromIcon {
    param([string]$DisplayIcon)

    if ([string]::IsNullOrWhiteSpace($DisplayIcon)) { return $null }

    $icon = $DisplayIcon.Trim()            # whitespace
    $icon = $icon.Trim('"')                # surrounding quotes
    $icon = $icon -replace ',\s*\d+$',''   # strip trailing ,index
    $icon = $icon -replace '\s+$',''       # trailing spaces

    # If there's a .exe somewhere, capture up to .exe
    $lower = $icon.ToLower()
    $exe = $null
    $idx = $lower.IndexOf('.exe')
    if ($idx -ge 0) {
        $exe = $icon.Substring(0, $idx + 4)
    } else {
        # No .exe; sometimes DisplayIcon points to .ico; try to switch .ico -> .exe heuristically
        if ($lower.EndsWith('.ico')) {
            $exe = $icon.Substring(0, $icon.Length - 4) + '.exe'
        }
    }

    if (-not $exe) { return $null }

    # Clean remaining quotes/spaces
    $exe = $exe.Trim().Trim('"')

    try {
        if (Test-Path -LiteralPath $exe) { 
            return (Resolve-Path -LiteralPath $exe).Path 
        }
    } catch {
        # swallow malformed cases
    }
    return $null
}

# If we only have an InstallLocation, try to pick the "primary" exe in that folder
function Find-PrimaryExeInFolder {
    param(
        [Parameter(Mandatory)][string]$Folder,
        [string]$HintName
    )
    if (-not (Test-Path $Folder)) { return $null }

    $candidates = Get-ChildItem -Path $Folder -Recurse -ErrorAction SilentlyContinue -Filter *.exe |
        Where-Object {
            $_.Length -gt 200KB -and
            $_.Name -as [string] -and
            ($_.Name -notmatch '(?i)(unins|uninstall|setup|install|updater|update|elevation|crashpad|service|daemon|agent|reporter|telemetry|console|cli|helper|tool|dbg|redist|vc_redist|dxsetup|repair|cleanup|watchdog|monitor)\.exe$')
        }

    if (-not $candidates) { return $null }

    if ($HintName) {
        $byName = $candidates | Where-Object { $_.BaseName -match [Regex]::Escape($HintName) } |
                  Sort-Object Length -Descending | Select-Object -First 1
        if ($byName) { return $byName.FullName }
    }

    # As a fallback, pick the largest .exe (often the main GUI)
    return ($candidates | Sort-Object Length -Descending | Select-Object -First 1).FullName
}

function Is-LauncherCandidate {
    param([System.IO.FileInfo]$File)
    if (-not $File) { return $false }
    if ($File.Length -lt 200KB) { return $false }

    $n = $File.Name.ToLower()
    $skipPatterns = @(
        'unins', 'uninstall', 'setup', 'install', 'updater', 'update', 'elevation',
        'crashpad', 'service', 'daemon', 'agent', 'reporter', 'telemetry',
        'console', 'cli', 'helper', 'tool', 'dbg', 'redist', 'vc_redist',
        'dxsetup', 'repair', 'cleanup', 'watchdog', 'monitor'
    )
    foreach ($pat in $skipPatterns) {
        if ($n -match [Regex]::Escape($pat)) { return $false }
    }
    return $true
}

# ------------------------- Shortcut creation -------------------------

function New-StartShortcut {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Exe,
        [Parameter(Mandatory)][string]$ProgramsFolder
    )

    if (-not (Test-Path -LiteralPath $Exe)) {
        Write-Warning "Target exe not found: $Exe"
        return
    }

    if (-not (Test-Path -LiteralPath $ProgramsFolder)) {
        New-Item -ItemType Directory -Force -Path $ProgramsFolder | Out-Null
    }

    $safeName = $Name -replace '[\\/:*?"<>|]', ' '
    $lnkPath = Join-Path $ProgramsFolder "$safeName.lnk"

    if ($PSCmdlet.ShouldProcess($lnkPath, "Create Start menu shortcut")) {
        try {
            $shell = New-Object -ComObject WScript.Shell
            $sc = $shell.CreateShortcut($lnkPath)
            $sc.TargetPath = $Exe
            $sc.WorkingDirectory = Split-Path $Exe -Parent
            $sc.IconLocation = "$Exe,0"
            $sc.Save()
            Write-Host "✅ Created: $lnkPath"
        } catch {
            Write-Warning "Failed to create shortcut for '$Name' -> $Exe. $_"
        }
    }
}

# ------------------------- GUI selection (checkbox + rename) -------------------------

function Select-AppsWithCheckboxUI {
    param(
        [Parameter(Mandatory)][object[]]$Candidates
    )

    try {
        Add-Type -AssemblyName System.Windows.Forms | Out-Null
        Add-Type -AssemblyName System.Drawing | Out-Null
    } catch {
        return $null
    }

    # WinForms requires STA; if not STA, warn so the user can relaunch with -STA
    try {
        if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
            Write-Warning 'Checkbox UI requires STA. Relaunch the script with -STA (e.g., powershell -STA -File script.ps1 -Interactive -CheckboxUI).'
            return $null
        }
    } catch {}

    try { [System.Windows.Forms.Application]::EnableVisualStyles() } catch {}

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Select apps to add to Start (edit names if needed)'
    $form.StartPosition = 'CenterScreen'
    $form.Width = 900
    $form.Height = 600

    # Top search panel
    $panelTop = New-Object System.Windows.Forms.Panel
    $panelTop.Dock = 'Top'
    $panelTop.Height = 36

    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Text = 'Search:'
    $lblSearch.AutoSize = $true
    $lblSearch.Left = 10
    $lblSearch.Top = 10

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Left = 70
    $txtSearch.Top = 6
    $txtSearch.Width = 300

    $btnClearSearch = New-Object System.Windows.Forms.Button
    $btnClearSearch.Text = 'Clear'
    $btnClearSearch.Left = 380
    $btnClearSearch.Top = 5
    $btnClearSearch.Width = 60
    $btnClearSearch.Add_Click({ $txtSearch.Text = '' })

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = 'Fill'
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AutoGenerateColumns = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.EditMode = 'EditOnKeystrokeOrF2'

    # Columns: Checked, Name (editable), Exe (read-only), Source (read-only)
    $colCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colCheck.Name = 'Select'
    $colCheck.HeaderText = 'Select'
    $colCheck.Width = 60

    $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colName.Name = 'Name'
    $colName.HeaderText = 'Shortcut Name'
    $colName.AutoSizeMode = 'Fill'

    $colExe = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colExe.Name = 'Exe'
    $colExe.HeaderText = 'Executable'
    $colExe.Width = 380
    $colExe.ReadOnly = $true

    $colSource = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colSource.Name = 'Source'
    $colSource.HeaderText = 'Source'
    $colSource.Width = 80
    $colSource.ReadOnly = $true

    # Add columns explicitly to avoid AddRange interop quirks
    [void]$grid.Columns.Add($colCheck)
    [void]$grid.Columns.Add($colName)
    [void]$grid.Columns.Add($colExe)
    [void]$grid.Columns.Add($colSource)

    if ($Candidates -and $Candidates.Count -gt 0) {
        foreach ($c in $Candidates) {
            [void]$grid.Rows.Add($false, [string]$c.Name, [string]$c.Exe, [string]$c.Source)
        }
    }

    $panelBottom = New-Object System.Windows.Forms.Panel
    $panelBottom.Dock = 'Bottom'
    $panelBottom.Height = 45

    $btnSelectAll = New-Object System.Windows.Forms.Button
    $btnSelectAll.Text = 'Select All'
    $btnSelectAll.Width = 90
    $btnSelectAll.Left = 10
    $btnSelectAll.Top = 10
    $btnSelectAll.Add_Click({
        foreach ($row in $grid.Rows) { $row.Cells['Select'].Value = $true }
    })

    $btnSelectNone = New-Object System.Windows.Forms.Button
    $btnSelectNone.Text = 'Select None'
    $btnSelectNone.Width = 100
    $btnSelectNone.Left = 110
    $btnSelectNone.Top = 10
    $btnSelectNone.Add_Click({
        foreach ($row in $grid.Rows) { $row.Cells['Select'].Value = $false }
    })

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = 'Add selected to Start Menu'
    $btnOK.Width = 180
    $btnOK.Left = 220
    $btnOK.Top = 10
    $btnOK.Enabled = $false
    $btnOK.Add_Click({ $form.Tag = 'OK'; $form.Close() })

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'
    $btnCancel.Width = 80
    $btnCancel.Left = 410
    $btnCancel.Top = 10
    $btnCancel.Add_Click({ $form.Tag = 'Cancel'; $form.Close() })

    $panelBottom.Controls.AddRange(@($btnSelectAll, $btnSelectNone, $btnOK, $btnCancel))

    $panelTop.Controls.AddRange(@($lblSearch, $txtSearch, $btnClearSearch))

    $form.Controls.Add($grid)
    $form.Controls.Add($panelTop)
    $form.Controls.Add($panelBottom)

    # Enable/disable OK button based on any checked rows
    $updateOkEnabled = {
        $anyChecked = $false
        foreach ($row in $grid.Rows) {
            if ($row.Visible -and [bool]($row.Cells['Select'].Value)) { $anyChecked = $true; break }
        }
        $btnOK.Enabled = $anyChecked
    }

    # Live filter logic
    $applyFilter = {
        $q = ($txtSearch.Text | ForEach-Object { $_.Trim() })
        if (-not $q -or $q.Length -eq 0) {
            foreach ($row in $grid.Rows) { $row.Visible = $true }
        } else {
            $qLower = $q.ToLower()
            foreach ($row in $grid.Rows) {
                $nameVal = [string]$row.Cells['Name'].Value
                $exeVal = [string]$row.Cells['Exe'].Value
                $hit = ($nameVal -and $nameVal.ToLower().Contains($qLower)) -or ($exeVal -and $exeVal.ToLower().Contains($qLower))
                $row.Visible = [bool]$hit
            }
        }
        & $updateOkEnabled
    }

    $txtSearch.Add_TextChanged($applyFilter)

    # Ensure checkbox edits commit immediately so CellValueChanged fires
    $grid.Add_CurrentCellDirtyStateChanged({ if ($grid.IsCurrentCellDirty) { $grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) } })
    $grid.Add_CellValueChanged({ & $updateOkEnabled })
    & $updateOkEnabled

    [void]$form.ShowDialog()

    if ($form.Tag -ne 'OK') { return @() }

    $selected = @()
    foreach ($row in $grid.Rows) {
        $isChecked = [bool]($row.Cells['Select'].Value)
        if ($isChecked) {
            $name = [string]$row.Cells['Name'].Value
            $exe = [string]$row.Cells['Exe'].Value
            $src = [string]$row.Cells['Source'].Value
            $selected += [pscustomobject]@{ Name = $name; Exe = $exe; Source = $src }
        }
    }
    return $selected
}

# ------------------------- Discovery -------------------------

function Get-CandidateApps {
    [CmdletBinding()]
    param(
        [int]$ScanLimit = 5000,
        [string[]]$IncludeRoots
    )

    $candidates = New-Object System.Collections.Generic.List[object]

    # 1) App Paths
    $appPathsRoots = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths"
    )
    foreach ($root in $appPathsRoots) {
        if (Test-Path $root) {
            Get-ChildItem $root -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    $p = Get-ItemProperty $_.PsPath -ErrorAction SilentlyContinue
                    # Default may hold exe; Path sometimes points to folder
                    # Use registry API to read unnamed default value reliably
                    $exe = (Get-Item $_.PsPath).GetValue('')
                    if (-not $exe -and $p.Path) {
                        # If it's a folder, try to find a primary exe
                        if (Test-Path $p.Path) {
                            $exe = Find-PrimaryExeInFolder -Folder $p.Path -HintName $null
                        }
                    }
                    if ($exe -and (Test-Path $exe)) {
                        $name = Normalize-Name -PathOrName $exe
                        $candidates.Add([pscustomobject]@{
                            Name   = $name
                            Exe    = (Resolve-Path $exe).Path
                            Source = 'AppPaths'
                        })
                    }
                } catch {}
            }
        }
    }

    # 2) Uninstall registry
    $uninstRoots = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    foreach ($root in $uninstRoots) {
        if (Test-Path $root) {
            Get-ChildItem $root -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    $p = Get-ItemProperty $_.PsPath -ErrorAction SilentlyContinue
                    $displayName = $p.DisplayName
                    $exe = Resolve-ExeFromIcon -DisplayIcon $p.DisplayIcon

                    # Fallback: some entries have InstallLocation + No DisplayIcon
                    if (-not $exe -and $p.InstallLocation) {
                        $exe = Find-PrimaryExeInFolder -Folder $p.InstallLocation -HintName $displayName
                    }

                    if ($displayName -and $exe) {
                        $candidates.Add([pscustomobject]@{
                            Name   = $displayName
                            Exe    = (Resolve-Path $exe).Path
                            Source = 'Uninstall'
                        })
                    }
                } catch {}
            }
        }
    }

    # 3) Heuristic scan of common roots
    $roots = @(
        $env:ProgramFiles,
        "${env:ProgramFiles(x86)}",
        (Join-Path $env:LOCALAPPDATA "Programs"),
        $env:ProgramData
    ) + ($IncludeRoots | Where-Object { $_ -and (Test-Path $_) })

    $roots = $roots | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique

    $scanned = 0
    foreach ($r in $roots) {
        try {
            Get-ChildItem $r -Recurse -ErrorAction SilentlyContinue -Filter *.exe |
            ForEach-Object {
                if ($scanned -ge $ScanLimit) { break }
                $scanned++
                if (Is-LauncherCandidate -File $_) {
                    $candidates.Add([pscustomobject]@{
                        Name   = $_.BaseName
                        Exe    = $_.FullName
                        Source = 'Scan'
                    })
                }
            }
        } catch {}
    }

    # De-dupe by full exe path (exact); prefer AppPaths and Uninstall over Scan
    $priority = @{ 'AppPaths' = 0; 'Uninstall' = 1; 'Scan' = 2 }
    $deduped = $candidates |
        Group-Object Exe | ForEach-Object {
            $_.Group | Sort-Object @{ Expression = { $priority[[string]$_.Source] }; Descending = $false } | Select-Object -First 1
        }

    # Sort nicely
    $deduped | Sort-Object Name, Exe
}

# ------------------------- Main -------------------------

$programsFolder = Get-ProgramsFolder -AllUsers:$AllUsers
if ($Subfolder -and $Subfolder.Trim().Length -gt 0) {
    $targetProgramsFolder = Join-Path $programsFolder $Subfolder.Trim()
} else {
    $targetProgramsFolder = $programsFolder
}
if ($AllUsers -and -not (Test-Admin)) {
    Write-Warning "You asked for -AllUsers but the shell is not elevated. Run PowerShell as Administrator or omit -AllUsers."
}

# Check duplicates across the entire Programs tree, not just the subfolder
$existing = Get-ExistingShortcutNames -ProgramsFolder $programsFolder

Write-Host "Scanning for installed apps… (this can take a minute)"
$candidates = Get-CandidateApps -ScanLimit $ScanLimit -IncludeRoots $IncludeRoots

# Filter out names we already have as shortcuts (simple name check)
$candidates = $candidates | Where-Object { $existing -notcontains ($_.Name.ToLower()) }

if (-not $candidates -or $candidates.Count -eq 0) {
    Write-Host "✔ No missing Start menu entries found for this user in $programsFolder."
    return
}

if ($Preview) {
    $candidates | Select-Object Name, Exe, Source | Sort-Object Name | Format-Table -AutoSize
    Write-Host "`n(Preview mode) Not creating any shortcuts. Target folder would be: $targetProgramsFolder"
    return
}

# Selection
$selection = $null

if ($Interactive) {
    $ogv = Get-Command Out-GridView -ErrorAction SilentlyContinue
    if ($CheckboxUI) {
        $selection = Select-AppsWithCheckboxUI -Candidates $candidates
    } elseif ($ogv) {
        try {
            $selection = $candidates | Select-Object Name,Exe,Source |
                Out-GridView -Title "Select apps to create Start menu shortcuts" -PassThru
        } catch {
            $selection = $null
        }
    }

    if (-not $selection -and -not $CheckboxUI) {
        # Console fallback
        $i = 0
        $indexed = $candidates | ForEach-Object {
            [pscustomobject]@{ Index = $i; Name = $_.Name; Exe = $_.Exe; Source = $_.Source }
            $i++
        }
        $indexed | Format-Table -AutoSize
        $inputIdx = Read-Host "Enter comma-separated indexes to add (e.g. 0,2,5), or press Enter to cancel. To rename, append =NewName (e.g. 2=My App)"
        if ([string]::IsNullOrWhiteSpace($inputIdx)) { return }
        $parts = $inputIdx -split ',' | ForEach-Object { $_.Trim() }
        $selection = @()
        foreach ($p in $parts) {
            if ($p -match '^(\d+)\s*=\s*(.+)$') {
                $k = [int]$Matches[1]
                $newName = $Matches[2]
                $item = $indexed | Where-Object { $_.Index -eq $k }
                if ($item) { $selection += [pscustomobject]@{ Name = $newName; Exe = $item.Exe; Source = $item.Source } }
            } elseif ($p -match '^\d+$') {
                $k = [int]$p
                $item = $indexed | Where-Object { $_.Index -eq $k } | Select-Object Name,Exe,Source
                if ($item) { $selection += $item }
            }
        }
    }
} elseif ($Auto) {
    $selection = $candidates
} else {
    # Confirm all by default
    $show = $candidates | Select-Object -First 20
    Write-Host "Found $($candidates.Count) candidate apps without Start entries. Showing first 20:"
    $show | Select-Object Name,Exe,Source | Format-Table -AutoSize
    $go = Read-Host "Create shortcuts for ALL of them? (y/n)"
    if ($go -notmatch '^[Yy]') { return }
    $selection = $candidates
}

if (-not $selection -or $selection.Count -eq 0) {
    Write-Host "Nothing selected."
    return
}

Write-Host "Creating shortcuts in: $targetProgramsFolder"
foreach ($app in $selection) {
    New-StartShortcut -Name $app.Name -Exe $app.Exe -ProgramsFolder $targetProgramsFolder
}

Write-Host "`nDone. Open Start → All apps to verify. If entries don’t show immediately, sign out/in or restart Windows Explorer (Task Manager → Windows Explorer → Restart)."
