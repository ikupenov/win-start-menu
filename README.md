## Start Menu Rebuild

Rebuild classic Start menu shortcuts for installed apps on Windows 10/11.

### What it does
- Scans registry (App Paths, Uninstall) and common folders to discover installed apps
- Skips uninstallers/updaters/helpers; avoids duplicates; safe heuristic parsing
- Lets you pick apps interactively (checkbox UI) or run fully automatic
- Can write shortcuts for current user or for all users (admin required)

### Requirements
- Windows 10 or Windows 11
- PowerShell 7 (`pwsh`) preferred, Windows PowerShell also supported
- Administrator only required when using `-AllUsers`

### Quick start
1) Double‑click `run.cmd` (prefers PowerShell 7, falls back to Windows PowerShell):
   - Opens interactive UI by default

Or run directly from a terminal (PowerShell 7):

```powershell
pwsh -NoProfile -STA -ExecutionPolicy Bypass -File .\script.ps1 -Interactive -UI
```

### Usage
The main script is `script.ps1` with the following switches/parameters:

- **-AllUsers**: Create shortcuts under `C:\ProgramData\Microsoft\Windows\Start Menu\Programs` (requires admin)
- **-Interactive**: Interactive selection of discovered apps
- **-UI**: Use an interactive UI with checkbox grid and rename support
- **-Auto**: Create shortcuts for all discovered candidates without prompting
- **-Preview**: Show what would be created, then exit without making changes
- **-ScanLimit <int>**: Cap the maximum number of `.exe` files inspected during folder scan (default 5000)
- **-IncludeRoots <string[]>**: Extra folders to scan (e.g., `"D:\Apps","D:\Games"`)
- **-Subfolder <string>**: Subfolder under Programs where new shortcuts are placed (default `Recovered`)

### Examples
- Interactive checkbox UI (recommended):
```powershell
pwsh -NoProfile -STA -ExecutionPolicy Bypass -File .\script.ps1 -Interactive -UI
```

- Dry‑run to preview actions only:
```powershell
pwsh -NoProfile -STA -ExecutionPolicy Bypass -File .\script.ps1 -Preview
```

- Automatically create shortcuts for all candidates (current user):
```powershell
pwsh -NoProfile -STA -ExecutionPolicy Bypass -File .\script.ps1 -Auto
```

- Create for all users (requires elevated PowerShell):
```powershell
Start-Process pwsh -Verb RunAs -ArgumentList '-NoProfile','-STA','-ExecutionPolicy','Bypass','-File','".\\script.ps1"','-Interactive','-UI','-AllUsers'
```

- Include extra scan roots and adjust subfolder name:
```powershell
pwsh -NoProfile -STA -ExecutionPolicy Bypass -File .\script.ps1 -Interactive -UI -IncludeRoots "D:\\Apps","D:\\Games" -Subfolder "Recovered Apps"
```

### Notes and limitations
- UWP/Store apps are not handled (they do not use `.lnk` shortcuts the same way)
- The script avoids obvious non‑launchers (e.g., uninstallers, updaters, services) using name/size heuristics

### Troubleshooting
- **"Neither PowerShell 7 nor Windows PowerShell was found" when running `run.cmd`**: Ensure `pwsh` or `powershell` is in PATH.
- **Execution policy prompts**: The runner uses `-ExecutionPolicy Bypass` for the session. If restricted by enterprise policy, run in an elevated prompt as permitted by your environment.
- **No checkbox grid UI**: The script falls back to a console picker if the richer grid UI components are unavailable.
- **Access denied when using `-AllUsers`**: Start PowerShell as Administrator.

### License
MIT
