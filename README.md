**Planning script and instructions**

I need to create a markdown file with instructions and a single script. The script should have variables for paths, like $PandocExe, $ToolsRoot, $HelperScriptPath, and $LogPath. I'll keep it concise but complete, with headings and bullet points. The script will set up a context menu for "Open Markdown in Word." The user only asked for installation, so I'll focus on converting MD to DOCX. I'll also include verification steps. I'll respond with markdown content only.
```markdown
# Markdown ↔ Word Right‑Click Integration

This document describes how to install a **right‑click “Open Markdown in Word”** workflow using Pandoc and Word, with logging and a single PowerShell installer script.

---

## 1. Prerequisites

- **Pandoc installed**
  - Verify:

    ```powershell
    pandoc -v
    ```

  - If this fails, install Pandoc (e.g., standard Windows installer) and note the full path to `pandoc.exe` (e.g. `C:\Program Files\Pandoc\pandoc-3.8.3\pandoc.exe`).

- **Microsoft Word installed**
  - Verify you can open Word normally from the Start menu.

---

## 2. Define your installation layout

Decide the following paths (you can keep these defaults):

- **Helper tools root folder** (writable, non‑protected):

  ```text
  C:\Tools\Pandoc
  ```

- **Helper script path**:

  ```text
  C:\Tools\Pandoc\OpenMdInWord.ps1
  ```

- **Log file path**:

  ```text
  C:\Tools\Pandoc\OpenMdInWord.log
  ```

- **Pandoc executable path** (adjust to your system):

  ```text
  C:\Program Files\Pandoc\pandoc-3.8.3\pandoc.exe
  ```

You will plug these into the installer script below.

---

## 3. One‑shot PowerShell installer script

> **Action:** Copy this entire script into a file named `Install-MarkdownWordIntegration.ps1`.  
> **Action:** Open **PowerShell as Administrator**.  
> **Action:** Run:
>
> ```powershell
> .\Install-MarkdownWordIntegration.ps1
> ```

```powershell
# Install-MarkdownWordIntegration.ps1
# One-shot installer for "Open Markdown in Word" right-click integration

# ==========================
# CONFIGURABLE PATHS
# ==========================
$PandocExe      = "C:\Program Files\Pandoc\pandoc-3.8.3\pandoc.exe"  # <-- adjust if needed
$ToolsRoot      = "C:\Tools\Pandoc"
$HelperScript   = Join-Path $ToolsRoot "OpenMdInWord.ps1"
$LogFile        = Join-Path $ToolsRoot "OpenMdInWord.log"

# ==========================
# STEP 1: Remove old context menu entries
# ==========================
Write-Host "`n=== STEP 1: Remove old context menu entries ==="

$oldKeys = @(
    "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\.md\shell\OpenInWord",
    "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\.md\shell\OpenMarkdownInWord"
)

foreach ($key in $oldKeys) {
    if (Test-Path $key) {
        Remove-Item -Path $key -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "Removed: $key"
    }
}

# ==========================
# STEP 2: Ensure tools folder exists
# ==========================
Write-Host "`n=== STEP 2: Ensure tools folder exists ==="

if (-not (Test-Path $ToolsRoot)) {
    New-Item -ItemType Directory -Path $ToolsRoot | Out-Null
    Write-Host "Created: $ToolsRoot"
} else {
    Write-Host "Exists: $ToolsRoot"
}

# ==========================
# STEP 3: Write helper script
# ==========================
Write-Host "`n=== STEP 3: Write helper script ==="

$helperScriptContent = @"
param([string]`$mdPath)

`$log = "$LogFile"

function Log(`$msg) {
    Add-Content -Path `$log -Value ("[{0}] {1}" -f (Get-Date), `$msg)
}

Log "---- New invocation ----"
Log "Raw input: `$mdPath"

try {
    `$mdPath = `$mdPath.Trim('"')
    Log "Normalized path: `$mdPath"

    if (-not (Test-Path `$mdPath)) {
        Log "ERROR: Markdown file not found"
        exit 1
    }

    `$docxPath = [System.IO.Path]::ChangeExtension(`$mdPath, ".docx")
    Log "DOCX path: `$docxPath"

    # Run Pandoc (absolute path for determinism)
    `$pandoc = "$PandocExe"
    `$args = @("`"`$mdPath`"", "-o", "`"`$docxPath`"")

    Log "Running Pandoc: `$pandoc `$mdPath -o `$docxPath"
    `$proc = Start-Process -FilePath `$pandoc -ArgumentList `$args -NoNewWindow -Wait -PassThru

    Log "Pandoc exit code: `$(`$proc.ExitCode)"

    if (`$proc.ExitCode -ne 0) {
        Log "ERROR: Pandoc conversion failed"
        exit 1
    }

    # Open Word
    Log "Launching Word: winword.exe `$docxPath"
    Start-Process "winword.exe" "`"`$docxPath`""

    Log "SUCCESS"
}
catch {
    Log "EXCEPTION: `$(`$_.Exception.Message)"
    exit 1
}
"@

Set-Content -Path $HelperScript -Value $helperScriptContent -Encoding UTF8
Write-Host "Helper script written to: $HelperScript"

# ==========================
# STEP 4: Create context menu entry
# ==========================
Write-Host "`n=== STEP 4: Create context menu entry ==="

$menuKey = "Registry::HKEY_CLASSES_ROOT\SystemFileAssociations\.md\shell\OpenMarkdownInWord"
$cmdKey  = "$menuKey\command"

New-Item -Path $menuKey -Force | Out-Null
Set-ItemProperty -Path $menuKey -Name "(default)" -Value "Open Markdown in Word"

New-Item -Path $cmdKey -Force | Out-Null
$cmd = "powershell.exe -ExecutionPolicy Bypass -File `"$HelperScript`" `"%1`""
Set-ItemProperty -Path $cmdKey -Name "(default)" -Value $cmd

Write-Host "Context menu installed: Open Markdown in Word"

# ==========================
# STEP 5: Validate Pandoc
# ==========================
Write-Host "`n=== STEP 5: Validate Pandoc path ==="

if (Test-Path $PandocExe) {
    Write-Host "Pandoc executable found at: $PandocExe"
} else {
    Write-Host "WARNING: Pandoc executable NOT found at:"
    Write-Host "  $PandocExe"
    Write-Host "Update `\$PandocExe in this script to the correct path."
}

# ==========================
# STEP 6: Summary
# ==========================
Write-Host "`n=== INSTALLATION COMPLETE ==="
Write-Host "Helper script: $HelperScript"
Write-Host "Log file:      $LogFile"
Write-Host "Context menu:  Right-click any .md → 'Open Markdown in Word'"
```

---

## 4. Verify the right‑click workflow

- **Action:** Right‑click any `.md` file  
  - Example: `C:\Repos\esp-fridge-temp-sensor\specs\esp-fridge-temp-sensor-fsd.md`
- **Action:** Choose **“Open Markdown in Word”**

You should see:

- **Result:** A `.docx` file created next to the `.md`  
  - e.g. `esp-fridge-temp-sensor-fsd.docx`
- **Result:** Word opens the `.docx`

If it fails:

- **Action:** Open the log file:

  ```powershell
  notepad C:\Tools\Pandoc\OpenMdInWord.log
  ```

- **Result:** Check for:
  - Missing Markdown file
  - Pandoc path issues
  - Word launch issues

---

## 5. Optional: Add Word macro for DOCX → Markdown

This section is **manual by design** and can be done later. It describes how to wire Word so that saving/exporting to Markdown calls a script (not provided here, but the pattern is clear).

### 5.1 Open the Word macro editor

- **Action:** Open Word.
- **Action:** Press `Alt+F11` to open the **VBA editor**.
- **Action:** In the left pane, locate **Normal** (your `Normal.dotm` template).

### 5.2 Create a new module

- **Action:** Right‑click **Normal** → **Insert** → **Module**.
- **Action:** A new module (e.g. `Module1`) appears.

### 5.3 Add a macro stub (example pattern)

> This is just a pattern; you can adapt it to call a DOCX → MD script later.

- **Action:** Paste something like:

  ```vba
  Sub ExportDocxToMarkdown()
      ' Example stub: call external script here
      ' Shell "powershell.exe -ExecutionPolicy Bypass -File ""C:\Tools\Pandoc\DocxToMd.ps1"" """ & ActiveDocument.FullName & """", vbNormalFocus
      MsgBox "DOCX → Markdown export not yet wired. Add your script call here.", vbInformation
  End Sub
  ```

- **Action:** Save and close the VBA editor.

### 5.4 Add a button or shortcut (optional)

- **Action:** In Word, right‑click the ribbon → **Customize the Ribbon**.
- **Action:** Add a custom group and assign the `ExportDocxToMarkdown` macro to a button.
- **Result:** You now have a manual entry point for DOCX → MD once you create the companion script.

---

## 6. Quick troubleshooting checklist

- **No context menu entry:**
  - Re‑run `Install-MarkdownWordIntegration.ps1` as Administrator.
- **Context menu appears but nothing happens:**
  - Check `C:\Tools\Pandoc\OpenMdInWord.log`.
- **Log shows “Markdown file not found”:**
  - Verify the file path in the log matches the actual file.
- **Log shows “Pandoc executable not found” or similar:**
  - Update `$PandocExe` in the installer script to the correct `pandoc.exe` path and re‑run.
- **DOCX created but Word doesn’t open:**
  - Check if `winword.exe` is on PATH; if not, you can hard‑code its full path in the helper script.

---

You now have a **copy‑pasteable, single‑script installation** plus a clear set of verification and troubleshooting steps for a deterministic Markdown → Word workflow.  
```
