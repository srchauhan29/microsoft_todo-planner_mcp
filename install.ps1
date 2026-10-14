<#
.SYNOPSIS
    Installs the Microsoft To Do & Planner MCP server for the current Windows user.

.DESCRIPTION
    1. Installs uv (per-user, no admin, no MSI) if it is missing.
    2. Installs Python + dependencies for this project (pinned by uv.lock).
    3. Signs you in to Microsoft 365 (option 1: own app registration, option 2: Microsoft-published app).
    4. Registers the server with Claude Desktop (and optionally Claude Code).
    5. Verifies the result.

    Run it from a normal PowerShell / Windows Terminal window, NOT from a terminal inside Claude:
    Claude Desktop is a packaged app and redirects files written from inside it.

    No secrets are written. Sign-in tokens are stored per user, encrypted with Windows DPAPI,
    in %USERPROFILE%\.todo-planner-mcp.

.EXAMPLE
    .\install.ps1
    Interactive install (asks which sign-in option to use).

.EXAMPLE
    .\install.ps1 -App microsoft -Tenant contoso.com.au
    Microsoft-published app, sign-in restricted to your own tenant.

.EXAMPLE
    .\install.ps1 -App own -ClientId 00000000-0000-0000-0000-000000000000 -Tenant contoso.com.au -ClaudeCode
    Own app registration; also register with Claude Code.

.EXAMPLE
    .\install.ps1 -Uninstall
    Removes the Claude registrations and cached sign-in (leaves this folder and uv in place).
#>
[CmdletBinding()]
param(
    [ValidateSet('own', 'microsoft')]
    [string]$App,
    [string]$ClientId,
    [string]$Tenant,
    [string]$TimeZone = 'auto',
    [switch]$SkipLogin,
    [switch]$SkipClaudeDesktop,
    [switch]$ClaudeCode,
    [switch]$Uninstall,
    [switch]$Yes
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProjectDir = $PSScriptRoot
$DataDir = Join-Path $env:USERPROFILE '.todo-planner-mcp'

function Write-Step([string]$Text) { Write-Host "`n==> $Text" -ForegroundColor Cyan }
function Write-Ok([string]$Text) { Write-Host "    OK  $Text" -ForegroundColor Green }
function Write-Note([string]$Text) { Write-Host "    $Text" -ForegroundColor Yellow }

function Confirm-Action([string]$Question) {
    if ($Yes) { return $true }
    $answer = Read-Host "$Question [Y/n]"
    return ($answer -eq '' -or $answer -match '^(y|yes)$')
}

function Invoke-Native {
    # Run a native command and stop on a non-zero exit code.
    # Windows PowerShell 5.1 turns a native tool's stderr (uv writes progress there) into error records
    # when output is redirected (logs, Intune); with ErrorActionPreference=Stop that would abort the
    # script. Judge success by exit code only.
    param([string]$Exe, [string[]]$Arguments, [string]$What)
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        & $Exe @Arguments 2>&1 | ForEach-Object { "$_" }
        $code = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $previous
    }
    if ($code -ne 0) { throw "$What failed (exit code $code)." }
}

function Invoke-Interactive {
    # For commands that prompt the user: keep the console attached (no piping) so prompts show.
    param([string]$Exe, [string[]]$Arguments, [string]$What)
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try { & $Exe @Arguments; $code = $LASTEXITCODE } finally { $ErrorActionPreference = $previous }
    if ($code -ne 0) { throw "$What failed (exit code $code)." }
}

function Invoke-Soft {
    # Like Invoke-Native, but a failure is only reported, not fatal.
    param([string]$Exe, [string[]]$Arguments)
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try { & $Exe @Arguments 2>&1 | ForEach-Object { "$_" } } finally { $ErrorActionPreference = $previous }
}

function Test-PackagedProcess {
    # True when this PowerShell's AppData writes are redirected into an MSIX package's private folder,
    # e.g. a terminal started from Claude Desktop. Child processes can be redirected without having a
    # package identity of their own, so test what actually happens to a write.
    $probe = Join-Path $env:APPDATA ("_todo_mcp_probe_{0}.tmp" -f [guid]::NewGuid().ToString('N'))
    try {
        Set-Content -LiteralPath $probe -Value 'probe'
        $target = @((Get-Item -LiteralPath $probe).Target) -join ''
        return $target -match '\\Packages\\[^\\]+\\LocalCache\\'
    } catch {
        return $false
    } finally {
        Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
    }
}

function Get-UvPath {
    $cmd = Get-Command uv -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    $candidate = Join-Path $env:USERPROFILE '.local\bin\uv.exe'
    if (Test-Path $candidate) { return $candidate }
    return $null
}

function Install-Uv {
    Write-Note 'uv is not installed. It installs per-user into %USERPROFILE%\.local\bin (no admin rights, no MSI).'
    if (-not (Confirm-Action 'Install uv now?')) {
        throw 'uv is required. Install it from https://docs.astral.sh/uv/ and re-run this script.'
    }
    $uv = $null
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        # Preferred: winget verifies the download's SHA-256 against the community-reviewed manifest.
        Write-Note 'Installing uv with winget (hash-verified, per-user)...'
        Invoke-Soft 'winget' @('install', '--id', 'astral-sh.uv', '--exact', '--source', 'winget', '--scope', 'user',
            '--silent', '--accept-package-agreements', '--accept-source-agreements') | Out-Host
        $env:Path = [Environment]::GetEnvironmentVariable('Path', 'User') + ';' + $env:Path
        $uv = Get-UvPath
    }
    if (-not $uv) {
        # Fallback: the official installer from astral.sh over HTTPS (no independent hash check).
        Write-Note 'Installing uv with the official installer from https://astral.sh/uv ...'
        $installer = Invoke-RestMethod -Uri 'https://astral.sh/uv/install.ps1' -UseBasicParsing
        Invoke-Expression $installer
        $uv = Get-UvPath
    }
    if (-not $uv) { throw 'uv installation did not produce uv.exe. Open a new terminal and re-run.' }
    $env:Path = (Split-Path $uv) + ';' + $env:Path
    return $uv
}

function Test-ClaudeDesktopRunning {
    # Claude Desktop: Store build under WindowsApps\Claude_*, classic build under AnthropicClaude.
    # Ignore the Claude Code CLI, which is also called claude.exe.
    $procs = Get-Process -Name 'claude' -ErrorAction SilentlyContinue | Where-Object {
        -not $_.Path -or $_.Path -match '\\WindowsApps\\Claude_|\\AnthropicClaude\\'
    }
    return [bool]$procs
}

function Wait-ClaudeClosed {
    while (Test-ClaudeDesktopRunning) {
        Write-Note 'Claude Desktop is running. Quit it from the tray icon (right-click -> Quit) so its config can be updated.'
        if ($Yes) { throw 'Claude Desktop is running; quit it and re-run (or use -SkipClaudeDesktop).' }
        $answer = Read-Host 'Press Enter once Claude Desktop is closed (or type S to skip Claude Desktop registration)'
        if ($answer -match '^(s|skip)$') { return $false }
    }
    return $true
}

# ------------------------------------------------------------------ preflight
Write-Host 'Microsoft To Do & Planner MCP - installer' -ForegroundColor White
Write-Host "Project folder: $ProjectDir"

if (-not (Test-Path (Join-Path $ProjectDir 'pyproject.toml'))) {
    throw "pyproject.toml not found next to this script. Run install.ps1 from the project folder."
}
if (Test-PackagedProcess) {
    throw ('This PowerShell is running inside a packaged app (probably Claude Desktop). ' +
           'Open a normal PowerShell or Windows Terminal window and run the script there.')
}

Write-Step 'Checking uv'
$uv = Get-UvPath
if (-not $uv) { $uv = Install-Uv }
Write-Ok ((& $uv --version) -join ' ')

# ------------------------------------------------------------------ uninstall
if ($Uninstall) {
    Write-Step 'Removing Claude Desktop registration'
    if (Wait-ClaudeClosed) {
        Invoke-Native $uv @('run', '--directory', $ProjectDir, 'python', 'scripts\register_claude_desktop.py', '--remove') 'Claude Desktop unregistration'
    }
    if (Get-Command claude -ErrorAction SilentlyContinue) {
        Write-Step 'Removing Claude Code registration (if present)'
        Invoke-Soft 'claude' @('mcp', 'remove', 'todo-planner', '--scope', 'user') | Out-Null
        Write-Ok 'Done'
    }
    Write-Step 'Signing out'
    Invoke-Soft $uv @('run', '--directory', $ProjectDir, 'todo-planner-mcp', 'logout')
    if ((Test-Path $DataDir) -and (Confirm-Action "Delete saved settings and token cache in $DataDir?")) {
        Remove-Item -LiteralPath $DataDir -Recurse -Force
        Write-Ok "Deleted $DataDir"
    }
    Write-Host "`nUninstalled. You can delete $ProjectDir if you no longer need it." -ForegroundColor Green
    return
}

# ------------------------------------------------------------------ install
Write-Step 'Installing Python and dependencies (pinned by uv.lock)'
try {
    Invoke-Native $uv @('sync', '--frozen', '--directory', $ProjectDir) 'Dependency install'
} catch {
    # A half-built .venv (e.g. from an interrupted run) blocks uv; it is fully derived, so rebuild it once.
    $venv = Join-Path $ProjectDir '.venv'
    if (-not (Test-Path $venv)) { throw }
    Write-Note 'Rebuilding the project environment (.venv) and retrying...'
    Remove-Item -LiteralPath $venv -Recurse -Force
    Invoke-Native $uv @('sync', '--frozen', '--directory', $ProjectDir) 'Dependency install'
}
Write-Ok 'Dependencies installed'

if (-not $SkipLogin) {
    Write-Step 'Signing in to Microsoft 365 (your browser will open)'
    $loginArgs = @('run', '--directory', $ProjectDir, 'todo-planner-mcp', 'login')
    if ($App) {
        $loginArgs += @('--app', $App, '--timezone', $TimeZone)
        if ($ClientId) { $loginArgs += @('--client-id', $ClientId) }
        if ($Tenant) { $loginArgs += @('--tenant', $Tenant) }
    } else {
        Write-Note "Choose 1 (your organisation's own app registration) or 2 (Microsoft-published app)."
    }
    Invoke-Interactive $uv $loginArgs 'Sign-in'
}

if (-not $SkipClaudeDesktop) {
    Write-Step 'Registering with Claude Desktop'
    if (Wait-ClaudeClosed) {
        Invoke-Native $uv @('run', '--directory', $ProjectDir, 'python', 'scripts\register_claude_desktop.py') 'Claude Desktop registration'
    } else {
        Write-Note 'Skipped. Re-run later with Claude Desktop closed, or install the .mcpb extension instead.'
    }
}

if ($ClaudeCode) {
    Write-Step 'Registering with Claude Code'
    if (Get-Command claude -ErrorAction SilentlyContinue) {
        Invoke-Soft 'claude' @('mcp', 'remove', 'todo-planner', '--scope', 'user') | Out-Null
        Invoke-Native 'claude' @('mcp', 'add', 'todo-planner', '--scope', 'user', '--', $uv, 'run', '--frozen', '--directory', $ProjectDir, 'todo-planner-mcp') 'Claude Code registration'
        Write-Ok 'Registered (user scope)'
    } else {
        Write-Note 'Claude Code CLI (claude) not found on PATH - skipped.'
    }
}

Write-Step 'Verifying'
Invoke-Soft $uv @('run', '--directory', $ProjectDir, 'todo-planner-mcp', 'status')

Write-Host @"

Done. Next:
  1. Start Claude Desktop.
  2. In any chat, Project or Cowork task, ask: "Check my To Do sign-in status"
  3. Then try: "Show my To Do lists" or "Show my Planner tasks"
If the Microsoft 365 session ever expires, a browser sign-in window opens automatically.
"@ -ForegroundColor Green
