<#
.SYNOPSIS
    Orchestrates MTG Art Downloader, Proxyshop, print-sheet assembly, review,
    and printing for the card-proxy workflow.

.DESCRIPTION
    Portable defaults are resolved relative to the folder containing this
    script:
      - MTG-Art-Downloader: .\MTG-Art-Downloader
      - Proxyshop:          .\Proxyshop
      - Sheet assembler:    .\assemble_card_sheets.py
      - Card-back PDF:      .\horizontal_cardBack_3x3.pdf

    The card-back PDF is not required when -SingleSided is used. Use the path
    parameters when those projects live elsewhere. Proxyshop
    renders are read from ProxyshopDir\out, and each run writes print sheets to
    ProxyshopDir\orchestrated_runs\<timestamp>.

    Important installation note:
    MTG-Art-Downloader is the Poetry-based application in this workflow and is
    launched with `poetry run python main.py`. The normal Proxyshop 1.13.2
    GitHub release is Proxyshop.exe and does not use Poetry. Its release CLI
    does not expose an equivalent to the GUI's "Render All" command.

    Proxyshop modes:
      1. GuiManual - Default for the normal GitHub Proxyshop.exe release. The
         script prepares art, opens Proxyshop, and asks you to click Render All.
      2. SourceHeadless - Optional advanced mode. It requires a separate
         Proxyshop source checkout with its Python dependencies installed. The
         script applies two minimal Python 3.10 compatibility fixes, then
         generates and runs a temporary headless render-all runner.

    Printing note:
    Windows does not provide a supported command-line API for selecting a named
    printer preset. ManualDialog mode is the default because it lets you choose
    the correct settings in the printer's system dialog. ExperimentalPreset
    mode attempts to select an optional named preset through UI Automation,
    then prints silently through SumatraPDF.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\orchestrate_proxy_workflow.ps1

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\orchestrate_proxy_workflow.ps1 `
      -DownloaderDir "C:\path\to\MTG-Art-Downloader" `
      -ProxyshopDir "C:\path\to\Proxyshop" `
      -ProxyshopMode SourceHeadless

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\orchestrate_proxy_workflow.ps1 `
      -SingleSided
#>

#requires -Version 5.1

[CmdletBinding()]
param(
    # Blank paths resolve to portable defaults relative to this script:
    # .\MTG-Art-Downloader, .\Proxyshop, .\assemble_card_sheets.py, and
    # .\horizontal_cardBack_3x3.pdf.
    [string]$DownloaderDir = "",
    [string]$ProxyshopDir = "",
    [string]$AssemblerScript = "",
    [string]$BackPdf = "",

    [ValidateSet("Auto", "SourceHeadless", "GuiManual")]
    [string]$ProxyshopMode = "Auto",

    [string]$PoetryExe = "poetry",
    [string]$PythonExe = "py",

    # Ensures MTG-Art-Downloader's Poetry environment has the dependencies from
    # pyproject.toml before main.py runs. Disable only if you manage that env manually.
    [bool]$InstallDownloaderDependencies = $true,

    # SourceHeadless mode requires a Poetry environment containing Proxyshop's
    # dependencies. The normal Proxyshop.exe release ignores this setting.
    [bool]$InstallProxyshopDependencies = $true,

    # Cleans only top-level renderable images from Proxyshop's art folder before
    # copying this run's card art. Files beginning with ! and non-image files remain.
    [bool]$CleanProxyshopArt = $true,

    # Deletes old images from MTG Art Downloader's configured download folder
    # before downloading. Disabled by default to avoid surprising data removal.
    [bool]$CleanDownloaderOutput = $false,

    [switch]$SkipDownload,
    [switch]$SkipProxyshop,
    [switch]$SkipAssembly,
    [switch]$SkipPrint,

    # Generates front-only sheets and does not require the card-back PDF.
    [switch]$SingleSided,

    [string]$PrinterName = "",

    # Optional driver preset name. Leave blank to use the printer's defaults.
    [string]$PrintPreset = "",

    [ValidateSet("ManualDialog", "AutomaticDefault", "ExperimentalPreset", "Off")]
    [string]$PrintMode = "ManualDialog",

    [ValidateSet("long-edge", "short-edge", "none")]
    [string]$DuplexBinding = "long-edge",

    [string]$SumatraPdfExe = ""
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

# MTG-Art-Downloader prints a Unicode block-art banner. Windows Python otherwise
# defaults piped stdout to cp1252 and raises UnicodeEncodeError. These process
# settings are inherited by Poetry and every Python child process it launches.
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"
try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
}
catch {
    Write-Host "WARNING: Could not switch the console to UTF-8; Python output will still use UTF-8." -ForegroundColor Yellow
}

$script:ImageExtensions = @(".jpg", ".jpeg", ".jpf", ".png", ".tif", ".tiff", ".webp")
$script:RunDir = $null
$script:RunLogDir = $null
$script:WorkflowStarted = Get-Date
$script:ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }


function Write-Section {
    param([Parameter(Mandatory = $true)][string]$Message)

    Write-Host ""
    Write-Host ("=" * 78) -ForegroundColor Cyan
    Write-Host " $Message" -ForegroundColor Cyan
    Write-Host ("=" * 78) -ForegroundColor Cyan
}


function Write-WarningLine {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Host "WARNING: $Message" -ForegroundColor Yellow
}


function Write-ErrorLine {
    param([Parameter(Mandatory = $true)][string]$Message)
    Write-Host "ERROR: $Message" -ForegroundColor Red
}


function Resolve-RequiredDirectory {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Description
    )

    $resolved = [System.IO.Path]::GetFullPath($Path)
    if (-not (Test-Path -LiteralPath $resolved -PathType Container)) {
        throw "$Description was not found: $resolved"
    }
    return $resolved
}


function Resolve-RequiredFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Description
    )

    $resolved = [System.IO.Path]::GetFullPath($Path)
    if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
        throw "$Description was not found: $resolved"
    }
    return $resolved
}


function Resolve-CommandPath {
    param(
        [Parameter(Mandatory = $true)][string]$Command,
        [string[]]$FallbackPaths = @()
    )

    if ($Command -match "[\\/]") {
        $expanded = [Environment]::ExpandEnvironmentVariables($Command)
        if (Test-Path -LiteralPath $expanded -PathType Leaf) {
            return (Resolve-Path -LiteralPath $expanded).Path
        }
    }

    $found = Get-Command $Command -ErrorAction SilentlyContinue
    if ($found) {
        return $found.Source
    }

    foreach ($candidate in $FallbackPaths) {
        $expanded = [Environment]::ExpandEnvironmentVariables($candidate)
        if (Test-Path -LiteralPath $expanded -PathType Leaf) {
            return (Resolve-Path -LiteralPath $expanded).Path
        }
    }

    throw "Could not find executable or command: $Command"
}


function Get-IniValue {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Section,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Default
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $Default
    }

    $currentSection = ""
    foreach ($line in Get-Content -LiteralPath $Path -Encoding UTF8) {
        $trimmed = $line.Trim()
        if (-not $trimmed -or $trimmed.StartsWith("#") -or $trimmed.StartsWith(";")) {
            continue
        }
        if ($trimmed -match "^\[(.+)\]$") {
            $currentSection = $Matches[1].Trim()
            continue
        }
        if ($currentSection -ieq $Section -and $trimmed -match "^([^=]+)=(.*)$") {
            $key = $Matches[1].Trim()
            if ($key -ieq $Name) {
                return $Matches[2].Trim()
            }
        }
    }
    return $Default
}


function Initialize-RunFolders {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $script:RunDir = Join-Path $ProxyshopDir ("orchestrated_runs\" + $timestamp)
    $script:RunLogDir = Join-Path $script:RunDir "logs"
    New-Item -ItemType Directory -Path $script:RunLogDir -Force | Out-Null
    Write-Host "Run folder: $script:RunDir"
}


function Show-LogTail {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [int]$LineCount = 80
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return
    }

    Write-Host ""
    Write-Host "--- Last $LineCount line(s): $Path ---" -ForegroundColor DarkYellow
    Get-Content -LiteralPath $Path -Tail $LineCount -ErrorAction SilentlyContinue |
        ForEach-Object { Write-Host $_ }
    Write-Host "--- End log excerpt ---" -ForegroundColor DarkYellow
}


function Show-FileContent {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return
    }

    Write-Host ""
    Write-Host "--- Contents: $Path ---" -ForegroundColor DarkYellow
    Get-Content -LiteralPath $Path -ErrorAction SilentlyContinue |
        ForEach-Object { Write-Host $_ }
    Write-Host "--- End file ---" -ForegroundColor DarkYellow
}


function Read-ContinueOrQuit {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$Message,
        [string[]]$LogPaths = @()
    )

    Write-Host ""
    Write-Host $Title -ForegroundColor Yellow
    Write-Host $Message
    foreach ($logPath in $LogPaths) {
        if ($logPath -and (Test-Path -LiteralPath $logPath -PathType Leaf)) {
            if ((Get-Item -LiteralPath $logPath).Name -ieq "failed.txt") {
                Show-FileContent -Path $logPath
            }
            else {
                Show-LogTail -Path $logPath -LineCount 80
            }
        }
    }

    while ($true) {
        $response = Read-Host "Type c to continue, or q to quit"
        switch -Regex ($response.Trim()) {
            "^(?i)c$" { return $true }
            "^(?i)q$" { return $false }
            default { Write-Host "Please enter only c or q." }
        }
    }
}


function Stop-AtGate {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$Message,
        [string[]]$LogPaths = @()
    )

    $continue = Read-ContinueOrQuit -Title $Title -Message $Message -LogPaths $LogPaths
    if (-not $continue) {
        Write-Host "Workflow stopped by user." -ForegroundColor Yellow
        return $false
    }
    return $true
}


function Invoke-NativeLogged {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [string[]]$Arguments = @(),
        [Parameter(Mandatory = $true)][string]$WorkingDirectory,
        [Parameter(Mandatory = $true)][string]$LogPath,
        [string[]]$StandardInputLines = @()
    )

    $logParent = Split-Path -Path $LogPath -Parent
    New-Item -ItemType Directory -Path $logParent -Force | Out-Null

    Write-Host "Running: $FilePath $($Arguments -join ' ')"
    Write-Host "Working directory: $WorkingDirectory"
    Write-Host "Log: $LogPath"

    Push-Location -LiteralPath $WorkingDirectory
    $previousErrorActionPreference = $ErrorActionPreference
    try {
        # PowerShell 5.1 converts native stderr lines into ErrorRecord objects.
        # With Stop as the global preference, even a non-fatal warning such as
        # Poetry's deprecation notice would otherwise terminate the workflow.
        # Continue lets the native process finish; its exit code remains authoritative.
        $ErrorActionPreference = "Continue"

        function Write-NativeLogLine {
            param([AllowNull()][object]$Value)

            $line = if ($Value -is [System.Management.Automation.ErrorRecord]) {
                $Value.Exception.Message
            }
            else {
                [string]$Value
            }
            Write-Host $line
            Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
        }

        if ($StandardInputLines.Count -gt 0) {
            $StandardInputLines | & $FilePath @Arguments 2>&1 |
                ForEach-Object { Write-NativeLogLine $_ }
        }
        else {
            & $FilePath @Arguments 2>&1 |
                ForEach-Object { Write-NativeLogLine $_ }
        }
        return $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previousErrorActionPreference
        Pop-Location
    }
}


function Copy-ToRunLogArchive {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Prefix
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return
    }

    $leaf = Split-Path -Path $Path -Leaf
    $destination = Join-Path $script:RunLogDir ("{0}_previous_{1}" -f $Prefix, $leaf)
    Copy-Item -LiteralPath $Path -Destination $destination -Force
    Remove-Item -LiteralPath $Path -Force
    Write-Host "Archived previous $(Split-Path -Path $Path -Leaf) to: $destination"
}


function Normalize-CardName {
    param([Parameter(Mandatory = $true)][string]$Value)

    $normalized = $Value.Normalize([Text.NormalizationForm]::FormKC)
    $normalized = $normalized.Replace("_", " ").ToLowerInvariant()
    return ([regex]::Replace($normalized, "\s+", " ")).Trim()
}


function Read-CardRequests {
    param([Parameter(Mandatory = $true)][string]$CardsTxt)

    $requests = New-Object System.Collections.Generic.List[object]
    $pattern = '^\s*(?:(?<count>\d+)\s+)?(?<name>.*?)(?:\s+\([^()]*\)\s+\S+)?\s*$'
    $lineNumber = 0

    foreach ($line in Get-Content -LiteralPath $CardsTxt -Encoding UTF8) {
        $lineNumber++
        $trimmed = $line.Trim()
        if (-not $trimmed -or $trimmed.StartsWith("#")) {
            continue
        }
        $match = [regex]::Match($line, $pattern)
        if (-not $match.Success) {
            Write-WarningLine "Could not parse cards.txt line ${lineNumber}: $trimmed"
            continue
        }

        $name = $match.Groups["name"].Value.Trim()
        $quantity = 1
        if ($match.Groups["count"].Success -and $match.Groups["count"].Value) {
            $quantity = [int]$match.Groups["count"].Value
        }
        if (-not $name -or $quantity -lt 1) {
            Write-WarningLine "Skipping invalid cards.txt line ${lineNumber}: $trimmed"
            continue
        }

        $requests.Add([pscustomobject]@{
            Name = $name
            NormalizedName = (Normalize-CardName -Value $name)
            Quantity = $quantity
            LineNumber = $lineNumber
        })
    }

    return $requests
}


function Test-ImageMatchesCard {
    param(
        [Parameter(Mandatory = $true)][string]$ImageStem,
        [Parameter(Mandatory = $true)][string]$NormalizedCardName
    )

    $stem = Normalize-CardName -Value $ImageStem
    if ($stem -eq $NormalizedCardName) {
        return $true
    }
    if (-not $stem.StartsWith($NormalizedCardName)) {
        return $false
    }
    if ($stem.Length -le $NormalizedCardName.Length) {
        return $false
    }

    $next = $stem.Substring($NormalizedCardName.Length, 1)
    return @(" ", "(", "[", "{", "-") -contains $next
}


function Get-DownloaderDownloadFolder {
    $configPath = Join-Path $DownloaderDir "config.ini"
    $configured = Get-IniValue -Path $configPath -Section "FILES" -Name "Download.Folder" -Default "downloaded"
    if ([System.IO.Path]::IsPathRooted($configured)) {
        return [System.IO.Path]::GetFullPath($configured)
    }
    return [System.IO.Path]::GetFullPath((Join-Path $DownloaderDir $configured))
}


function Clear-ImageFiles {
    param(
        [Parameter(Mandatory = $true)][string]$Folder,
        [bool]$Recursive,
        [string]$Description
    )

    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) {
        return
    }

    $items = if ($Recursive) {
        Get-ChildItem -LiteralPath $Folder -File -Recurse
    }
    else {
        Get-ChildItem -LiteralPath $Folder -File
    }

    $images = @($items | Where-Object {
        ($script:ImageExtensions -contains $_.Extension.ToLowerInvariant()) -and
        (-not $_.Name.StartsWith("!"))
    })

    foreach ($image in $images) {
        Remove-Item -LiteralPath $image.FullName -Force
    }
    Write-Host "Removed $($images.Count) old image file(s) from $Description."
}


function Ensure-DownloaderDependencies {
    param([Parameter(Mandatory = $true)][string]$Poetry)

    if (-not $InstallDownloaderDependencies) {
        return $true
    }

    $preflightLog = Join-Path $script:RunLogDir "00_downloader_dependency_check.log"
    $preflightCode = Invoke-NativeLogged `
        -FilePath $Poetry `
        -Arguments @("run", "python", "-c", "import pathvalidate") `
        -WorkingDirectory $DownloaderDir `
        -LogPath $preflightLog

    if ($preflightCode -eq 0) {
        Write-Host "MTG Art Downloader dependency check passed." -ForegroundColor Green
        return $true
    }

    Write-WarningLine "MTG Art Downloader dependencies are missing; running poetry install --no-root."
    $installLog = Join-Path $script:RunLogDir "00_downloader_poetry_install.log"
    $installCode = Invoke-NativeLogged `
        -FilePath $Poetry `
        -Arguments @("install", "--no-root") `
        -WorkingDirectory $DownloaderDir `
        -LogPath $installLog

    if ($installCode -ne 0) {
        return Stop-AtGate `
            -Title "poetry install failed for MTG Art Downloader." `
            -Message "Review the dependency installation output before deciding whether to continue." `
            -LogPaths @($installLog, $preflightLog)
    }

    $retryLog = Join-Path $script:RunLogDir "00_downloader_dependency_recheck.log"
    $retryCode = Invoke-NativeLogged `
        -FilePath $Poetry `
        -Arguments @("run", "python", "-c", "import pathvalidate") `
        -WorkingDirectory $DownloaderDir `
        -LogPath $retryLog

    if ($retryCode -ne 0) {
        return Stop-AtGate `
            -Title "MTG Art Downloader dependencies are still unavailable." `
            -Message "poetry install completed, but the downloader environment still cannot import its dependencies." `
            -LogPaths @($retryLog, $installLog)
    }

    Write-Host "MTG Art Downloader dependencies installed." -ForegroundColor Green
    return $true
}


function Invoke-DownloaderStage {
    if ($SkipDownload) {
        Write-Section "Stage 1/4: MTG Art Downloader (skipped)"
        return $true
    }

    Write-Section "Stage 1/4: MTG Art Downloader"

    $cardsTxt = Resolve-RequiredFile -Path (Join-Path $DownloaderDir "cards.txt") -Description "cards.txt"
    $failedTxt = Join-Path $DownloaderDir "failed.txt"
    Copy-ToRunLogArchive -Path $failedTxt -Prefix "downloader"

    if ($CleanDownloaderOutput) {
        $downloadFolder = Get-DownloaderDownloadFolder
        Clear-ImageFiles -Folder $downloadFolder -Recursive $true -Description "MTG Art Downloader output"
    }

    $poetry = Resolve-CommandPath -Command $PoetryExe
    if (-not (Ensure-DownloaderDependencies -Poetry $poetry)) {
        return $false
    }

    $logPath = Join-Path $script:RunLogDir "01_mtg_art_downloader.log"

    # MTG Art Downloader asks for Enter once at startup and once at completion.
    $exitCode = Invoke-NativeLogged `
        -FilePath $poetry `
        -Arguments @("run", "python", "main.py") `
        -WorkingDirectory $DownloaderDir `
        -LogPath $logPath `
        -StandardInputLines @("", "")

    if ($exitCode -ne 0) {
        return Stop-AtGate `
            -Title "MTG Art Downloader exited with code $exitCode." `
            -Message "Review the downloader output before deciding whether to continue." `
            -LogPaths @($logPath)
    }

    if (Test-Path -LiteralPath $failedTxt -PathType Leaf) {
        $failedLength = (Get-Item -LiteralPath $failedTxt).Length
        if ($failedLength -gt 0) {
            return Stop-AtGate `
                -Title "MTG Art Downloader created failed.txt." `
                -Message "Some requested images could not be downloaded. Review the failures, then continue or quit." `
                -LogPaths @($failedTxt, $logPath)
        }
    }

    Write-Host "MTG Art Downloader completed." -ForegroundColor Green
    return $true
}


function Copy-DownloadedArtStage {
    Write-Section "Stage 2/4: Sync downloaded art to Proxyshop"

    $cardsTxt = Resolve-RequiredFile -Path (Join-Path $DownloaderDir "cards.txt") -Description "cards.txt"
    $downloadFolder = Get-DownloaderDownloadFolder
    $artFolder = Join-Path $ProxyshopDir "art"
    New-Item -ItemType Directory -Path $artFolder -Force | Out-Null

    if (-not (Test-Path -LiteralPath $downloadFolder -PathType Container)) {
        Write-WarningLine "Downloader output folder does not exist: $downloadFolder"
        $downloadFolder = $DownloaderDir
    }

    if ($CleanProxyshopArt) {
        Clear-ImageFiles -Folder $artFolder -Recursive $false -Description "Proxyshop art folder"
    }

    $requests = Read-CardRequests -CardsTxt $cardsTxt
    if ($requests.Count -eq 0) {
        throw "No valid card entries were found in $cardsTxt"
    }

    $allImages = @(
        Get-ChildItem -LiteralPath $downloadFolder -File -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $script:ImageExtensions -contains $_.Extension.ToLowerInvariant() }
    )

    $selected = New-Object System.Collections.Generic.List[object]
    $missing = New-Object System.Collections.Generic.List[object]

    foreach ($request in $requests) {
        $matches = @(
            $allImages | Where-Object {
                Test-ImageMatchesCard -ImageStem $_.BaseName -NormalizedCardName $request.NormalizedName
            } | Sort-Object LastWriteTimeUtc -Descending
        )

        if ($matches.Count -eq 0) {
            $missing.Add($request)
            continue
        }

        # If the same output name exists in both mtgpics and scryfall, keep the newest.
        $uniqueMatches = @($matches | Group-Object Name | ForEach-Object { $_.Group | Select-Object -First 1 })
        foreach ($image in $uniqueMatches) {
            $selected.Add($image)
        }
    }

    foreach ($request in $missing) {
        Write-WarningLine "No downloaded art matched '$($request.Name)' from cards.txt line $($request.LineNumber)."
    }

    foreach ($image in ($selected | Sort-Object FullName -Unique)) {
        Copy-Item -LiteralPath $image.FullName -Destination (Join-Path $artFolder $image.Name) -Force
    }

    $artImages = @(
        Get-ChildItem -LiteralPath $artFolder -File |
            Where-Object {
                ($script:ImageExtensions -contains $_.Extension.ToLowerInvariant()) -and
                (-not $_.Name.StartsWith("!"))
            }
    )

    Write-Host "Copied $($selected.Count) matching art image(s) into: $artFolder"
    Write-Host "Proxyshop will see $($artImages.Count) renderable art image(s)."

    if ($artImages.Count -eq 0) {
        return Stop-AtGate `
            -Title "No Proxyshop art is available." `
            -Message "The downloader output could not be matched to cards.txt. Review the paths and downloader output." `
            -LogPaths @()
    }

    if ($missing.Count -gt 0) {
        return Stop-AtGate `
            -Title "$($missing.Count) cards.txt line(s) have no downloaded art." `
            -Message "Proxyshop cannot render missing art. Continue only if you intentionally want to skip those cards." `
            -LogPaths @((Join-Path $DownloaderDir "failed.txt"))
    }

    return $true
}


function Get-HeadlessProxyshopRunnerCode {
    return @'
"""
Temporary headless Render All runner generated by orchestrate_proxy_workflow.ps1.
This mirrors the render-all pipeline in Proxyshop 1.13.2 without importing Kivy.
"""

from concurrent.futures import ThreadPoolExecutor
from multiprocessing import cpu_count
import os
from pathlib import Path
from threading import Event
from time import perf_counter
import typing
from typing import Optional

# Select Proxyshop's TerminalConsole before src initializes its globals. Without
# this, src creates GUIConsole and render status updates attempt to access Kivy
# widget IDs that do not exist in this non-GUI process.
os.environ["PROXYSHOP_HEADLESS"] = "1"

# Proxyshop 1.13.2 imports NotRequired from typing, which is unavailable before
# Python 3.11. Its declared Python range starts at 3.10, so provide the
# equivalent typing_extensions symbol before importing any Proxyshop modules.
from typing_extensions import NotRequired

if not hasattr(typing, "NotRequired"):
    typing.NotRequired = NotRequired

from photoshop.api import PurgeTarget, SaveOptions

from src import APP, CFG, CON, CONSOLE, ENV, TEMPLATE_DEFAULTS
from src._loader import get_template_map_selected
from src.console import get_bullet_points, msg_bold, msg_error, msg_success, msg_warn
from src.layouts import assign_layout, join_dual_card_layouts
from src.utils.adobe import get_photoshop_error_message


class HeadlessProxyshop:
    """Small non-GUI equivalent of the parts of ProxyshopGUIApp used by Render All."""

    def __init__(self):
        self.app = APP
        self.cfg = CFG
        self.con = CON
        self.env = ENV
        self.console = CONSOLE
        self.templates_default = TEMPLATE_DEFAULTS
        self.templates_selected = {}
        self.current_render = None

    @property
    def timer(self) -> float:
        return perf_counter()

    @property
    def docref(self):
        if self.current_render and hasattr(self.current_render, "docref"):
            return self.current_render.docref or None
        return None

    @property
    def thread(self) -> Optional[Event]:
        if self.current_render and hasattr(self.current_render, "event"):
            return self.current_render.event or None
        return None

    @property
    def thread_cancelled(self) -> bool:
        thread = self.thread
        return bool(not isinstance(thread, Event) or thread.is_set())

    def get_art_files(self):
        from src import PATH

        art_folder = PATH.ART
        extensions = (".png", ".jpg", ".tif", ".jpeg", ".jpf")
        files = [
            path for path in art_folder.iterdir()
            if path.is_file()
            and path.suffix.lower() in extensions
            and not path.name.startswith("!")
        ]

        webp_files = [
            path for path in art_folder.iterdir()
            if path.is_file()
            and path.suffix.lower() == ".webp"
            and not path.name.startswith("!")
        ]
        if webp_files and not self.app.supports_webp:
            self.console.update(msg_warn("Skipped WEBP image, WEBP requires Photoshop ^23.2.0"))
        elif webp_files:
            files.extend(webp_files)
        return sorted(files)

    def close_document(self) -> None:
        if self.docref:
            try:
                self.docref.close(SaveOptions.DoNotSaveChanges)
                self.app.purge(PurgeTarget.AllCaches)
            except Exception as error:
                print("Couldn't close corresponding document!")
                self.console.log_exception(error)
        self.current_render = None

    def reset(self, close_document: bool = False) -> None:
        self.cfg.load()
        self.con.reload()
        if close_document:
            self.close_document()

    def start_render(self, card, template, loaded_class, reload_config=False, reload_constants=False):
        if not self.env.TEST_MODE:
            self.console.update(msg_success(f"---- {card.display_name} ----"))

        if reload_config:
            self.cfg.load(config=template["config"])
        if reload_constants:
            self.con.reload()

        try:
            card.template_file = template["object"].path_psd
            self.current_render = loaded_class(card)

            with ThreadPoolExecutor() as executor:
                executor.submit(self.console.start_await_cancel, self.current_render.event)

            start_time = self.timer
            result = self.current_render.execute()
            elapsed = round(self.timer - start_time, 1)

            if not self.thread.is_set() and result:
                if not self.env.TEST_MODE:
                    self.console.update(f"[i]Time completed: {elapsed} seconds[/i]\n")
                return elapsed

        except Exception as error:
            if self.docref:
                self.current_render.reset()
            self.console.log_error(
                self.thread or Event(),
                card=card.name,
                template=template["name"],
                msg=msg_error(
                    "Encountered a general error!\n"
                    "Check [b]/logs/error.txt[/b] for details."
                ),
                e=error,
            )
        return None

    def render_all(self) -> int:
        """Render every supported image in Proxyshop's art folder."""
        self.reset()

        while check := self.app.refresh_app():
            if not self.console.await_choice(
                Event(),
                get_photoshop_error_message(check),
                end="Hit Continue to try again, or Cancel to end the operation.\n",
            ):
                return 1

        templates = get_template_map_selected(self.templates_selected, self.templates_default)
        files = self.get_art_files()
        if not files:
            self.console.update(msg_error("No art images found!"))
            return 2

        with ThreadPoolExecutor(max_workers=cpu_count()) as pool:
            cards = pool.map(assign_layout, files)
        cards = join_dual_card_layouts(list(cards))

        layouts = {}
        failed = []
        for card in cards:
            if isinstance(card, str):
                failed.append(card)
                continue

            if not templates[card.card_class]["object"].is_installed:
                card = msg_error(
                    msg=card.display_name,
                    reason=(
                        f"Template '{templates[card.card_class]['name']}' with type "
                        f"'{card.card_class}' is not installed!"
                    ),
                )
                failed.append(card)
                continue

            layouts.setdefault(
                str(templates[card.card_class]["object"].path_psd), {}
            ).setdefault(card.card_class, []).append(card)

        if failed:
            failure_list = "\n".join(failed)
            if not layouts:
                self.console.update(
                    f"\n{msg_bold(msg_error('Failed to render all cards!'))}\n{failure_list}"
                )
                return 3
            if not self.console.error(
                msg=f"\n{msg_error('Unable to render these cards:')}{failure_list}"
            ):
                return 4

        self.console.update()
        times = []
        for index, (_path, class_map) in enumerate(layouts.items()):
            for layout_type, layout_cards in class_map.items():
                loaded_class = templates[layout_type]["object"].get_template_class(
                    templates[layout_type]["class_name"]
                )
                if not loaded_class:
                    self.console.update(
                        msg_error(
                            "Unable to load Python class: "
                            f"{msg_bold(templates[layout_type]['class_name'])}"
                        )
                    )
                    if index + 1 < len(layouts):
                        failed_cards = get_bullet_points(
                            text=[str(card.display_name) for card in layout_cards], char="-"
                        )
                        if self.console.error(
                            msg=(
                                f"{msg_error('The following cards have been cancelled:')}"
                                f"{failed_cards}"
                            )
                        ):
                            self.console.update()
                            continue
                    return 5

                self.cfg.load(templates[layout_type]["config"])
                self.con.reload()
                for card in layout_cards:
                    result = self.start_render(card, templates[layout_type], loaded_class)
                    if self.thread_cancelled:
                        return 6
                    if result is not None:
                        times.append(result)

            self.close_document()

        self.console.update(msg_success("Renders Completed!"))
        if times:
            average = round(sum(times) / len(times), 1)
            self.console.update(f"Average time: {average} seconds")
        return 0


if __name__ == "__main__":
    raise SystemExit(HeadlessProxyshop().render_all())
'@
}


function Test-ProxyshopSourceCheckout {
    return (
        (Test-Path -LiteralPath (Join-Path $ProxyshopDir "main.py") -PathType Leaf) -and
        (Test-Path -LiteralPath (Join-Path $ProxyshopDir "pyproject.toml") -PathType Leaf) -and
        (Test-Path -LiteralPath (Join-Path $ProxyshopDir "src") -PathType Container)
    )
}


function Get-EffectiveProxyshopMode {
    if ($ProxyshopMode -ne "Auto") {
        return $ProxyshopMode
    }

    if (Test-ProxyshopSourceCheckout) {
        return "SourceHeadless"
    }
    return "GuiManual"
}


function Repair-ProxyshopPython310Compatibility {
    param([Parameter(Mandatory = $true)][string]$LogPath)

    # Proxyshop 1.13.2 declares Python 3.10 support, but this annotation uses
    # PEP 604 union syntax with typing.ForwardRef. Python 3.10 cannot evaluate
    # that combination at import time. typing.Union with a string forward
    # reference is equivalent and works on every supported Python version.
    $layoutsPath = Join-Path $ProxyshopDir "src\layouts.py"
    if (-not (Test-Path -LiteralPath $layoutsPath -PathType Leaf)) {
        return Stop-AtGate `
            -Title "Proxyshop layouts.py was not found." `
            -Message "SourceHeadless mode expected this file: $layoutsPath" `
            -LogPaths @()
    }

    $source = Get-Content -LiteralPath $layoutsPath -Raw -Encoding UTF8
    $updated = $source.Replace(
        "from typing import Optional, Match, Union, Type, ForwardRef",
        "from typing import Optional, Match, Union, Type"
    ).Replace(
        "def assign_layout(filename: Path) -> str | ForwardRef('CardLayout'):",
        "def assign_layout(filename: Path) -> Union[str, 'CardLayout']:"
    )

    if ([string]::Equals($source, $updated, [System.StringComparison]::Ordinal)) {
        return $true
    }

    $backupPath = "$layoutsPath.orchestrator-backup"
    if (-not (Test-Path -LiteralPath $backupPath -PathType Leaf)) {
        Copy-Item -LiteralPath $layoutsPath -Destination $backupPath
    }

    Set-Content -LiteralPath $layoutsPath -Value $updated -Encoding UTF8
    $message = "Patched src/layouts.py for Python 3.10; original saved as $backupPath"
    Write-WarningLine $message
    Add-Content -LiteralPath $LogPath -Value $message -Encoding UTF8
    return $true
}


function Ensure-ProxyshopDependencies {
    param([Parameter(Mandatory = $true)][string]$Poetry)

    if (-not $InstallProxyshopDependencies) {
        return $true
    }

    $preflightLog = Join-Path $script:RunLogDir "02_proxyshop_dependency_check.log"
    $preflightCode = Invoke-NativeLogged `
        -FilePath $Poetry `
        -Arguments @("run", "python", "-c", "import photoshop; import typing_extensions") `
        -WorkingDirectory $ProxyshopDir `
        -LogPath $preflightLog

    if ($preflightCode -eq 0) {
        Write-Host "Proxyshop dependency check passed." -ForegroundColor Green
        return $true
    }

    Write-WarningLine "Proxyshop source dependencies are missing; running poetry install --no-root."
    $installLog = Join-Path $script:RunLogDir "02_proxyshop_poetry_install.log"
    $installCode = Invoke-NativeLogged `
        -FilePath $Poetry `
        -Arguments @("install", "--no-root") `
        -WorkingDirectory $ProxyshopDir `
        -LogPath $installLog

    if ($installCode -ne 0) {
        return Stop-AtGate `
            -Title "poetry install --no-root failed for Proxyshop." `
            -Message "Review the dependency installation output before deciding whether to continue." `
            -LogPaths @($installLog, $preflightLog)
    }

    $retryLog = Join-Path $script:RunLogDir "02_proxyshop_dependency_recheck.log"
    $retryCode = Invoke-NativeLogged `
        -FilePath $Poetry `
        -Arguments @("run", "python", "-c", "import photoshop; import typing_extensions") `
        -WorkingDirectory $ProxyshopDir `
        -LogPath $retryLog

    if ($retryCode -ne 0) {
        return Stop-AtGate `
            -Title "Proxyshop dependencies are still unavailable." `
            -Message "poetry install --no-root completed, but the Proxyshop environment still cannot import photoshop and typing_extensions." `
            -LogPaths @($retryLog, $installLog)
    }

    Write-Host "Proxyshop dependencies installed." -ForegroundColor Green
    return $true
}


function Invoke-ProxyshopHeadlessStage {
    $poetry = Resolve-CommandPath -Command $PoetryExe
    if (-not (Ensure-ProxyshopDependencies -Poetry $poetry)) {
        return $false
    }

    $helperPath = Join-Path $ProxyshopDir "orchestrator_render_all_headless.py"
    $logPath = Join-Path $script:RunLogDir "02_proxyshop_headless.log"

    if (-not (Repair-ProxyshopPython310Compatibility -LogPath $logPath)) {
        return $false
    }

    Set-Content -LiteralPath $helperPath -Value (Get-HeadlessProxyshopRunnerCode) -Encoding UTF8
    try {
        $exitCode = Invoke-NativeLogged `
            -FilePath $poetry `
            -Arguments @("run", "python", $helperPath) `
            -WorkingDirectory $ProxyshopDir `
            -LogPath $logPath
    }
    finally {
        Remove-Item -LiteralPath $helperPath -Force -ErrorAction SilentlyContinue
    }

    $failedLog = Join-Path $ProxyshopDir "logs\failed.txt"
    $errorLog = Join-Path $ProxyshopDir "logs\error.txt"

    if ($exitCode -ne 0) {
        return Stop-AtGate `
            -Title "Headless Proxyshop exited with code $exitCode." `
            -Message "Review the render log, Proxyshop failed log, and Proxyshop error log before continuing." `
            -LogPaths @($logPath, $failedLog, $errorLog)
    }

    if ((Test-Path -LiteralPath $failedLog -PathType Leaf) -and
        ((Get-Item -LiteralPath $failedLog).Length -gt 0)) {
        return Stop-AtGate `
            -Title "Proxyshop reported failed renders." `
            -Message "Review the failed render list. Continue only if the remaining cards are enough for this run." `
            -LogPaths @($failedLog, $errorLog, $logPath)
    }

    Write-Host "Headless Proxyshop render completed." -ForegroundColor Green
    return $true
}


function Invoke-ProxyshopGuiManualStage {
    $proxyshopExe = Resolve-RequiredFile -Path (Join-Path $ProxyshopDir "Proxyshop.exe") -Description "Proxyshop.exe"
    $failedLog = Join-Path $ProxyshopDir "logs\failed.txt"
    $errorLog = Join-Path $ProxyshopDir "logs\error.txt"

    Write-WarningLine "The GitHub Proxyshop.exe release does not provide a headless Render All command."
    Write-WarningLine "Poetry is only required for MTG-Art-Downloader in this workflow. SourceHeadless is an optional advanced mode for a separate dependency-installed Proxyshop source checkout."
    Write-Host ""
    Write-Host "Proxyshop will now open. In Proxyshop:" -ForegroundColor Yellow
    Write-Host "  1. Confirm the correct templates/settings are selected."
    Write-Host "  2. Click Render All."
    Write-Host "  3. Respond to any Proxyshop errors or Continue prompts."
    Write-Host "  4. Return to this window only after rendering is finished."

    Start-Process -FilePath $proxyshopExe -WorkingDirectory $ProxyshopDir | Out-Null

    $ready = Read-ContinueOrQuit `
        -Title "Manual Proxyshop render gate" `
        -Message "After Proxyshop's Render All operation has completely finished, type c. Type q to stop." `
        -LogPaths @()
    if (-not $ready) {
        return $false
    }

    if ((Test-Path -LiteralPath $failedLog -PathType Leaf) -and
        ((Get-Item -LiteralPath $failedLog).Length -gt 0)) {
        return Stop-AtGate `
            -Title "Proxyshop reported failed renders." `
            -Message "Review the failed render list. Continue only if the remaining cards are enough for this run." `
            -LogPaths @($failedLog, $errorLog)
    }

    Write-Host "Manual Proxyshop render gate accepted." -ForegroundColor Green
    return $true
}


function Invoke-ProxyshopStage {
    if ($SkipProxyshop) {
        Write-Section "Stage 3/4: Proxyshop (skipped)"
        return $true
    }

    Write-Section "Stage 3/4: Proxyshop"

    $failedLog = Join-Path $ProxyshopDir "logs\failed.txt"
    $errorLog = Join-Path $ProxyshopDir "logs\error.txt"
    Copy-ToRunLogArchive -Path $failedLog -Prefix "proxyshop"
    Copy-ToRunLogArchive -Path $errorLog -Prefix "proxyshop"

    $effectiveMode = Get-EffectiveProxyshopMode
    Write-Host "Proxyshop mode: $effectiveMode"

    if (($effectiveMode -eq "SourceHeadless") -and -not (Test-ProxyshopSourceCheckout)) {
        Write-WarningLine "ProxyshopDir is not a complete Proxyshop source checkout: $ProxyshopDir"
        Write-WarningLine "Expected main.py, pyproject.toml, and the src folder. The normal Proxyshop.exe release must use GuiManual mode."

        if ($ProxyshopMode -eq "Auto") {
            $effectiveMode = "GuiManual"
        }
        else {
            $fallBack = Read-ContinueOrQuit `
                -Title "SourceHeadless mode is unavailable." `
                -Message "Type c to fall back to the normal Proxyshop.exe GUI mode, or q to stop." `
                -LogPaths @()
            if (-not $fallBack) {
                return $false
            }
            $effectiveMode = "GuiManual"
        }
        Write-Host "Proxyshop mode changed to: $effectiveMode" -ForegroundColor Yellow
    }

    if ($effectiveMode -eq "SourceHeadless") {
        return Invoke-ProxyshopHeadlessStage
    }

    return Invoke-ProxyshopGuiManualStage
}


function Resolve-PythonInvocation {
    $fallbacks = @(
        "$env:LocalAppData\Programs\Python\Python312\python.exe",
        "$env:LocalAppData\Programs\Python\Python311\python.exe",
        "$env:LocalAppData\Programs\Python\Python310\python.exe",
        "C:\Python312\python.exe",
        "C:\Python311\python.exe",
        "C:\Python310\python.exe"
    )

    $python = Resolve-CommandPath -Command $PythonExe -FallbackPaths $fallbacks
    $arguments = @()
    if ((Split-Path -Path $python -Leaf) -ieq "py.exe" -or $PythonExe -ieq "py") {
        $arguments += "-3"
    }
    return [pscustomobject]@{ Path = $python; PrefixArguments = $arguments }
}


function Invoke-AssemblyStage {
    if ($SkipAssembly) {
        Write-Section "Stage 4/4: Print-sheet assembly (skipped)"
        return $true
    }

    Write-Section "Stage 4/4: Print-sheet assembly"

    $assembler = Resolve-RequiredFile -Path $AssemblerScript -Description "assemble_card_sheets.py"
    $back = $null
    if (-not $SingleSided) {
        $back = Resolve-RequiredFile -Path $BackPdf -Description "card-back PDF"
    }
    $cardsTxt = Resolve-RequiredFile -Path (Join-Path $DownloaderDir "cards.txt") -Description "cards.txt"
    $renderFolder = Join-Path $ProxyshopDir "out"
    if (-not (Test-Path -LiteralPath $renderFolder -PathType Container)) {
        throw "Proxyshop output folder was not found: $renderFolder"
    }

    $frontsBase = Join-Path $script:RunDir "merge1.pdf"
    $outputBase = Join-Path $script:RunDir "OutputMerge.pdf"
    $logPath = Join-Path $script:RunLogDir "03_assemble_card_sheets.log"
    $python = Resolve-PythonInvocation

    $arguments = @()
    $arguments += $python.PrefixArguments
    $arguments += @(
        $assembler,
        "--input-folder", $renderFolder,
        "--cards-txt", $cardsTxt,
        "--fronts-pdf", $frontsBase,
        "--output", $outputBase
    )

    if ($SingleSided) {
        $arguments += "--single-sided"
    }
    else {
        $arguments += @(
            "--back-pdf", $back,
            "--duplex-binding", $DuplexBinding
        )
    }

    $exitCode = Invoke-NativeLogged `
        -FilePath $python.Path `
        -Arguments $arguments `
        -WorkingDirectory $ProxyshopDir `
        -LogPath $logPath

    $logText = ""
    if (Test-Path -LiteralPath $logPath -PathType Leaf) {
        $logText = Get-Content -LiteralPath $logPath -Raw -ErrorAction SilentlyContinue
    }

    if ($exitCode -ne 0) {
        return Stop-AtGate `
            -Title "Print-sheet assembly exited with code $exitCode." `
            -Message "Review the assembly errors before deciding whether to continue." `
            -LogPaths @($logPath)
    }

    if ($logText -match "(?m)^ERROR:") {
        return Stop-AtGate `
            -Title "Print-sheet assembly reported missing or invalid cards." `
            -Message "The script completed, but at least one requested card was missing or invalid." `
            -LogPaths @($logPath)
    }

    $outputs = Get-FinalOutputPdfs
    if ($outputs.Count -eq 0) {
        return Stop-AtGate `
            -Title "No OutputMerge PDF files were created." `
            -Message "Review the assembly log before continuing." `
            -LogPaths @($logPath)
    }

    Write-Host "Created $($outputs.Count) output PDF(s):" -ForegroundColor Green
    foreach ($output in $outputs) {
        Write-Host "  $($output.FullName)"
    }
    return $true
}


function Get-FinalOutputPdfs {
    if (-not $script:RunDir -or -not (Test-Path -LiteralPath $script:RunDir -PathType Container)) {
        return @()
    }

    return @(
        Get-ChildItem -LiteralPath $script:RunDir -File -Filter "OutputMerge*.pdf" |
            Where-Object { $_.BaseName -match "^OutputMerge(_\d+)?$" } |
            Sort-Object {
                if ($_.BaseName -eq "OutputMerge") { 1 }
                else { [int]($_.BaseName -replace "^OutputMerge_", "") }
            }
    )
}


function Resolve-DefaultPrinterName {
    if ($PrinterName) {
        $printer = Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue
        if (-not $printer) {
            throw "Printer was not found: $PrinterName"
        }
        return $printer.Name
    }

    $defaultPrinter = Get-CimInstance -ClassName Win32_Printer -ErrorAction SilentlyContinue |
        Where-Object { $_.Default -eq $true } |
        Select-Object -First 1
    if ($defaultPrinter) {
        Write-Host "Using Windows default printer: $($defaultPrinter.Name)" -ForegroundColor Cyan
        return $defaultPrinter.Name
    }

    throw "No Windows default printer was found. Provide -PrinterName explicitly."
}


function Find-SumatraPdf {
    if ($SumatraPdfExe) {
        return (Resolve-RequiredFile -Path $SumatraPdfExe -Description "SumatraPDF.exe")
    }

    $candidates = @(
        "$env:LOCALAPPDATA\SumatraPDF\SumatraPDF.exe",
        "$env:ProgramFiles\SumatraPDF\SumatraPDF.exe",
        "${env:ProgramFiles(x86)}\SumatraPDF\SumatraPDF.exe"
    )

    $command = Get-Command "SumatraPDF.exe" -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }
    $command = Get-Command "SumatraPDF" -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    foreach ($candidate in $candidates) {
        $expanded = [Environment]::ExpandEnvironmentVariables($candidate)
        if (Test-Path -LiteralPath $expanded -PathType Leaf) {
            return (Resolve-Path -LiteralPath $expanded).Path
        }
    }

    try {
        $appPath = Get-ItemProperty -LiteralPath "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\SumatraPDF.exe" -ErrorAction Stop
        if ($appPath."(default)" -and (Test-Path -LiteralPath $appPath."(default)" -PathType Leaf)) {
            return $appPath."(default)"
        }
    }
    catch {
        # Not installed in the common registry location.
    }

    return $null
}


function Find-AcrobatReader {
    $candidates = @(
        "$env:ProgramFiles\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
        "${env:ProgramFiles(x86)}\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
        "$env:ProgramFiles\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
        "${env:ProgramFiles(x86)}\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
    )

    foreach ($candidate in $candidates) {
        $expanded = [Environment]::ExpandEnvironmentVariables($candidate)
        if (Test-Path -LiteralPath $expanded -PathType Leaf) {
            return (Resolve-Path -LiteralPath $expanded).Path
        }
    }

    foreach ($registryPath in @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Acrobat.exe",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\AcroRd32.exe"
    )) {
        try {
            $item = Get-ItemProperty -LiteralPath $registryPath -ErrorAction Stop
            if ($item."(default)" -and (Test-Path -LiteralPath $item."(default)" -PathType Leaf)) {
                return $item."(default)"
            }
        }
        catch {
            # Try the next known location.
        }
    }

    return $null
}


function Read-SelectedPrintPreset {
    Write-Host ""
    if ($PrintPreset) {
        Write-Host "Configured printing preset: $PrintPreset" -ForegroundColor Cyan
        $entered = Read-Host "Enter a printer preset to use, or press Enter to keep the configured preset"
        if ([string]::IsNullOrWhiteSpace($entered)) {
            return $PrintPreset
        }
        return $entered.Trim()
    }

    $entered = Read-Host "Optional printer preset to use; press Enter to use the printer defaults"
    if ([string]::IsNullOrWhiteSpace($entered)) {
        return ""
    }
    return $entered.Trim()
}


function Set-PrinterPresetWithUIA {
    param(
        [Parameter(Mandatory = $true)][string]$Printer,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Preset
    )

    try {
        Add-Type -AssemblyName UIAutomationClient
        Add-Type -AssemblyName UIAutomationTypes
    }
    catch {
        Write-WarningLine "Windows UI Automation is unavailable: $($_.Exception.Message)"
        return $false
    }

    $argument = "printui.dll,PrintUIEntry /n `"$Printer`" /e"
    $process = Start-Process -FilePath "rundll32.exe" -ArgumentList $argument -PassThru

    try {
        $deadline = (Get-Date).AddSeconds(12)
        $window = $null

        while ((Get-Date) -lt $deadline -and -not $window) {
            Start-Sleep -Milliseconds 400
            $root = [System.Windows.Automation.AutomationElement]::RootElement
            $children = $root.FindAll(
                [System.Windows.Automation.TreeScope]::Children,
                [System.Windows.Automation.Condition]::TrueCondition
            )

            for ($index = 0; $index -lt $children.Count; $index++) {
                $candidate = $children.Item($index)
                $candidateName = ""
                $candidateProcessId = -1
                try {
                    $candidateName = $candidate.Current.Name
                    $candidateProcessId = $candidate.Current.ProcessId
                }
                catch {
                    continue
                }

                if (($candidateProcessId -eq $process.Id) -or
                    ($candidateName -match [regex]::Escape($Printer)) -or
                    ($candidateName -match "Printing Preferences|Printer Preferences")) {
                    $window = $candidate
                    break
                }
            }
        }

        if (-not $window) {
            Write-WarningLine "Could not find the printer preferences window."
            return $false
        }

        $presetCondition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::NameProperty,
            $Preset
        )
        $presetElement = $window.FindFirst(
            [System.Windows.Automation.TreeScope]::Descendants,
            $presetCondition
        )

        if (-not $presetElement) {
            Write-WarningLine "Could not find a printer preset named '$Preset' in the preferences window."
            return $false
        }

        $selected = $false
        try {
            $selectionPattern = $presetElement.GetCurrentPattern(
                [System.Windows.Automation.SelectionItemPattern]::Pattern
            ) -as [System.Windows.Automation.SelectionItemPattern]
            if ($selectionPattern) {
                $selectionPattern.Select()
                $selected = $true
            }
        }
        catch {
            # Try Invoke below.
        }

        if (-not $selected) {
            try {
                $invokePattern = $presetElement.GetCurrentPattern(
                    [System.Windows.Automation.InvokePattern]::Pattern
                ) -as [System.Windows.Automation.InvokePattern]
                if ($invokePattern) {
                    $invokePattern.Invoke()
                    $selected = $true
                }
            }
            catch {
                # Try walking to an invokable parent below.
            }
        }

        if (-not $selected) {
            $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
            $parent = $walker.GetParent($presetElement)
            while ($parent -and -not $selected) {
                try {
                    $invokePattern = $parent.GetCurrentPattern(
                        [System.Windows.Automation.InvokePattern]::Pattern
                    ) -as [System.Windows.Automation.InvokePattern]
                    if ($invokePattern) {
                        $invokePattern.Invoke()
                        $selected = $true
                    }
                }
                catch {
                    # Keep walking upward.
                }
                $parent = $walker.GetParent($parent)
            }
        }

        if (-not $selected) {
            Write-WarningLine "Found '$Preset', but Windows UI Automation could not select it."
            return $false
        }

        Start-Sleep -Milliseconds 700

        $okNameCondition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::NameProperty,
            "OK"
        )
        $okTypeCondition = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::Button
        )
        $okCondition = New-Object System.Windows.Automation.AndCondition(
            $okNameCondition,
            $okTypeCondition
        )
        $okButton = $window.FindFirst(
            [System.Windows.Automation.TreeScope]::Descendants,
            $okCondition
        )

        if (-not $okButton) {
            Write-WarningLine "Could not find the OK button in printer preferences."
            return $false
        }

        $okInvoke = $okButton.GetCurrentPattern(
            [System.Windows.Automation.InvokePattern]::Pattern
        ) -as [System.Windows.Automation.InvokePattern]
        if (-not $okInvoke) {
            Write-WarningLine "The printer preferences OK button did not expose an invoke action."
            return $false
        }

        $okInvoke.Invoke()
        Wait-Process -Id $process.Id -Timeout 10 -ErrorAction SilentlyContinue
        Write-Host "Selected printer preset: $Preset" -ForegroundColor Green
        return $true
    }
    catch {
        Write-WarningLine "Could not apply the printer preset automatically: $($_.Exception.Message)"
        return $false
    }
    finally {
        if ($process -and -not $process.HasExited) {
            try { $process.CloseMainWindow() | Out-Null } catch { }
            Start-Sleep -Milliseconds 500
            if (-not $process.HasExited) {
                try { $process.Kill() } catch { }
            }
        }
    }
}


function Print-WithManualDialog {
    param(
        [Parameter(Mandatory = $true)][System.IO.FileInfo[]]$Pdfs,
        [Parameter(Mandatory = $true)][string]$Printer,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Preset,
        [switch]$SingleSided
    )

    $sumatra = Find-SumatraPdf
    $acrobat = Find-AcrobatReader

    Write-Host ""
    Write-Host "Manual print mode" -ForegroundColor Cyan
    Write-Host "Printer: $Printer"
    if ($Preset) {
        Write-Host "Select this printer preset in the print system dialog: $Preset" -ForegroundColor Yellow
    }
    else {
        Write-Host "Select the desired paper, quality, scaling, and duplex settings in the print system dialog." -ForegroundColor Yellow
    }
    if ($SingleSided) {
        Write-Host "Turn off two-sided/duplex printing for this single-sided output." -ForegroundColor Yellow
    }

    if ($sumatra) {
        Write-Host "Using SumatraPDF print dialogs: $sumatra"
    }
    elseif ($acrobat) {
        Write-Host "Using Adobe Acrobat/Reader print dialogs: $acrobat"
    }
    else {
        Write-WarningLine "Neither SumatraPDF nor Adobe Acrobat/Reader was found. PDFs will be opened in the default viewer; press Ctrl+P manually."
    }

    $index = 0
    foreach ($pdf in $Pdfs) {
        $index++
        Write-Host ""
        Write-Host "Print file $index of $($Pdfs.Count): $($pdf.FullName)" -ForegroundColor Cyan

        if ($sumatra) {
            $arguments = "-print-dialog -exit-when-done `"$($pdf.FullName)`""
            $process = Start-Process -FilePath $sumatra -ArgumentList $arguments -PassThru -Wait
            if ($process.ExitCode -ne 0) {
                Write-WarningLine "SumatraPDF returned exit code $($process.ExitCode). Verify that the print was sent."
            }
        }
        elseif ($acrobat) {
            $arguments = "/p `"$($pdf.FullName)`""
            Start-Process -FilePath $acrobat -ArgumentList $arguments | Out-Null
            Read-Host "After Acrobat's print dialog has been handled, press Enter"
        }
        else {
            Start-Process -FilePath $pdf.FullName | Out-Null
            Read-Host "Print from the default viewer, then press Enter"
        }

        if ($index -lt $Pdfs.Count) {
            while ($true) {
                $response = Read-Host "Type n for the next PDF, or s to stop printing"
                if ($response -match "^(?i)n$") { break }
                if ($response -match "^(?i)s$") { return }
                Write-Host "Please enter n or s."
            }
        }
    }
}


function Print-Automatically {
    param(
        [Parameter(Mandatory = $true)][System.IO.FileInfo[]]$Pdfs,
        [Parameter(Mandatory = $true)][string]$Printer,
        [switch]$SingleSided
    )

    $sumatra = Find-SumatraPdf
    if (-not $sumatra) {
        Write-WarningLine "Automatic PDF printing requires SumatraPDF, but it was not found."
        return $false
    }

    $printSettings = "paper=letter,noscale,center"
    if ($SingleSided) {
        $printSettings += ",simplex"
    }

    foreach ($pdf in $Pdfs) {
        Write-Host "Printing: $($pdf.FullName)"
        $arguments = (
            "-print-to `"$Printer`" " +
            "-print-settings `"$printSettings`" " +
            "`"$($pdf.FullName)`""
        )
        $process = Start-Process -FilePath $sumatra -ArgumentList $arguments -PassThru -Wait
        if ($process.ExitCode -ne 0) {
            Write-ErrorLine "SumatraPDF failed with exit code $($process.ExitCode) for $($pdf.Name)."
            return $false
        }
    }

    return $true
}


function Invoke-ReviewAndPrintStage {
    if ($SkipPrint -or $PrintMode -eq "Off") {
        Write-Section "Review and printing (skipped)"
        return $true
    }

    Write-Section "Review and printing"

    $pdfs = Get-FinalOutputPdfs
    if ($pdfs.Count -eq 0) {
        Write-WarningLine "No final OutputMerge PDF files are available to review."
        return $true
    }

    Write-Host "Review these output PDF(s) before printing:" -ForegroundColor Cyan
    foreach ($pdf in $pdfs) {
        Write-Host "  $($pdf.FullName)"
    }

    try {
        Start-Process -FilePath "explorer.exe" -ArgumentList "`"$script:RunDir`"" | Out-Null
    }
    catch {
        Write-WarningLine "Could not open the output folder in Explorer."
    }

    while ($true) {
        $response = Read-Host "Do the output PDFs look correct? Type p to print, or q to quit without printing"
        if ($response -match "^(?i)p$") { break }
        if ($response -match "^(?i)q$") {
            Write-Host "Printing cancelled. Output files remain in: $script:RunDir" -ForegroundColor Yellow
            return $true
        }
        Write-Host "Please enter p or q."
    }

    $printer = Resolve-DefaultPrinterName
    $preset = ""
    if ($PrintMode -ne "AutomaticDefault") {
        $preset = Read-SelectedPrintPreset
    }

    $presetDisplay = if ($preset) { $preset } else { "(printer defaults)" }
    Write-Host "Printer: $printer" -ForegroundColor Cyan
    Write-Host "Preset:  $presetDisplay" -ForegroundColor Cyan

    if ($PrintMode -eq "ManualDialog") {
        Print-WithManualDialog -Pdfs $pdfs -Printer $printer -Preset $preset -SingleSided:$SingleSided
        return $true
    }

    if ($PrintMode -eq "ExperimentalPreset") {
        $applied = $false
        if ($preset) {
            $applied = Set-PrinterPresetWithUIA -Printer $printer -Preset $preset
        }
        else {
            Write-WarningLine "ExperimentalPreset requires a named printer preset; falling back to manual dialogs."
        }

        if (-not $applied) {
            $continue = Read-ContinueOrQuit `
                -Title "The printer preset could not be applied automatically." `
                -Message "Choose c to fall back to manual print dialogs, or q to stop before printing." `
                -LogPaths @()
            if (-not $continue) {
                return $false
            }
            Print-WithManualDialog -Pdfs $pdfs -Printer $printer -Preset $preset -SingleSided:$SingleSided
            return $true
        }
    }

    $printed = Print-Automatically -Pdfs $pdfs -Printer $printer -SingleSided:$SingleSided
    if (-not $printed) {
        $continue = Read-ContinueOrQuit `
            -Title "Automatic printing failed." `
            -Message "Choose c to fall back to manual print dialogs, or q to stop." `
            -LogPaths @()
        if (-not $continue) {
            return $false
        }
        Print-WithManualDialog -Pdfs $pdfs -Printer $printer -Preset $preset -SingleSided:$SingleSided
    }

    return $true
}


function Test-Configuration {
    Write-Section "Configuration"

    if ([string]::IsNullOrWhiteSpace($DownloaderDir)) {
        $DownloaderDir = Join-Path $script:ScriptRoot "MTG-Art-Downloader"
    }
    if ([string]::IsNullOrWhiteSpace($ProxyshopDir)) {
        $ProxyshopDir = Join-Path $script:ScriptRoot "Proxyshop"
    }
    if ([string]::IsNullOrWhiteSpace($AssemblerScript)) {
        $AssemblerScript = Join-Path $script:ScriptRoot "assemble_card_sheets.py"
    }
    if ([string]::IsNullOrWhiteSpace($BackPdf)) {
        $BackPdf = Join-Path $script:ScriptRoot "horizontal_cardBack_3x3.pdf"
    }

    $script:DownloaderDir = Resolve-RequiredDirectory -Path $DownloaderDir -Description "MTG Art Downloader folder"
    $script:ProxyshopDir = Resolve-RequiredDirectory -Path $ProxyshopDir -Description "Proxyshop folder"
    $script:AssemblerScript = [System.IO.Path]::GetFullPath($AssemblerScript)
    $script:BackPdf = [System.IO.Path]::GetFullPath($BackPdf)

    $backDisplay = if ($SingleSided) { "(disabled; single-sided output)" } else { $BackPdf }
    $duplexDisplay = if ($SingleSided) { "(not used; single-sided output)" } else { $DuplexBinding }

    Write-Host "Downloader:  $DownloaderDir"
    Write-Host "Proxyshop:   $ProxyshopDir"
    Write-Host "Assembler:   $AssemblerScript"
    Write-Host "Back PDF:    $backDisplay"
    Write-Host "Clean art:   $CleanProxyshopArt"
    Write-Host "Print mode:  $PrintMode"
    Write-Host "Duplex:      $duplexDisplay"

    Resolve-RequiredFile -Path (Join-Path $DownloaderDir "cards.txt") -Description "cards.txt" | Out-Null

    if (-not $SkipAssembly) {
        Resolve-RequiredFile -Path $AssemblerScript -Description "assemble_card_sheets.py" | Out-Null
        if (-not $SingleSided) {
            Resolve-RequiredFile -Path $BackPdf -Description "card-back PDF" | Out-Null
        }
    }
}


function Invoke-Workflow {
    Test-Configuration
    Initialize-RunFolders

    if (-not (Invoke-DownloaderStage)) {
        return 1
    }
    if (-not (Copy-DownloadedArtStage)) {
        return 1
    }
    if (-not (Invoke-ProxyshopStage)) {
        return 1
    }
    if (-not (Invoke-AssemblyStage)) {
        return 1
    }
    if (-not (Invoke-ReviewAndPrintStage)) {
        return 1
    }

    Write-Section "Workflow complete"
    Write-Host "Run output: $script:RunDir" -ForegroundColor Green
    return 0
}


try {
    $exitCode = Invoke-Workflow
    exit $exitCode
}
catch {
    Write-Host ""
    Write-ErrorLine $_.Exception.Message
    if ($_.InvocationInfo) {
        Write-Host "At: $($_.InvocationInfo.PositionMessage)" -ForegroundColor Red
    }
    if ($script:RunDir) {
        Write-Host "Run folder: $script:RunDir" -ForegroundColor Yellow
    }
    exit 1
}
