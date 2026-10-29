<#
.SYNOPSIS
  Build the Windows app: PyInstaller (one folder) -> smoke test -> zip.

.EXAMPLE
  .\build.ps1                 # tests + build + zip
  .\build.ps1 -SkipTests      # faster, when the tests were just run
  .\build.ps1 -SkipSelfTest   # when you cannot run a GUI process (rare)

Output: dist\ResultTestGenerator-<version>-win64.zip
Needs: pip install -r requirements.txt   (includes pyinstaller)
#>
param(
    [switch]$SkipTests,
    [switch]$SkipSelfTest
)

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

function Invoke-Step([string]$Title, [scriptblock]$Block) {
    Write-Host "==> $Title" -ForegroundColor Cyan
    # Windows PowerShell 5.1 turns native stderr (PyInstaller logs there) into errors;
    # judge the step by its exit code instead.
    $previous = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try { & $Block } finally { $ErrorActionPreference = $previous }
    if ($LASTEXITCODE -ne 0) { throw "$Title failed (exit code $LASTEXITCODE)" }
}

$appName = "Result Test Generator"
$version = (python -c "import rtg; print(rtg.__version__)").Trim()
if (-not $version) { throw "Could not read rtg.__version__" }
$zipName = "ResultTestGenerator-$version-win64.zip"

if (-not $SkipTests) {
    Invoke-Step "Run tests" { python -m pytest }
}

Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue

Invoke-Step "PyInstaller ($appName $version)" {
    python -m PyInstaller --noconfirm --clean --onedir --windowed `
        --name $appName `
        --add-data "rtg/data/column_aliases.json;rtg/data" `
        --collect-all qtawesome `
        --hidden-import pynput.keyboard._win32 `
        --hidden-import pynput.mouse._win32 `
        app.py
}

$exe = Join-Path $PSScriptRoot "dist\$appName\$appName.exe"
if (-not (Test-Path $exe)) { throw "Build did not produce $exe" }

if (-not $SkipSelfTest) {
    Write-Host "==> Smoke test (--selftest)" -ForegroundColor Cyan
    $env:QT_QPA_PLATFORM = "offscreen"
    $process = Start-Process -FilePath $exe -ArgumentList "--selftest" -Wait -PassThru
    Remove-Item Env:QT_QPA_PLATFORM
    if ($process.ExitCode -ne 0) {
        throw "Self-test failed (exit code $($process.ExitCode)): the packaged app does not start correctly."
    }
}

Write-Host "==> Zip" -ForegroundColor Cyan
$zipPath = Join-Path $PSScriptRoot "dist\$zipName"
Compress-Archive -Path (Join-Path $PSScriptRoot "dist\$appName") -DestinationPath $zipPath -Force

$sizeMb = [math]::Round((Get-Item $zipPath).Length / 1MB, 1)
Write-Host "Done: $zipPath ($sizeMb MB)" -ForegroundColor Green
