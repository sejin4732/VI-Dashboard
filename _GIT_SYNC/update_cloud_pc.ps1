$ErrorActionPreference = "Stop"

function Write-Step {
    param([string]$Message)
    Write-Host "[CLOUD] $Message" -ForegroundColor Cyan
}

$projectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $projectRoot

$gitCommand = Get-Command git -ErrorAction SilentlyContinue
if (-not $gitCommand) {
    Write-Host "Git이 설치되어 있지 않습니다. 먼저 Git for Windows를 설치해 주세요." -ForegroundColor Red
    exit 1
}

$pythonCommand = Get-Command python -ErrorAction SilentlyContinue
if (-not $pythonCommand) {
    Write-Host "python 명령을 찾을 수 없습니다. Python 설치 또는 PATH를 확인해 주세요." -ForegroundColor Red
    exit 1
}

if (-not (Test-Path (Join-Path $projectRoot ".git"))) {
    Write-Host "이 폴더는 아직 clone 되지 않았습니다." -ForegroundColor Yellow
    Write-Host "먼저 아래처럼 실행해 주세요."
    Write-Host "git clone <REPO_URL> VICT"
    exit 1
}

Write-Step "최신 코드 가져오기"
git pull | Out-Host

Write-Step "requirements 설치/업데이트"
python -m pip install -r requirements.txt | Out-Host

$configPath = Join-Path $projectRoot "config.json"
$dataRoot = "D:/VICT_DATA/data"
$bomLakeRoot = "D:/VICT_DATA/bom_lake"
$cacheRoot = "D:/VICT_DATA/cache"
$outputRoot = "D:/VICT_DATA/outputs"
$logRoot = "D:/VICT_DATA/logs"

$requiredFolders = @(
    "D:/VICT_DATA",
    $dataRoot,
    $bomLakeRoot,
    $cacheRoot,
    $outputRoot,
    $logRoot
)

foreach ($folder in $requiredFolders) {
    if (-not (Test-Path $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        Write-Step "폴더 생성: $folder"
    }
}

if (-not (Test-Path $configPath)) {
    Write-Step "config.json이 없어 기본 파일을 생성합니다."
    $configObject = [ordered]@{
        app_version   = "0.1.0"
        data_root     = $dataRoot
        bom_lake_root = $bomLakeRoot
        cache_root    = $cacheRoot
        output_root   = $outputRoot
        log_root      = $logRoot
    }
    $configObject | ConvertTo-Json -Depth 3 | Set-Content -Path $configPath -Encoding UTF8
}
else {
    Write-Step "기존 config.json을 유지합니다."
}

Write-Host ""
Write-Host "실행 명령:" -ForegroundColor Yellow
Write-Host "python vi_report_gui.py"
