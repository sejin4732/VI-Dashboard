$ErrorActionPreference = "Stop"

function Write-Step {
    param([string]$Message)
    Write-Host "[SETUP] $Message" -ForegroundColor Cyan
}

$projectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $projectRoot

$gitCommand = Get-Command git -ErrorAction SilentlyContinue
if (-not $gitCommand) {
    Write-Host "Git이 설치되어 있지 않습니다. 먼저 Git for Windows를 설치해 주세요." -ForegroundColor Red
    exit 1
}

Write-Step "프로젝트 경로: $projectRoot"

if (-not (Test-Path (Join-Path $projectRoot ".git"))) {
    Write-Step "Git 저장소가 없어 새로 초기화합니다."
    try {
        git init -b main | Out-Host
    }
    catch {
        git init | Out-Host
        git branch -M main | Out-Host
    }
}
else {
    Write-Step "기존 Git 저장소를 사용합니다."
}

if (git rev-parse --verify HEAD 2>$null) {
    Write-Step "기존 커밋이 있어 staged 파일을 정리합니다."
    git reset | Out-Host
}
else {
    Write-Step "아직 커밋이 없어 staged 정리는 건너뜁니다."
}

Write-Step ".gitignore 기준으로 소스 파일만 stage 합니다."
git add .gitignore requirements.txt | Out-Host
git add _GIT_SYNC/config.cloud.example.json _GIT_SYNC/README_DEV_SYNC.md _GIT_SYNC/setup_git_personal_pc.ps1 _GIT_SYNC/update_cloud_pc.ps1 | Out-Host
git add *.py *.md *.txt *.json 2>$null | Out-Null
git add dashboard compat_imports workflow 2>$null | Out-Host

$hasCommit = $false
git rev-parse --verify HEAD 2>$null | Out-Null
if ($LASTEXITCODE -eq 0) {
    $hasCommit = $true
}

if (-not $hasCommit) {
    Write-Step "첫 커밋을 생성합니다."
    git commit -m "Initial VICT source setup" | Out-Host
}
else {
    Write-Step "기존 저장소입니다. 필요하면 직접 commit 해주세요."
}

Write-Host ""
Write-Host "다음 수동 단계:" -ForegroundColor Yellow
Write-Host "git remote add origin <REPO_URL>"
Write-Host "git push -u origin main"
Write-Host ""
Write-Host "현재 상태 확인:"
git status --short --branch | Out-Host
