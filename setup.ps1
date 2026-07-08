# setup.ps1 — 가상환경 생성 및 패키지 설치
# 사용법: .\setup.ps1
# 개발자(HTTP 모드 포함): .\setup.ps1 -Dev

param([switch]$Dev)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host ""
Write-Host "=== royal-mcp 환경 설정 ===" -ForegroundColor Cyan

# Python 설치 상태 확인
$pythonInfo = Get-Command python -ErrorAction SilentlyContinue
if (-not $pythonInfo) {
    Write-Host ""
    Write-Host "❌ Python을 찾을 수 없습니다." -ForegroundColor Red
    Write-Host ""
    Write-Host "   해결 방법:" -ForegroundColor Cyan
    Write-Host "   1. https://www.python.org/ 에서 Python 3.11 이상을 설치하세요." -ForegroundColor White
    Write-Host "      ( https://www.python.org/ftp/python/3.14.5/python-3.14.5-amd64.exe )" -ForegroundColor White
    Write-Host "   2. 설치 중 'Add Python to PATH' 항목을 반드시 체크하세요." -ForegroundColor White
    Write-Host "   3. 설치 완료 후 PowerShell을 새로 열고 .\setup.ps1 을 다시 실행하세요." -ForegroundColor White
    exit 1
} elseif ($pythonInfo.Source -like "*WindowsApps*") {
    Write-Host ""
    Write-Host "❌ Python이 올바르게 설치되지 않았습니다." -ForegroundColor Red
    Write-Host "   현재 'python' 명령이 Microsoft Store로 연결되어 있습니다." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "   해결 방법:" -ForegroundColor Cyan
    Write-Host "   1. https://www.python.org/ 에서 Python 3.11 이상을 설치하세요." -ForegroundColor White
    Write-Host "      ( https://www.python.org/ftp/python/3.14.5/python-3.14.5-amd64.exe )" -ForegroundColor White
    Write-Host "   2. 설치 중 'Add Python to PATH' 항목을 반드시 체크하세요." -ForegroundColor White
    Write-Host "   3. 설치 완료 후 PowerShell을 새로 열고 .\setup.ps1 을 다시 실행하세요." -ForegroundColor White
    exit 1
}

# 1. 가상환경 생성
if (Test-Path "$scriptDir\.venv") {
    Write-Host "[1/3] .venv 이미 존재 — 재사용합니다." -ForegroundColor Yellow
} else {
    Write-Host "[1/3] 가상환경 생성 중..." -ForegroundColor Cyan
    python -m venv "$scriptDir\.venv"
    Write-Host "      완료" -ForegroundColor Green
}

$pip    = "$scriptDir\.venv\Scripts\pip.exe"
$python = "$scriptDir\.venv\Scripts\python.exe"

# 2. pip 업그레이드
Write-Host "[2/3] pip 업그레이드 중..." -ForegroundColor Cyan
& $python -m pip install --upgrade pip --quiet
Write-Host "      완료" -ForegroundColor Green

# 3. 패키지 설치
if ($Dev) {
    Write-Host "[3/3] 패키지 설치 중 (개발자 모드)..." -ForegroundColor Cyan
    & $pip install -r "$scriptDir\requirements-dev.txt"
} else {
    Write-Host "[3/3] 패키지 설치 중..." -ForegroundColor Cyan
    & $pip install -r "$scriptDir\requirements.txt"
}
Write-Host "      완료" -ForegroundColor Green

Write-Host ""
Write-Host "✅ 설치 완료!" -ForegroundColor Green
Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host " Claude Code MCP 등록 안내" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host ""
Write-Host "⚠  TENANT_ID · CLIENT_ID 는 정보관리팀에 문의하여 발급받으세요." -ForegroundColor Yellow
Write-Host ""
Write-Host " 환경변수(--env) 설명:" -ForegroundColor Cyan
Write-Host "  MCP_SERVER_NAME         MCP 서버 표시 이름 (기본값 그대로 사용)" -ForegroundColor DarkGray
Write-Host "  TENANT_ID               Azure AD 테넌트 ID" -ForegroundColor DarkGray
Write-Host "  CLIENT_ID               Azure AD 앱 클라이언트 ID" -ForegroundColor DarkGray
Write-Host "  MS365_TOKEN_CACHE_PATH  토큰 캐시 파일 경로 (project scope 사용 시 권장, 생략 시 기본값)" -ForegroundColor DarkGray
Write-Host "  MS365_AUTO_AUTH         true: 서버 시작 시 자동 로그인 시도 / false: 수동 로그인(ms365 로그인해줘)" -ForegroundColor DarkGray
Write-Host ""
Write-Host "[1단계] 기존 등록 제거  (처음 설치라면 건너뜁니다)" -ForegroundColor Cyan
Write-Host "  claude mcp remove royal-mcp --scope project" -ForegroundColor DarkGray
Write-Host ""
Write-Host "[2단계 — project] 프로젝트 폴더에서 실행  (프로젝트 전용 · 기본 권장)" -ForegroundColor Cyan
Write-Host ""
Write-Host "  claude mcp add --scope project royal-mcp ``" -ForegroundColor Yellow
Write-Host "    `"$python`" `"$scriptDir\server.py`" ``" -ForegroundColor Yellow
Write-Host "    --env `"MCP_SERVER_NAME=royal-mcp`" ``" -ForegroundColor Yellow
Write-Host "    --env `"TENANT_ID=여기에_테넌트ID_입력`" ``" -ForegroundColor Yellow
Write-Host "    --env `"CLIENT_ID=여기에_클라이언트ID_입력`" ``" -ForegroundColor Yellow
Write-Host "    --env `"MS365_TOKEN_CACHE_PATH=`$(Resolve-Path .)\.royal-mcp\token_cache.bin`" ``" -ForegroundColor Yellow
Write-Host "    --env `"MS365_AUTO_AUTH=false`"" -ForegroundColor Yellow


Write-Host ""
Write-Host "[2단계 — user]  어느 폴더에서나 실행 가능  (모든 프로젝트 공용)" -ForegroundColor Cyan
Write-Host ""
Write-Host "  claude mcp add --scope user royal-mcp ``" -ForegroundColor Yellow
Write-Host "    `"$python`" `"$scriptDir\server.py`" ``" -ForegroundColor Yellow
Write-Host "    --env `"MCP_SERVER_NAME=royal-mcp`" ``" -ForegroundColor Yellow
Write-Host "    --env `"TENANT_ID=여기에_테넌트ID_입력`" ``" -ForegroundColor Yellow
Write-Host "    --env `"CLIENT_ID=여기에_클라이언트ID_입력`" ``" -ForegroundColor Yellow
Write-Host "    --env `"MS365_AUTO_AUTH=false`"" -ForegroundColor Yellow
Write-Host "  (MS365_TOKEN_CACHE_PATH 생략 시 기본값: ~\.royal-mcp\token_cache.bin)" -ForegroundColor DarkGray
Write-Host ""
