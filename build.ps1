# PowerShell builder for Budget App (.exe)
$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "  [예산관리대장.exe] 단독 실행 파일 빌드를 시작합니다..." -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host ""

# 기존 실행 중인 프로세스 종료 시도
Stop-Process -Name "예산관리대장" -Force -ErrorAction SilentlyContinue

$pythonPath = $null

$candidates = @(
    "$env:USERPROFILE\.local\bin\python3.12.exe",
    "$env:USERPROFILE\.local\bin\python3.14.exe",
    "python",
    "py",
    "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe",
    "$env:LOCALAPPDATA\Programs\Python\Python311\python.exe",
    "$env:LOCALAPPDATA\Programs\Python\Python310\python.exe",
    "C:\Python312\python.exe",
    "C:\Python311\python.exe"
)

foreach ($cand in $candidates) {
    try {
        if ($cand -match "\\") {
            if (Test-Path $cand) {
                $pythonPath = $cand
                break
            }
        } else {
            $cmd = Get-Command $cand -ErrorAction SilentlyContinue
            if ($cmd -and $cmd.Source) {
                if ($cmd.Source -notmatch "WindowsApps") {
                    $pythonPath = $cmd.Source
                    break
                }
            }
        }
    } catch {}
}

if (-not $pythonPath) {
    Write-Host "[오류] 사용 가능한 Python을 찾을 수 없습니다." -ForegroundColor Red
    Read-Host "엔터 키를 누르면 종료합니다..."
    exit 1
}

Write-Host "[확인] Python 경로: $pythonPath" -ForegroundColor Green

Write-Host "1. 필수 빌드 패키지(openpyxl, pypdf, pyinstaller) 설치 확인 중..." -ForegroundColor Yellow
& $pythonPath -m pip install openpyxl pypdf pyinstaller --break-system-packages --quiet

Write-Host "2. PyInstaller 빌드 실행 중..." -ForegroundColor Yellow
$currentDir = Get-Location

& $pythonPath -m PyInstaller --noconsole --onefile --clean `
    --name="예산관리대장" `
    --add-data "config;config" `
    --add-data "sample;sample" `
    --add-data "engine;engine" `
    main.py

Write-Host ""
if (Test-Path "dist\예산관리대장.exe") {
    Write-Host "========================================================" -ForegroundColor Green
    Write-Host "  [빌드 성공!]" -ForegroundColor Green
    Write-Host "  생성된 파일: $currentDir\dist\예산관리대장.exe" -ForegroundColor Green
    Write-Host "========================================================" -ForegroundColor Green
} else {
    Write-Host "[오류] 빌드 생성에 실패했습니다." -ForegroundColor Red
}

Write-Host ""
Read-Host "계속하려면 엔터 키를 누르십시오..."
