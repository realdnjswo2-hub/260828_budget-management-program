# PowerShell runner for Budget App
$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "========================================================" -ForegroundColor Cyan
Write-Host "  [예산관리대장] 프로그램을 시작합니다..." -ForegroundColor Cyan
Write-Host "========================================================" -ForegroundColor Cyan
Write-Host ""

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
    Write-Host "Python이 설치되어 있는지 확인해주세요." -ForegroundColor Yellow
    Read-Host "엔터 키를 누르면 종료합니다..."
    exit 1
}

Write-Host "[확인] Python 경로: $pythonPath" -ForegroundColor Green

Write-Host "필수 패키지(openpyxl, pypdf) 확인 중..." -ForegroundColor Gray
& $pythonPath -m pip install openpyxl pypdf --break-system-packages --quiet

Write-Host "프로그램 창을 엽니다..." -ForegroundColor Green
& $pythonPath main.py
