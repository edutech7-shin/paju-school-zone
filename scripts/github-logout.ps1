<#!
  Git Credential Manager에서 GitHub 계정을 제거합니다 (다른 계정으로 브라우저 로그인 전에 사용).

  사용:
    .\scripts\github-logout.ps1 -Account "이전GitHub아이디"

  예:
    .\scripts\github-logout.ps1 -Account "khamissemrekr"
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$Account
)

$ErrorActionPreference = "Stop"

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Error "Git이 PATH에 없습니다."
}

Write-Host "계정 제거 중: $Account" -ForegroundColor Yellow
& git credential-manager github logout $Account
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

Write-Host "남은 계정:" -ForegroundColor Green
& git credential-manager github list
