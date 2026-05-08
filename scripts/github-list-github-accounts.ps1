<#!
  Git Credential Manager에 저장된 GitHub.com 계정 목록을 출력합니다.
#>
$ErrorActionPreference = "Stop"

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Error "Git이 PATH에 없습니다."
}

Write-Host "저장된 GitHub 계정:" -ForegroundColor Cyan
& git credential-manager github list
