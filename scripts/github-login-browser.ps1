<#!
  GitHub를 웹 브라우저로 로그인합니다 (Git Credential Manager).

  사용:
    .\scripts\github-login-browser.ps1
    .\scripts\github-login-browser.ps1 -Force

  -Force : 이미 저장된 같은 GitHub 계정의 인증을 다시 받을 때(토큰 갱신).

  다른 계정으로 바꾸려면 먼저 scripts\github-list-github-accounts.ps1 로 계정을 확인한 뒤
  scripts\github-logout.ps1 로 기존 계정을 제거하고, 이 스크립트를 다시 실행하세요.
#>
param(
    [switch]$Force
)

$ErrorActionPreference = "Stop"

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Error "Git이 PATH에 없습니다. Git for Windows를 설치한 뒤 터미널을 다시 여세요."
}

$gitArgs = @("credential-manager", "github", "login", "--browser")
if ($Force) {
    $gitArgs += "--force"
}

Write-Host "브라우저에서 GitHub 로그인 창이 열립니다. 완료 후 이 창으로 돌아오세요." -ForegroundColor Cyan
& git @gitArgs
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

Write-Host "`n등록된 GitHub 계정:" -ForegroundColor Green
& git credential-manager github list
