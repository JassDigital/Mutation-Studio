$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
$Remote = 'https://github.com/JassDigital/Mutation-Studio.git'

if (-not (Get-Command git -ErrorAction SilentlyContinue)) { throw 'Git is not installed or not on PATH.' }
if (-not (Get-Command gh -ErrorAction SilentlyContinue)) { throw 'GitHub CLI is required. Install it from https://cli.github.com/ and run gh auth login.' }
gh auth status
if ($LASTEXITCODE -ne 0) { throw 'GitHub CLI is not authenticated. Run gh auth login and retry.' }

Write-Host 'Target repository: https://github.com/JassDigital/Mutation-Studio' -ForegroundColor Cyan
Write-Host 'The included sample data is synthetic. Check that no private data has been added before publishing.' -ForegroundColor Yellow
if (-not (Test-Path '.git')) { git init }
git branch -M main
git add -A
if (-not (git rev-parse --verify HEAD 2>$null)) {
    git commit -m 'Initial release: Mutation Studio v1.0'
} else {
    git diff --cached --quiet
    if ($LASTEXITCODE -ne 0) { git commit -m 'Update Mutation Studio' }
}
$existing = git remote get-url origin 2>$null
if ($LASTEXITCODE -eq 0) { git remote set-url origin $Remote } else { git remote add origin $Remote }
git push -u origin main
if ($LASTEXITCODE -ne 0) { throw 'Push failed. Confirm the repository is empty or reconcile existing remote history, then retry.' }
Write-Host 'Published successfully: https://github.com/JassDigital/Mutation-Studio' -ForegroundColor Green
