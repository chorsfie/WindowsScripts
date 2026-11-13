param(
    [string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path,
    [string]$Branch = "main"
)

$ErrorActionPreference = "Stop"

Write-Host "=== Git push readiness check ==="
Write-Host "Repo: $RepoRoot"
Write-Host "Branch: $Branch"

if (-not (Test-Path $RepoRoot)) {
    throw "Repo path not found: $RepoRoot"
}

Write-Host ""
Write-Host "[1/5] Current branch"
git -C $RepoRoot --no-pager branch --show-current

Write-Host ""
Write-Host "[2/5] Remotes"
git -C $RepoRoot --no-pager remote -v

Write-Host ""
Write-Host "[3/5] SSH auth test to GitHub"
Write-Host "Command: ssh -T git@github.com"
try {
    & $env:ComSpec /c "ssh -T git@github.com"
} catch {
    Write-Warning "SSH auth test returned an error. Review key setup if push fails."
}

Write-Host ""
Write-Host "[4/5] Working tree status"
git -C $RepoRoot --no-pager status --short

Write-Host ""
Write-Host "[5/5] Dry-run push"
git -C $RepoRoot --no-pager push --dry-run origin $Branch

Write-Host ""
Write-Host "Push readiness check complete."
