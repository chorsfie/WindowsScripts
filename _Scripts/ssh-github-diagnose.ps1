param(
    [string]$RepoRoot = "C:\Users\Craig\GIT",
    [string]$ExpectedRemote = "git@github.com:chorsfie/WindowsScripts.git"
)

$ErrorActionPreference = "Continue"

function Write-Section($title) {
    Write-Host ""
    Write-Host "=== $title ===" -ForegroundColor Cyan
}

Write-Host "GitHub SSH diagnostics (non-destructive)" -ForegroundColor Green
Write-Host "Repo: $RepoRoot"

Write-Section "1) Git version"
git --version

Write-Section "2) SSH version"
ssh -V

Write-Section "3) SSH key files in ~/.ssh"
$sshPath = Join-Path $HOME ".ssh"
if (Test-Path $sshPath) {
    Get-ChildItem $sshPath -File | Select-Object Name, Length, LastWriteTime
} else {
    Write-Warning "No .ssh folder found at $sshPath"
}

Write-Section "4) Candidate public keys"
$pubKeys = @(
    (Join-Path $sshPath "id_ed25519.pub"),
    (Join-Path $sshPath "id_rsa.pub"),
    (Join-Path $sshPath "id_ecdsa.pub")
) | Where-Object { Test-Path $_ }
if ($pubKeys.Count -eq 0) {
    Write-Warning "No common public key files found."
} else {
    $pubKeys | ForEach-Object { Write-Host "Found: $_" }
}

Write-Section "5) ssh-agent service state"
try {
    Get-Service ssh-agent | Select-Object Name, Status, StartType
} catch {
    Write-Warning "Could not query ssh-agent service: $($_.Exception.Message)"
}

Write-Section "6) Keys loaded in agent"
try {
    ssh-add -l
} catch {
    Write-Warning "ssh-add failed (agent may not be running): $($_.Exception.Message)"
}

Write-Section "7) SSH config for github.com (if present)"
$sshConfig = Join-Path $sshPath "config"
if (Test-Path $sshConfig) {
    Get-Content $sshConfig
} else {
    Write-Host "No SSH config file found at $sshConfig"
}

Write-Section "8) Repo remote check"
if (Test-Path $RepoRoot) {
    git -C $RepoRoot --no-pager remote -v
    $originUrl = (git -C $RepoRoot remote get-url origin 2>$null)
    if ($originUrl) {
        if ($originUrl -eq $ExpectedRemote) {
            Write-Host "Origin URL matches expected SSH remote." -ForegroundColor Green
        } else {
            Write-Warning "Origin URL differs from expected."
            Write-Host "Expected: $ExpectedRemote"
            Write-Host "Actual:   $originUrl"
        }
    } else {
        Write-Warning "No origin remote found."
    }
} else {
    Write-Warning "Repo path not found: $RepoRoot"
}

Write-Section "9) GitHub auth test"
Write-Host "Running: ssh -T git@github.com"
& $env:ComSpec /c "ssh -T git@github.com"

Write-Section "10) Verbose key negotiation test"
Write-Host "Running: ssh -vT git@github.com"
& $env:ComSpec /c "ssh -vT git@github.com"

Write-Section "11) Suggested next actions"
Write-Host "If auth failed:"
Write-Host "  a) Create key: ssh-keygen -t ed25519 -C `"your_email@example.com`""
Write-Host "  b) Start agent: Get-Service ssh-agent | Set-Service -StartupType Automatic; Start-Service ssh-agent"
Write-Host "  c) Add key:     ssh-add $HOME\.ssh\id_ed25519"
Write-Host "  d) Add public key to GitHub: Settings -> SSH and GPG keys"
Write-Host "  e) Retest:      ssh -T git@github.com"
Write-Host "  f) Push:        git -C $RepoRoot push origin main"

Write-Host ""
Write-Host "Diagnostics complete." -ForegroundColor Green
