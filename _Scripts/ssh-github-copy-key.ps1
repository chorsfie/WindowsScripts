param(
    [string]$PublicKeyPath = "$HOME\.ssh\id_ed25519_github.pub"
)

$ErrorActionPreference = "Stop"

Write-Host "GitHub SSH key helper" -ForegroundColor Green
Write-Host "Public key path: $PublicKeyPath"

if (-not (Test-Path $PublicKeyPath)) {
    throw "Public key not found: $PublicKeyPath"
}

$publicKey = Get-Content -Raw -Path $PublicKeyPath
if ([string]::IsNullOrWhiteSpace($publicKey)) {
    throw "Public key file is empty: $PublicKeyPath"
}

try {
    Set-Clipboard -Value $publicKey
    Write-Host "Public key copied to clipboard." -ForegroundColor Green
} catch {
    Write-Warning "Could not copy to clipboard automatically: $($_.Exception.Message)"
}

Write-Host ""
Write-Host "Now add this key to GitHub:"
Write-Host "https://github.com/settings/keys"
Write-Host ""
Write-Host "If clipboard did not work, key content is below:"
Write-Host $publicKey
Write-Host ""
Write-Host "Then test:"
Write-Host "  ssh -T git@github.com"
Write-Host "And push:"
Write-Host "  git -C C:\Users\Craig\GIT push origin main"
