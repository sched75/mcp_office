# MCP Office - Script de démarrage du serveur
# Démarre le serveur MCP Office avec logging

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  MCP OFFICE - Démarrage du serveur" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Vérifier que l'environnement virtuel existe
if (-not (Test-Path ".\venv")) {
    Write-Host "❌ Environnement virtuel non trouvé !" -ForegroundColor Red
    Write-Host "Exécutez d'abord : .\scripts\install.ps1" -ForegroundColor Yellow
    exit 1
}

# Activer l'environnement virtuel
Write-Host "Activation de l'environnement virtuel..." -ForegroundColor Yellow
& ".\venv\Scripts\Activate.ps1"

# Définir les variables d'environnement
$env:PYTHONPATH = (Get-Location).Path
$env:PYTHON_UNBUFFERED = "1"

Write-Host "✅ Environnement configuré" -ForegroundColor Green
Write-Host ""
Write-Host "Démarrage du serveur MCP Office..." -ForegroundColor Yellow
Write-Host "295 outils disponibles :" -ForegroundColor Cyan
Write-Host "  • Word (65 outils)" -ForegroundColor White
Write-Host "  • Excel (82 outils)" -ForegroundColor White
Write-Host "  • PowerPoint (63 outils)" -ForegroundColor White
Write-Host "  • Outlook (85 outils)" -ForegroundColor White
Write-Host ""
Write-Host "Le serveur démarre... (Ctrl+C pour arrêter)" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Démarrer le serveur
try {
    python -m src.server
} catch {
    Write-Host ""
    Write-Host "❌ Erreur lors du démarrage du serveur" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "Consultez docs\troubleshooting.md pour l'aide au dépannage" -ForegroundColor Yellow
    exit 1
}
