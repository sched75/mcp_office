# MCP Office - Script d'installation automatique
# Ce script configure l'environnement complet pour le serveur MCP Office

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  MCP OFFICE - Installation Automatique" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 1. Vérifier Python
Write-Host "[1/6] Vérification de Python..." -ForegroundColor Yellow
$pythonVersion = python --version 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ Python n'est pas installé !" -ForegroundColor Red
    Write-Host "Veuillez installer Python 3.8+ depuis https://www.python.org/" -ForegroundColor Red
    exit 1
}
Write-Host "✅ $pythonVersion" -ForegroundColor Green

# 2. Créer l'environnement virtuel
Write-Host "[2/6] Création de l'environnement virtuel..." -ForegroundColor Yellow
if (Test-Path ".\venv") {
    Write-Host "⚠️  L'environnement virtuel existe déjà" -ForegroundColor Yellow
} else {
    python -m venv venv
    Write-Host "✅ Environnement virtuel créé" -ForegroundColor Green
}

# 3. Activer l'environnement virtuel et installer les dépendances
Write-Host "[3/6] Installation des dépendances..." -ForegroundColor Yellow
& ".\venv\Scripts\Activate.ps1"
pip install --upgrade pip
pip install -r requirements.txt
Write-Host "✅ Dépendances installées" -ForegroundColor Green

# 4. Vérifier Microsoft Office
Write-Host "[4/6] Vérification de Microsoft Office..." -ForegroundColor Yellow
$officeApps = @()
$apps = @{
    "Word" = "HKLM:\SOFTWARE\Microsoft\Office\*\Word\InstallRoot"
    "Excel" = "HKLM:\SOFTWARE\Microsoft\Office\*\Excel\InstallRoot"
    "PowerPoint" = "HKLM:\SOFTWARE\Microsoft\Office\*\PowerPoint\InstallRoot"
    "Outlook" = "HKLM:\SOFTWARE\Microsoft\Office\*\Outlook\InstallRoot"
}

foreach ($app in $apps.Keys) {
    if (Test-Path $apps[$app]) {
        $officeApps += $app
        Write-Host "  ✅ $app détecté" -ForegroundColor Green
    } else {
        Write-Host "  ⚠️  $app non détecté" -ForegroundColor Yellow
    }
}

if ($officeApps.Count -eq 0) {
    Write-Host "❌ Aucune application Office détectée !" -ForegroundColor Red
    Write-Host "Le serveur MCP nécessite Microsoft Office installé" -ForegroundColor Red
    exit 1
}

# 5. Configurer Claude Desktop
Write-Host "[5/6] Configuration de Claude Desktop..." -ForegroundColor Yellow
$claudeConfigPath = "$env:APPDATA\Claude\claude_desktop_config.json"
$claudeConfigDir = Split-Path $claudeConfigPath

if (-not (Test-Path $claudeConfigDir)) {
    Write-Host "⚠️  Répertoire de configuration Claude Desktop introuvable" -ForegroundColor Yellow
    Write-Host "Configuration manuelle requise - voir docs/installation.md" -ForegroundColor Yellow
} else {
    $currentDir = (Get-Location).Path
    $configTemplate = Get-Content ".\config\claude_desktop_config.json" -Raw
    $configTemplate = $configTemplate -replace "C:\\\\Users\\\\dsi\\\\OneDrive\\\\Documents\\\\Personnel\\\\mcp_office", $currentDir.Replace('\', '\\')
    
    if (Test-Path $claudeConfigPath) {
        Write-Host "⚠️  Configuration Claude Desktop existante détectée" -ForegroundColor Yellow
        $response = Read-Host "Voulez-vous ajouter/fusionner la configuration MCP Office ? (O/N)"
        if ($response -eq 'O' -or $response -eq 'o') {
            $existingConfig = Get-Content $claudeConfigPath -Raw | ConvertFrom-Json
            $newConfig = $configTemplate | ConvertFrom-Json
            
            if (-not $existingConfig.mcpServers) {
                $existingConfig | Add-Member -NotePropertyName "mcpServers" -NotePropertyValue @{}
            }
            
            $existingConfig.mcpServers.'mcp-office' = $newConfig.mcpServers.'mcp-office'
            $existingConfig | ConvertTo-Json -Depth 10 | Set-Content $claudeConfigPath
            Write-Host "✅ Configuration fusionnée" -ForegroundColor Green
        }
    } else {
        $configTemplate | Set-Content $claudeConfigPath
        Write-Host "✅ Configuration Claude Desktop créée" -ForegroundColor Green
    }
}

# 6. Test du serveur
Write-Host "[6/6] Test du serveur MCP..." -ForegroundColor Yellow
Write-Host "Tentative de démarrage du serveur (appuyez sur Ctrl+C pour arrêter)" -ForegroundColor Yellow
Start-Sleep -Seconds 2

# Le serveur va démarrer en mode interactif - l'utilisateur devra le tester manuellement
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Installation terminée !" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Applications Office détectées:" -ForegroundColor Green
foreach ($app in $officeApps) {
    Write-Host "  • $app" -ForegroundColor Green
}
Write-Host ""
Write-Host "Prochaines étapes:" -ForegroundColor Yellow
Write-Host "1. Redémarrez Claude Desktop" -ForegroundColor White
Write-Host "2. Vérifiez que 'mcp-office' apparaît dans les serveurs disponibles" -ForegroundColor White
Write-Host "3. Testez avec une commande simple : 'Crée un document Word avec le texte Bonjour'" -ForegroundColor White
Write-Host ""
Write-Host "Pour démarrer le serveur manuellement:" -ForegroundColor Yellow
Write-Host "  .\scripts\start_server.ps1" -ForegroundColor White
Write-Host ""
Write-Host "Pour la documentation complète:" -ForegroundColor Yellow
Write-Host "  docs\installation.md" -ForegroundColor White
Write-Host "  docs\user_guide.md" -ForegroundColor White
Write-Host ""
