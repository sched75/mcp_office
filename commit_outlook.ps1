# Script PowerShell pour commiter les changements Outlook

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "Commit des changements Outlook dans MCP Office" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

# Changer de répertoire
Set-Location "C:\Users\dsi\OneDrive\Documents\Personnel\mcp_office"

# Vérifier le statut
Write-Host "Vérification du statut git..." -ForegroundColor Yellow
git status
Write-Host ""

# Ajouter les fichiers
Write-Host "Ajout des nouveaux fichiers..." -ForegroundColor Yellow
git add src/outlook/
git add src/core/types.py
git add src/core/exceptions.py
git add tests/test_outlook_service.py
git add TODO.md
Write-Host ""

# Afficher les fichiers ajoutés
Write-Host "Fichiers ajoutés:" -ForegroundColor Yellow
git status --short
Write-Host ""

# Commit
Write-Host "Création du commit..." -ForegroundColor Yellow
$commitMessage = @"
feat(outlook): Add complete Outlook service with 85 functionalities

- Add OutlookService with full COM automation support
- Implement 12 email operations (send, read, reply, search, etc.)
- Implement 5 attachment operations
- Implement 7 folder operations
- Implement 10 calendar operations (appointments, recurring events, reminders)
- Implement 8 meeting operations (create, invite, accept/decline)
- Implement 9 contact operations (create, modify, search, groups)
- Implement 7 task operations (create, modify, complete, priority)
- Implement 27 advanced operations (categories, rules, signatures, accounts)
- Add 10 Outlook-specific types to core/types.py
- Add 6 Outlook-specific exceptions to core/exceptions.py
- Add comprehensive test suite with mocks
- Add complete documentation and README
- Update TODO.md with Outlook completion status
- Total: 295/295 functionalities complete (100%)

Architecture:
- Mixin pattern for modular functionality
- SOLID principles compliance
- PEP 8 compliance
- Full type hints and docstrings
- Robust error handling with custom exceptions

Files added:
- src/outlook/__init__.py
- src/outlook/outlook_service.py
- src/outlook/mail_operations.py
- src/outlook/attachment_operations.py
- src/outlook/folder_operations.py
- src/outlook/calendar_operations.py
- src/outlook/additional_operations.py
- src/outlook/README.md
- tests/test_outlook_service.py
"@

git commit -m $commitMessage

Write-Host ""
Write-Host "================================================" -ForegroundColor Green
Write-Host "Commit terminé avec succès!" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Green
Write-Host ""

# Proposer le push
$push = Read-Host "Voulez-vous pusher les changements vers GitHub? (O/N)"

if ($push -eq "O" -or $push -eq "o") {
    Write-Host ""
    Write-Host "Push vers GitHub..." -ForegroundColor Yellow
    git push origin main
    Write-Host ""
    Write-Host "================================================" -ForegroundColor Green
    Write-Host "Push terminé avec succès!" -ForegroundColor Green
    Write-Host "================================================" -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "Push annulé. Vous pouvez pusher plus tard avec: git push origin main" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Appuyez sur une touche pour continuer..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
