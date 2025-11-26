#!/usr/bin/env python3
"""Script pour commiter automatiquement les changements Outlook."""

import os
import subprocess
import sys

# Changer le répertoire de travail
project_dir = r"C:\Users\dsi\OneDrive\Documents\Personnel\mcp_office"
os.chdir(project_dir)

print("=" * 60)
print("COMMIT AUTOMATIQUE DES CHANGEMENTS OUTLOOK")
print("=" * 60)
print()

try:
    # Vérifier le statut
    print("📋 Statut actuel:")
    result = subprocess.run(
        ["git", "status", "--short"], capture_output=True, text=True, check=True
    )
    print(result.stdout)

    # Ajouter les fichiers
    print("➕ Ajout des fichiers...")
    files_to_add = [
        "src/outlook/",
        "src/core/types.py",
        "src/core/exceptions.py",
        "tests/test_outlook_service.py",
        "TODO.md",
    ]

    for file in files_to_add:
        subprocess.run(["git", "add", file], check=True)
        print(f"   ✓ {file}")

    print()

    # Message de commit
    commit_message = """feat(outlook): Add complete Outlook service with 85 functionalities

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
- tests/test_outlook_service.py"""

    # Créer le commit
    print("💾 Création du commit...")
    subprocess.run(["git", "commit", "-m", commit_message], check=True)
    print("   ✓ Commit créé avec succès!")
    print()

    # Afficher le dernier commit
    print("📝 Détails du commit:")
    result = subprocess.run(
        ["git", "log", "-1", "--oneline"], capture_output=True, text=True, check=True
    )
    print(result.stdout)

    # Proposer le push
    print("=" * 60)
    print("✅ COMMIT RÉUSSI!")
    print("=" * 60)
    print()
    print("Pour pusher vers GitHub, exécutez:")
    print("   git push origin main")
    print()

    sys.exit(0)

except subprocess.CalledProcessError as e:
    print(f"\n❌ Erreur lors de l'exécution de git: {e}")
    sys.exit(1)
except Exception as e:
    print(f"\n❌ Erreur inattendue: {e}")
    sys.exit(1)
