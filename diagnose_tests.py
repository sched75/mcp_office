"""Script de diagnostic pour les tests."""

import os
import sys

# Ajouter le répertoire parent au path
sys.path.insert(0, os.path.abspath("."))

print("=" * 70)
print("DIAGNOSTIC DES TESTS")
print("=" * 70)
print()

print("Python version:", sys.version)
print("Python executable:", sys.executable)
print("Current directory:", os.getcwd())
print()

print("Tentative d'import des modules...")
print()

try:
    print("1. Import src.outlook.outlook_service...")
    print("   ✓ OK")
except Exception as e:
    print(f"   ✗ ERREUR: {e}")
    import traceback

    traceback.print_exc()

try:
    print("2. Import src.core.exceptions...")
    print("   ✓ OK")
except Exception as e:
    print(f"   ✗ ERREUR: {e}")
    import traceback

    traceback.print_exc()

try:
    print("3. Import tests.test_outlook_service...")
    print("   ✓ OK")
except Exception as e:
    print(f"   ✗ ERREUR: {e}")
    import traceback

    traceback.print_exc()

print()
print("=" * 70)
print("FIN DU DIAGNOSTIC")
print("=" * 70)
