"""Script de validation pour vérifier Ruff et Radon."""

import subprocess
import sys


def run_ruff_check():
    """Vérifier PEP 8 avec Ruff."""
    print("\n" + "=" * 70)
    print("VÉRIFICATION RUFF (PEP 8 COMPLIANCE)")
    print("=" * 70 + "\n")

    cmd = [
        r".\venv\Scripts\ruff.exe",
        "check",
        "src/outlook/",
        "tests/test_outlook_service.py",
        "src/core/types.py",
        "src/core/exceptions.py",
    ]

    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode == 0:
        print("✓ Aucune erreur PEP 8 détectée!")
        print("✓ Code 100% conforme PEP 8\n")
        return True
    else:
        print("✗ Erreurs PEP 8 détectées:")
        print(result.stdout)
        if result.stderr:
            print(result.stderr)
        return False


def run_radon_cc():
    """Vérifier la complexité cyclomatique avec Radon."""
    print("\n" + "=" * 70)
    print("COMPLEXITÉ CYCLOMATIQUE (RADON)")
    print("=" * 70 + "\n")

    cmd = [r".\venv\Scripts\radon.exe", "cc", "src/outlook/", "-a", "-s"]

    result = subprocess.run(cmd, capture_output=True, text=True)

    print(result.stdout)

    # Vérifier s'il y a des grades C, D, E, F
    if any(grade in result.stdout for grade in [" (C)", " (D)", " (E)", " (F)"]):
        print("\n✗ Des fonctions ont une complexité élevée (C, D, E ou F)")
        return False
    else:
        print("\n✓ Toutes les fonctions ont une complexité faible (A ou B)")
        return True


def run_radon_mi():
    """Vérifier l'index de maintenabilité avec Radon."""
    print("\n" + "=" * 70)
    print("INDEX DE MAINTENABILITÉ (RADON)")
    print("=" * 70 + "\n")

    cmd = [r".\venv\Scripts\radon.exe", "mi", "src/outlook/", "-s"]

    result = subprocess.run(cmd, capture_output=True, text=True)

    print(result.stdout)

    # Vérifier s'il y a des grades C, D, E, F
    if any(
        grade in result.stdout
        for grade in [" (C)", " (D)", " (E)", " (F)", " - C", " - D", " - E", " - F"]
    ):
        print("\n✗ Certains fichiers ont une faible maintenabilité (C, D, E ou F)")
        return False
    else:
        print("\n✓ Tous les fichiers ont une bonne maintenabilité (A ou B)")
        return True


def main():
    """Fonction principale."""
    print("\n" + "=" * 70)
    print("VALIDATION COMPLÈTE DU CODE MCP OFFICE - OUTLOOK")
    print("=" * 70)

    results = []

    # Test 1: Ruff
    results.append(run_ruff_check())

    # Test 2: Radon CC
    results.append(run_radon_cc())

    # Test 3: Radon MI
    results.append(run_radon_mi())

    # Résumé
    print("\n" + "=" * 70)
    print("RÉSUMÉ FINAL")
    print("=" * 70 + "\n")

    passed = sum(results)
    total = len(results)

    print(f"Tests réussis: {passed}/{total}")

    if all(results):
        print("\n✓✓✓ TOUS LES TESTS PASSÉS ✓✓✓")
        print("✓ Code 100% conforme PEP 8")
        print("✓ Complexité cyclomatique: grades A ou B uniquement")
        print("✓ Maintenabilité: grades A ou B uniquement")
        print("\n🎉 LE CODE EST PRÊT POUR LE COMMIT FINAL ! 🎉\n")
        return 0
    else:
        print("\n✗ Certains tests ont échoué")
        print("Veuillez corriger les problèmes identifiés ci-dessus.\n")
        return 1


if __name__ == "__main__":
    sys.exit(main())
