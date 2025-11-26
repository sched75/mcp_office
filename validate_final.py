"""Script de validation qui écrit dans un fichier."""

import subprocess
import sys
from pathlib import Path

output_file = Path("validation_results.txt")


def write_output(text):
    """Écrire dans le fichier et sur stdout."""
    with open(output_file, "a", encoding="utf-8") as f:
        f.write(text + "\n")
    print(text)


def run_ruff_check():
    """Vérifier PEP 8 avec Ruff."""
    write_output("\n" + "=" * 70)
    write_output("VÉRIFICATION RUFF (PEP 8 COMPLIANCE)")
    write_output("=" * 70 + "\n")

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
        write_output("✓ Aucune erreur PEP 8 détectée!")
        write_output("✓ Code 100% conforme PEP 8\n")
        return True
    else:
        write_output("✗ Erreurs PEP 8 détectées:")
        write_output(result.stdout)
        if result.stderr:
            write_output(result.stderr)
        return False


def run_radon_cc():
    """Vérifier la complexité cyclomatique avec Radon."""
    write_output("\n" + "=" * 70)
    write_output("COMPLEXITÉ CYCLOMATIQUE (RADON)")
    write_output("=" * 70 + "\n")

    cmd = [r".\venv\Scripts\radon.exe", "cc", "src/outlook/", "-a", "-s"]

    result = subprocess.run(cmd, capture_output=True, text=True)

    write_output(result.stdout if result.stdout else "Aucune sortie")

    # Vérifier s'il y a des grades C, D, E, F
    if any(grade in result.stdout for grade in ["(C)", "(D)", "(E)", "(F)"]):
        write_output("\n✗ Des fonctions ont une complexité élevée (C, D, E ou F)")
        return False
    else:
        write_output("\n✓ Toutes les fonctions ont une complexité faible (A ou B)")
        return True


def run_radon_mi():
    """Vérifier l'index de maintenabilité avec Radon."""
    write_output("\n" + "=" * 70)
    write_output("INDEX DE MAINTENABILITÉ (RADON)")
    write_output("=" * 70 + "\n")

    cmd = [r".\venv\Scripts\radon.exe", "mi", "src/outlook/", "-s"]

    result = subprocess.run(cmd, capture_output=True, text=True)

    write_output(result.stdout if result.stdout else "Aucune sortie")

    # Vérifier s'il y a des grades C, D, E, F
    if any(
        grade in result.stdout for grade in ["(C)", "(D)", "(E)", "(F)", "- C", "- D", "- E", "- F"]
    ):
        write_output("\n✗ Certains fichiers ont une faible maintenabilité (C, D, E ou F)")
        return False
    else:
        write_output("\n✓ Tous les fichiers ont une bonne maintenabilité (A ou B)")
        return True


def main():
    """Fonction principale."""
    # Supprimer le fichier s'il existe
    if output_file.exists():
        output_file.unlink()

    write_output("\n" + "=" * 70)
    write_output("VALIDATION COMPLÈTE DU CODE MCP OFFICE - OUTLOOK")
    write_output("=" * 70)

    results = []

    # Test 1: Ruff
    results.append(run_ruff_check())

    # Test 2: Radon CC
    results.append(run_radon_cc())

    # Test 3: Radon MI
    results.append(run_radon_mi())

    # Résumé
    write_output("\n" + "=" * 70)
    write_output("RÉSUMÉ FINAL")
    write_output("=" * 70 + "\n")

    passed = sum(results)
    total = len(results)

    write_output(f"Tests réussis: {passed}/{total}")

    if all(results):
        write_output("\n✓✓✓ TOUS LES TESTS PASSÉS ✓✓✓")
        write_output("✓ Code 100% conforme PEP 8")
        write_output("✓ Complexité cyclomatique: grades A ou B uniquement")
        write_output("✓ Maintenabilité: grades A ou B uniquement")
        write_output("\n🎉 LE CODE EST PRÊT POUR LE COMMIT FINAL ! 🎉\n")
        return 0
    else:
        write_output("\n✗ Certains tests ont échoué")
        write_output("Veuillez corriger les problèmes identifiés ci-dessus.\n")
        return 1


if __name__ == "__main__":
    sys.exit(main())
