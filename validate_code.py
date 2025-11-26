"""Script de validation complète du code MCP Office."""

import subprocess
import sys

# Couleurs pour le terminal
GREEN = "\033[92m"
RED = "\033[91m"
YELLOW = "\033[93m"
BLUE = "\033[94m"
RESET = "\033[0m"


def print_header(text):
    """Afficher un en-tête."""
    print(f"\n{BLUE}{'=' * 70}{RESET}")
    print(f"{BLUE}{text:^70}{RESET}")
    print(f"{BLUE}{'=' * 70}{RESET}\n")


def run_command(cmd, description):
    """Exécuter une commande et afficher le résultat."""
    print(f"{YELLOW}▶ {description}...{RESET}")
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, check=False, shell=True)

        if result.returncode == 0:
            print(f"{GREEN}✓ {description} - OK{RESET}")
            if result.stdout:
                print(result.stdout)
            return True
        else:
            print(f"{RED}✗ {description} - ERREURS DÉTECTÉES{RESET}")
            if result.stdout:
                print(result.stdout)
            if result.stderr:
                print(result.stderr)
            return False
    except Exception as e:
        print(f"{RED}✗ Erreur lors de l'exécution: {e}{RESET}")
        return False


def main():
    """Fonction principale."""
    print_header("VALIDATION COMPLÈTE DU CODE MCP OFFICE")

    results = []

    # 1. Ruff Check
    print_header("1. VÉRIFICATION RUFF (PEP 8)")
    results.append(
        run_command(
            r".\venv\Scripts\ruff.exe check src/outlook/ tests/test_outlook_service.py --output-format=text",
            "Ruff check - Outlook",
        )
    )

    # 2. Ruff Format Check
    print_header("2. VÉRIFICATION DU FORMATAGE")
    results.append(
        run_command(
            r".\venv\Scripts\ruff.exe format src/outlook/ --check", "Ruff format check - Outlook"
        )
    )

    # 3. Radon - Complexité cyclomatique
    print_header("3. COMPLEXITÉ CYCLOMATIQUE (RADON)")
    results.append(
        run_command(r".\venv\Scripts\radon.exe cc src/outlook/ -a -s -n C", "Radon - Complexité")
    )

    # 4. Radon - Maintainability Index
    print_header("4. INDEX DE MAINTENABILITÉ (RADON)")
    results.append(
        run_command(r".\venv\Scripts\radon.exe mi src/outlook/ -s", "Radon - Maintenabilité")
    )

    # 5. Type checking (si mypy est installé)
    print_header("5. STATISTIQUES DU CODE")
    run_command(r".\venv\Scripts\radon.exe raw src/outlook/ -s", "Statistiques du code")

    # Résumé final
    print_header("RÉSUMÉ DE LA VALIDATION")

    passed = sum(results)
    total = len(results)

    print(f"Tests réussis: {passed}/{total}")

    if passed == total:
        print(f"\n{GREEN}✓ VALIDATION COMPLÈTE RÉUSSIE !{RESET}")
        print(f"{GREEN}Le code est conforme PEP 8 et de bonne qualité.{RESET}\n")
        return 0
    else:
        print(f"\n{YELLOW}⚠ Certaines vérifications nécessitent attention{RESET}")
        print(f"{YELLOW}Consultez les détails ci-dessus.{RESET}\n")
        return 1


if __name__ == "__main__":
    sys.exit(main())
