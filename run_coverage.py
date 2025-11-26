"""Script pour mesurer la couverture des tests Outlook."""

import subprocess
import sys

print("=" * 70)
print("MESURE DE COUVERTURE DES TESTS OUTLOOK")
print("=" * 70)
print()

# Exécuter pytest avec coverage
cmd = [
    r".\venv\Scripts\pytest.exe",
    "tests/test_outlook_service.py",
    "--cov=src.outlook",
    "--cov-report=term-missing",
    "--cov-report=html",
    "-v",
]

print("Exécution des tests avec coverage...")
print()

result = subprocess.run(cmd, capture_output=True, text=True)

print(result.stdout)
if result.stderr:
    print("STDERR:")
    print(result.stderr)

# Écrire dans un fichier
with open("coverage_report.txt", "w", encoding="utf-8") as f:
    f.write("=" * 70 + "\n")
    f.write("RAPPORT DE COUVERTURE DES TESTS OUTLOOK\n")
    f.write("=" * 70 + "\n\n")
    f.write(result.stdout)
    if result.stderr:
        f.write("\n\nSTDERR:\n")
        f.write(result.stderr)

print()
print("=" * 70)
print("Rapport sauvegardé dans coverage_report.txt")
print("Rapport HTML disponible dans htmlcov/index.html")
print("=" * 70)

sys.exit(result.returncode)
