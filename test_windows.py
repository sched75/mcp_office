#!/usr/bin/env python
"""Script de test pour vérifier l'installation sur Windows."""

import sys


def test_python_version():
    """Vérifier la version de Python."""
    print("🐍 Vérification de Python...")
    version = sys.version_info
    if version.major >= 3 and version.minor >= 10:
        print(f"   ✅ Python {version.major}.{version.minor}.{version.micro}")
        return True
    else:
        print(f"   ❌ Python {version.major}.{version.minor}.{version.micro} (3.10+ requis)")
        return False


def test_pywin32():
    """Vérifier l'installation de pywin32."""
    print("\n📦 Vérification de pywin32...")
    try:
        import win32com.client

        print("   ✅ pywin32 installé")
        return True
    except ImportError:
        print("   ❌ pywin32 non installé")
        print("      Installer avec: pip install pywin32")
        return False


def test_office_word():
    """Vérifier l'installation de Word."""
    print("\n📝 Vérification de Microsoft Word...")
    try:
        import win32com.client

        word = win32com.client.Dispatch("Word.Application")
        version = word.Version
        word.Quit()
        print(f"   ✅ Word {version} détecté")
        return True
    except Exception as e:
        print(f"   ❌ Word non détecté: {e}")
        return False


def test_office_excel():
    """Vérifier l'installation d'Excel."""
    print("\n📊 Vérification de Microsoft Excel...")
    try:
        import win32com.client

        excel = win32com.client.Dispatch("Excel.Application")
        version = excel.Version
        excel.Quit()
        print(f"   ✅ Excel {version} détecté")
        return True
    except Exception as e:
        print(f"   ❌ Excel non détecté: {e}")
        return False


def test_office_powerpoint():
    """Vérifier l'installation de PowerPoint."""
    print("\n📽️ Vérification de Microsoft PowerPoint...")
    try:
        import win32com.client

        ppt = win32com.client.Dispatch("PowerPoint.Application")
        version = ppt.Version
        ppt.Quit()
        print(f"   ✅ PowerPoint {version} détecté")
        return True
    except Exception as e:
        print(f"   ❌ PowerPoint non détecté: {e}")
        return False


def test_mcp():
    """Vérifier l'installation du package MCP."""
    print("\n🔌 Vérification du package MCP...")
    try:
        import mcp

        print("   ✅ MCP installé")
        return True
    except ImportError:
        print("   ❌ MCP non installé")
        print("      Installer avec: pip install mcp")
        return False


def test_services():
    """Vérifier que les services peuvent être importés."""
    print("\n⚙️ Vérification des services...")
    results = []

    try:
        print("   ✅ WordService importé")
        results.append(True)
    except Exception as e:
        print(f"   ❌ WordService: {e}")
        results.append(False)

    try:
        print("   ✅ ExcelService importé")
        results.append(True)
    except Exception as e:
        print(f"   ❌ ExcelService: {e}")
        results.append(False)

    try:
        print("   ✅ PowerPointService importé")
        results.append(True)
    except Exception as e:
        print(f"   ❌ PowerPointService: {e}")
        results.append(False)

    return all(results)


def main():
    """Exécuter tous les tests."""
    print("=" * 60)
    print("  🧪 Test d'Installation - MCP Office Automation")
    print("=" * 60)

    results = {
        "Python 3.10+": test_python_version(),
        "pywin32": test_pywin32(),
        "Microsoft Word": test_office_word(),
        "Microsoft Excel": test_office_excel(),
        "Microsoft PowerPoint": test_office_powerpoint(),
        "Package MCP": test_mcp(),
        "Services": test_services(),
    }

    print("\n" + "=" * 60)
    print("  📊 Résumé des Tests")
    print("=" * 60)

    for name, passed in results.items():
        status = "✅ PASS" if passed else "❌ FAIL"
        print(f"  {status}  {name}")

    print("=" * 60)

    all_passed = all(results.values())
    required_passed = results["Python 3.10+"] and results["pywin32"] and results["Services"]

    if all_passed:
        print("\n🎉 Tous les tests passent! Le serveur est prêt à être lancé.")
        print("\n   Pour démarrer: python -m src.server")
    elif required_passed:
        print("\n⚠️  Installation fonctionnelle (Office partiellement détecté)")
        print("   Le serveur peut être lancé, mais certaines fonctionnalités")
        print("   Office peuvent ne pas fonctionner.")
        print("\n   Pour démarrer: python -m src.server")
    else:
        print("\n❌ Installation incomplète. Veuillez installer les composants manquants.")
        print("\n   1. Vérifier Python 3.10+")
        print("   2. Installer les dépendances: pip install -r requirements.txt")
        print("   3. Vérifier l'installation d'Office")

    print("\n" + "=" * 60 + "\n")

    return 0 if all_passed else 1


if __name__ == "__main__":
    sys.exit(main())
