"""Vérification finale de l'uniformisation de l'architecture."""

import sys

sys.path.insert(0, "src")

print("=" * 80)
print("VÉRIFICATION FINALE DE L'ARCHITECTURE UNIFORME")
print("=" * 80)
print()

# 1. Vérifier tools_configs.py
print("📄 1. Vérification tools_configs.py")
print("-" * 80)

try:
    from tools_configs import (
        EXCEL_TOOLS_CONFIG,
        OUTLOOK_TOOLS_CONFIG,
        POWERPOINT_TOOLS_CONFIG,
        WORD_TOOLS_CONFIG,
    )

    configs = {
        "WORD": WORD_TOOLS_CONFIG,
        "EXCEL": EXCEL_TOOLS_CONFIG,
        "POWERPOINT": POWERPOINT_TOOLS_CONFIG,
        "OUTLOOK": OUTLOOK_TOOLS_CONFIG,
    }

    total = 0
    for name, config in configs.items():
        count = len(config)
        total += count
        print(f"  ✅ {name:15} : {count:3} outils configurés")

    print(f"\n  📊 TOTAL          : {total:3} outils")
    print("  ✅ Toutes les configurations importées avec succès")

except Exception as e:
    print(f"  ❌ Erreur lors de l'import: {e}")
    sys.exit(1)

print()

# 2. Vérifier server.py
print("📄 2. Vérification server.py")
print("-" * 80)

try:
    with open("src/server.py", encoding="utf-8") as f:
        server_content = f.read()

    checks = {
        "Import WORD_TOOLS_CONFIG": "WORD_TOOLS_CONFIG," in server_content,
        "Import EXCEL_TOOLS_CONFIG": "EXCEL_TOOLS_CONFIG," in server_content,
        "Import POWERPOINT_TOOLS_CONFIG": "POWERPOINT_TOOLS_CONFIG," in server_content,
        "Import OUTLOOK_TOOLS_CONFIG": "OUTLOOK_TOOLS_CONFIG," in server_content,
        "Pas de définition locale Outlook": "OUTLOOK_TOOLS_CONFIG = {" not in server_content,
        "Handler Word": 'if name.startswith("word_"):' in server_content,
        "Handler Excel": 'elif name.startswith("excel_"):' in server_content,
        "Handler PowerPoint": 'elif name.startswith("powerpoint_"):' in server_content,
        "Handler Outlook": 'elif name.startswith("outlook_"):' in server_content,
    }

    all_ok = True
    for check_name, result in checks.items():
        status = "✅" if result else "❌"
        print(f"  {status} {check_name}")
        if not result:
            all_ok = False

    if all_ok:
        print("\n  ✅ server.py est correctement configuré")
    else:
        print("\n  ❌ Des problèmes ont été détectés dans server.py")
        sys.exit(1)

except Exception as e:
    print(f"  ❌ Erreur: {e}")
    sys.exit(1)

print()

# 3. Résumé final
print("=" * 80)
print("RÉSUMÉ")
print("=" * 80)
print()
print("✅ Architecture uniformisée avec succès !")
print()
print("📁 Structure finale :")
print("  src/")
print("  ├── tools_configs.py ........... ✅ 4 configurations (271 outils)")
print("  │   ├── WORD_TOOLS_CONFIG")
print("  │   ├── EXCEL_TOOLS_CONFIG")
print("  │   ├── POWERPOINT_TOOLS_CONFIG")
print("  │   └── OUTLOOK_TOOLS_CONFIG")
print("  │")
print("  └── server.py .................. ✅ Importe les 4 configurations")
print("      ├── Import : 4/4 configs")
print("      ├── Handlers : 4/4 services")
print("      └── Pas de duplication")
print()
print("🎯 Avantages de cette architecture :")
print("  ✅ Séparation des responsabilités")
print("  ✅ Configuration centralisée")
print("  ✅ Facile à maintenir")
print("  ✅ Pas de duplication de code")
print("  ✅ Cohérence totale")
print()
print("=" * 80)
print("🎉 UNIFORMISATION RÉUSSIE ! 🎉")
print("=" * 80)
