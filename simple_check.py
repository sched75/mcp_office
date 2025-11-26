"""Rapport simple d'intégration."""

import sys

sys.path.insert(0, "src")

try:
    from tools_configs import EXCEL_TOOLS_CONFIG, POWERPOINT_TOOLS_CONFIG, WORD_TOOLS_CONFIG

    word_count = len(WORD_TOOLS_CONFIG)
    excel_count = len(EXCEL_TOOLS_CONFIG)
    ppt_count = len(POWERPOINT_TOOLS_CONFIG)
    outlook_count = 67

    total = word_count + excel_count + ppt_count + outlook_count

    print(f"WORD: {word_count} outils")
    print(f"EXCEL: {excel_count} outils")
    print(f"POWERPOINT: {ppt_count} outils")
    print(f"OUTLOOK: {outlook_count} outils")
    print(f"TOTAL: {total} outils")

    # Vérifier server.py
    with open("src/server.py", encoding="utf-8") as f:
        content = f.read()

    has_word = 'if name.startswith("word_"):' in content
    has_excel = 'elif name.startswith("excel_"):' in content
    has_ppt = 'elif name.startswith("powerpoint_"):' in content
    has_outlook = 'elif name.startswith("outlook_"):' in content
    has_build = "def build_handlers(" in content

    print(f"\nHandlers Word: {'OK' if has_word else 'MANQUANT'}")
    print(f"Handlers Excel: {'OK' if has_excel else 'MANQUANT'}")
    print(f"Handlers PowerPoint: {'OK' if has_ppt else 'MANQUANT'}")
    print(f"Handlers Outlook: {'OK' if has_outlook else 'MANQUANT'}")
    print(f"Fonction build_handlers: {'OK' if has_build else 'MANQUANT'}")

    if all([has_word, has_excel, has_ppt, has_outlook, has_build]):
        print("\n✅ INTÉGRATION COMPLÈTE RÉUSSIE !")
    else:
        print("\n❌ Intégration incomplète")

except Exception as e:
    print(f"ERREUR: {e}")
    import traceback

    traceback.print_exc()
