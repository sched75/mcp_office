"""Test simple d'import."""

import sys

sys.path.insert(0, "src")

try:
    print("Test d'import...")
    from tools_configs import (
        EXCEL_TOOLS_CONFIG,
        OUTLOOK_TOOLS_CONFIG,
        POWERPOINT_TOOLS_CONFIG,
        WORD_TOOLS_CONFIG,
    )

    print(f"✅ WORD: {len(WORD_TOOLS_CONFIG)} outils")
    print(f"✅ EXCEL: {len(EXCEL_TOOLS_CONFIG)} outils")
    print(f"✅ POWERPOINT: {len(POWERPOINT_TOOLS_CONFIG)} outils")
    print(f"✅ OUTLOOK: {len(OUTLOOK_TOOLS_CONFIG)} outils")
    print(
        f"\n✅ TOTAL: {len(WORD_TOOLS_CONFIG) + len(EXCEL_TOOLS_CONFIG) + len(POWERPOINT_TOOLS_CONFIG) + len(OUTLOOK_TOOLS_CONFIG)} outils"
    )

    print("\n✅ Tous les imports fonctionnent !")

except Exception as e:
    print(f"❌ Erreur: {e}")
    import traceback

    traceback.print_exc()
