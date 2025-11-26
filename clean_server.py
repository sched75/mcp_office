"""Script pour nettoyer server.py en supprimant OUTLOOK_TOOLS_CONFIG."""

import re

# Lire server.py
with open("src/server.py", encoding="utf-8") as f:
    content = f.read()

# 1. Modifier l'import pour inclure OUTLOOK_TOOLS_CONFIG
old_import = """from tools_configs import (
    WORD_TOOLS_CONFIG,
    EXCEL_TOOLS_CONFIG,
    POWERPOINT_TOOLS_CONFIG,
)"""

new_import = """from tools_configs import (
    WORD_TOOLS_CONFIG,
    EXCEL_TOOLS_CONFIG,
    POWERPOINT_TOOLS_CONFIG,
    OUTLOOK_TOOLS_CONFIG,
)"""

content = content.replace(old_import, new_import)

# 2. Trouver et supprimer la section OUTLOOK_TOOLS_CONFIG
# Trouver le début
start_pattern = r"# =+\n# OUTLOOK TOOLS CONFIGURATION.*?\n# =+\n\nOUTLOOK_TOOLS_CONFIG = \{"

start_match = re.search(start_pattern, content, re.DOTALL)
if start_match:
    start_pos = start_match.start()

    # Trouver la fin du dictionnaire
    brace_count = 0
    in_dict = False
    end_pos = start_pos

    # Commencer à compter à partir du premier '{'
    first_brace_pos = content.find("{", start_match.start())

    for i, char in enumerate(content[first_brace_pos:], start=first_brace_pos):
        if char == "{":
            brace_count += 1
            in_dict = True
        elif char == "}":
            brace_count -= 1
            if brace_count == 0 and in_dict:
                end_pos = i + 1
                break

    # Supprimer toute la section (y compris les commentaires)
    # Trouver la fin de ligne après la fermeture de l'accolade
    next_newline = content.find("\n", end_pos)
    if next_newline != -1:
        end_pos = next_newline + 1

    # Supprimer la section
    new_content = content[:start_pos] + content[end_pos:]

    # Sauvegarder
    with open("src/server.py", "w", encoding="utf-8") as f:
        f.write(new_content)

    print("✅ server.py modifié avec succès")
    print("   Import mis à jour: OUTLOOK_TOOLS_CONFIG ajouté")
    print(f"   Définition locale supprimée: {end_pos - start_pos} caractères")

    # Vérifier
    with open("src/server.py", encoding="utf-8") as f:
        verify = f.read()

    has_import = "OUTLOOK_TOOLS_CONFIG," in verify
    has_definition = "OUTLOOK_TOOLS_CONFIG = {" in verify

    if has_import and not has_definition:
        print("✅ Vérification OK:")
        print("   - Import présent: ✅")
        print("   - Définition locale absente: ✅")
    else:
        print("⚠️  Vérification:")
        print(f"   - Import présent: {'✅' if has_import else '❌'}")
        print(f"   - Définition locale absente: {'✅' if not has_definition else '❌'}")
else:
    print("❌ Pattern OUTLOOK_TOOLS_CONFIG non trouvé dans server.py")
