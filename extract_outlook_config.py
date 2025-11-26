"""Script pour extraire OUTLOOK_TOOLS_CONFIG de server.py."""

import re

# Lire server.py
with open("src/server.py", encoding="utf-8") as f:
    content = f.read()

# Trouver le début et la fin de OUTLOOK_TOOLS_CONFIG
start_pattern = r"OUTLOOK_TOOLS_CONFIG = \{"
end_pattern = r"\n\}"

# Extraire la configuration
start_match = re.search(start_pattern, content)
if start_match:
    start_pos = start_match.start()

    # Trouver la fin du dictionnaire (compter les accolades)
    brace_count = 0
    in_dict = False
    end_pos = start_pos

    for i, char in enumerate(content[start_pos:], start=start_pos):
        if char == "{":
            brace_count += 1
            in_dict = True
        elif char == "}":
            brace_count -= 1
            if brace_count == 0 and in_dict:
                end_pos = i + 1
                break

    outlook_config = content[start_pos:end_pos]

    # Sauvegarder
    with open("outlook_config_extracted.txt", "w", encoding="utf-8") as f:
        f.write(outlook_config)

    print(f"✅ Configuration Outlook extraite ({len(outlook_config)} caractères)")
    print(f"   Début: position {start_pos}")
    print(f"   Fin: position {end_pos}")

    # Compter les outils
    tool_count = outlook_config.count('    "')
    print(f"   Outils détectés: ~{tool_count}")
else:
    print("❌ OUTLOOK_TOOLS_CONFIG non trouvé")
