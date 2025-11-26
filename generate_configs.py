"""Générateur automatique de configurations pour server.py."""

import json

# Charger les données
with open("services_methods.json", encoding="utf-8") as f:
    data = json.load(f)


def generate_tool_config(service_name, methods):
    """Génère la configuration des outils pour un service."""
    config = {}

    for method in methods:
        required = [p["name"] for p in method["params"] if not p["optional"]]
        optional = [p["name"] for p in method["params"] if p["optional"]]

        config[method["name"]] = {
            "required": required,
            "optional": optional,
            "desc": method["docstring"] or f"{method['name'].replace('_', ' ').title()}",
        }

    return config


# Générer les configurations
word_config = generate_tool_config("word", data["word"])
excel_config = generate_tool_config("excel", data["excel"])
powerpoint_config = generate_tool_config("powerpoint", data["powerpoint"])

# Sauvegarder dans un fichier Python
output = f'''"""
Configurations des outils MCP pour Word, Excel et PowerPoint.
Généré automatiquement.
"""

WORD_TOOLS_CONFIG = {json.dumps(word_config, indent=4, ensure_ascii=False)}

EXCEL_TOOLS_CONFIG = {json.dumps(excel_config, indent=4, ensure_ascii=False)}

POWERPOINT_TOOLS_CONFIG = {json.dumps(powerpoint_config, indent=4, ensure_ascii=False)}
'''

with open("tools_configs.py", "w", encoding="utf-8") as f:
    f.write(output)

print("✅ Configurations générées dans tools_configs.py")
print("\nRésumé:")
print(f"  Word: {len(word_config)} outils")
print(f"  Excel: {len(excel_config)} outils")
print(f"  PowerPoint: {len(powerpoint_config)} outils")
print(f"  TOTAL: {len(word_config) + len(excel_config) + len(powerpoint_config)} outils")
