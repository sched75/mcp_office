"""Analyse toutes les méthodes des services Office pour générer les configurations."""

import ast
import json
from pathlib import Path


def extract_methods_from_file(filepath):
    """Extrait les méthodes publiques avec leurs paramètres."""
    with open(filepath, encoding="utf-8") as f:
        content = f.read()

    try:
        tree = ast.parse(content)
    except Exception:
        return []

    methods = []

    for node in ast.walk(tree):
        if isinstance(node, ast.ClassDef):
            for item in node.body:
                if isinstance(item, ast.FunctionDef) and not item.name.startswith("_"):
                        # Extraire les paramètres
                        params = []
                        defaults_start = len(item.args.args) - len(item.args.defaults)

                        for i, arg in enumerate(item.args.args):
                            if arg.arg == "self":
                                continue

                            # Déterminer si c'est requis ou optionnel
                            is_optional = i >= defaults_start
                            params.append({"name": arg.arg, "optional": is_optional})

                        # Extraire la docstring
                        docstring = ast.get_docstring(item)
                        first_line = ""
                        if docstring:
                            first_line = docstring.split("\n")[0].strip()

                        methods.append(
                            {"name": item.name, "params": params, "docstring": first_line}
                        )

    return methods


# Analyser chaque service
services = {
    "word": "src/word/word_service.py",
    "excel": "src/excel/excel_service.py",
    "powerpoint": "src/powerpoint/powerpoint_service.py",
}

results = {}

for service_name, filepath in services.items():
    path = Path(filepath)
    if path.exists():
        methods = extract_methods_from_file(path)

        # Exclure les méthodes héritées de BaseOfficeService
        base_methods = [
            "initialize",
            "cleanup",
            "create_document",
            "open_document",
            "save_document",
            "close_document",
            "get_application",
        ]

        filtered_methods = [m for m in methods if m["name"] not in base_methods]

        results[service_name] = filtered_methods

        print(f"\n{'=' * 80}")
        print(f"{service_name.upper()} SERVICE - {len(filtered_methods)} méthodes")
        print("=" * 80)

        for method in filtered_methods[:10]:  # Afficher les 10 premières
            required = [p["name"] for p in method["params"] if not p["optional"]]
            optional = [p["name"] for p in method["params"] if p["optional"]]

            print(f"\n📌 {method['name']}")
            if method["docstring"]:
                print(f"   {method['docstring']}")
            if required:
                print(f"   Required: {', '.join(required)}")
            if optional:
                print(f"   Optional: {', '.join(optional)}")

        if len(filtered_methods) > 10:
            print(f"\n   ... et {len(filtered_methods) - 10} autres méthodes")

# Sauvegarder en JSON
with open("services_methods.json", "w", encoding="utf-8") as f:
    json.dump(results, f, indent=2, ensure_ascii=False)

print(f"\n\n{'=' * 80}")
print("RÉSUMÉ")
print("=" * 80)
for service_name, methods in results.items():
    print(f"{service_name.upper()}: {len(methods)} méthodes")

total = sum(len(methods) for methods in results.values())
print(f"\nTOTAL (sans Outlook): {total} méthodes")
print(f"TOTAL AVEC OUTLOOK (67): {total + 67} méthodes")
print("\n✅ Données sauvegardées dans services_methods.json")
