"""Liste toutes les méthodes publiques d'OutlookService pour générer les handlers."""

import inspect
import sys
from pathlib import Path

# Ajouter le projet au path
sys.path.insert(0, str(Path(__file__).parent))

from src.outlook.outlook_service import OutlookService

# Obtenir toutes les méthodes publiques
methods = []
for name in dir(OutlookService):
    if not name.startswith("_") and callable(getattr(OutlookService, name)):
        method = getattr(OutlookService, name)
        # Exclure les méthodes héritées de BaseOfficeService
        if name not in [
            "initialize",
            "cleanup",
            "create_document",
            "open_document",
            "save_document",
            "close_document",
            "namespace",
        ]:
            # Obtenir la signature
            try:
                sig = inspect.signature(method)
                params = list(sig.parameters.keys())
                if "self" in params:
                    params.remove("self")

                methods.append({"name": name, "params": params, "signature": str(sig)})
            except Exception:
                methods.append({"name": name, "params": [], "signature": "(?)"})

# Trier par nom
methods.sort(key=lambda x: x["name"])

print("=" * 80)
print(f"MÉTHODES PUBLIQUES OUTLOOK SERVICE : {len(methods)}")
print("=" * 80)
print()

# Grouper par catégorie (basé sur le préfixe du nom)
categories = {}
for method in methods:
    name = method["name"]

    # Déterminer la catégorie
    if any(
        x in name
        for x in [
            "email",
            "mail",
            "send",
            "reply",
            "forward",
            "read_email",
            "delete_email",
            "mark_as",
            "flag",
            "move_email",
            "search_emails",
        ]
    ):
        cat = "Mail Operations"
    elif "attachment" in name:
        cat = "Attachment Operations"
    elif "folder" in name:
        cat = "Folder Operations"
    elif any(
        x in name
        for x in ["appointment", "calendar", "reminder", "busy_status", "export_appointment"]
    ):
        cat = "Calendar Operations"
    elif "meeting" in name or name in [
        "invite_participants",
        "accept_meeting",
        "decline_meeting",
        "propose_new_time",
        "cancel_meeting",
        "update_meeting",
        "check_availability",
    ]:
        cat = "Meeting Operations"
    elif "contact" in name:
        cat = "Contact Operations"
    elif "task" in name:
        cat = "Task Operations"
    elif any(
        x in name
        for x in [
            "account",
            "category",
            "inbox_count",
            "rule",
            "signature",
            "delegate",
            "out_of_office",
            "free_busy",
        ]
    ):
        cat = "Advanced Operations"
    else:
        cat = "Other Operations"

    if cat not in categories:
        categories[cat] = []
    categories[cat].append(method)

# Afficher par catégorie
for cat_name in sorted(categories.keys()):
    print(f"\n{cat_name}: {len(categories[cat_name])} méthodes")
    print("-" * 80)
    for method in categories[cat_name]:
        params_str = ", ".join(method["params"]) if method["params"] else "()"
        print(f"  - {method['name']}({params_str})")

print()
print("=" * 80)
print(f"TOTAL: {len(methods)} méthodes")
print("=" * 80)

# Sauvegarder dans un fichier
output_file = Path(__file__).parent / "outlook_methods_list.txt"
with open(output_file, "w", encoding="utf-8") as f:
    f.write(f"MÉTHODES PUBLIQUES OUTLOOK SERVICE : {len(methods)}\n")
    f.write("=" * 80 + "\n\n")

    for cat_name in sorted(categories.keys()):
        f.write(f"\n{cat_name}: {len(categories[cat_name])} méthodes\n")
        f.write("-" * 80 + "\n")
        for method in categories[cat_name]:
            params_str = ", ".join(method["params"]) if method["params"] else ""
            f.write(f"  - {method['name']}({params_str})\n")

    f.write(f"\n\nTOTAL: {len(methods)} méthodes\n")

print(f"\n✅ Liste sauvegardée dans {output_file}")
