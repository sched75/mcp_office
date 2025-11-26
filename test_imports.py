"""Test simple de couverture."""

import os
import sys

sys.path.insert(0, os.path.abspath("."))

# Test d'import basique
print("Test 1: Import OutlookService")
try:
    from src.outlook.outlook_service import OutlookService

    print("✓ Import réussi")

    # Vérifier les méthodes
    print("\nMéthodes disponibles:")
    methods = [
        m
        for m in dir(OutlookService)
        if not m.startswith("_") and callable(getattr(OutlookService, m))
    ]
    print(f"Nombre de méthodes publiques: {len(methods)}")

    # Compter les méthodes par catégorie
    email_methods = [
        m
        for m in methods
        if "email" in m.lower()
        or "mail" in m.lower()
        or m
        in ["send_email", "create_new_message", "reply_to_email", "forward_email", "search_emails"]
    ]
    print(f"  - Email methods: {len(email_methods)}")

    attachment_methods = [m for m in methods if "attachment" in m.lower()]
    print(f"  - Attachment methods: {len(attachment_methods)}")

    folder_methods = [m for m in methods if "folder" in m.lower()]
    print(f"  - Folder methods: {len(folder_methods)}")

    calendar_methods = [
        m
        for m in methods
        if "appointment" in m.lower() or "calendar" in m.lower() or "reminder" in m.lower()
    ]
    print(f"  - Calendar methods: {len(calendar_methods)}")

    meeting_methods = [m for m in methods if "meeting" in m.lower()]
    print(f"  - Meeting methods: {len(meeting_methods)}")

    contact_methods = [m for m in methods if "contact" in m.lower()]
    print(f"  - Contact methods: {len(contact_methods)}")

    task_methods = [m for m in methods if "task" in m.lower()]
    print(f"  - Task methods: {len(task_methods)}")

    other_methods = [
        m
        for m in methods
        if m
        not in email_methods
        + attachment_methods
        + folder_methods
        + calendar_methods
        + meeting_methods
        + contact_methods
        + task_methods
    ]
    print(f"  - Other methods: {len(other_methods)}")

    print(f"\nTotal methods: {len(methods)}")

except Exception as e:
    print(f"✗ Erreur: {e}")
    import traceback

    traceback.print_exc()

print("\n" + "=" * 70)
