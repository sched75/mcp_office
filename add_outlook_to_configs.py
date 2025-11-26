"""Script pour ajouter OUTLOOK_TOOLS_CONFIG à tools_configs.py."""

# Lire le fichier actuel
with open("src/tools_configs.py", encoding="utf-8") as f:
    current_content = f.read()

# Lire la config Outlook extraite
with open("outlook_config_extracted.txt", encoding="utf-8") as f:
    outlook_config = f.read()

# Vérifier si OUTLOOK_TOOLS_CONFIG est déjà présent
if "OUTLOOK_TOOLS_CONFIG" in current_content:
    print("⚠️  OUTLOOK_TOOLS_CONFIG déjà présent dans tools_configs.py")
else:
    # Ajouter OUTLOOK_TOOLS_CONFIG à la fin
    new_content = current_content.rstrip() + "\n\n" + outlook_config + "\n"

    # Sauvegarder
    with open("src/tools_configs.py", "w", encoding="utf-8") as f:
        f.write(new_content)

    print("✅ OUTLOOK_TOOLS_CONFIG ajouté à tools_configs.py")
    print(f"   Taille ajoutée: {len(outlook_config)} caractères")

    # Vérifier
    with open("src/tools_configs.py", encoding="utf-8") as f:
        verify_content = f.read()

    if "OUTLOOK_TOOLS_CONFIG" in verify_content:
        print("✅ Vérification OK: OUTLOOK_TOOLS_CONFIG est bien présent")
    else:
        print("❌ Erreur: OUTLOOK_TOOLS_CONFIG non trouvé après ajout")
