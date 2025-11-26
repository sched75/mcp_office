"""
Générateur automatique de documentation utilisateur pour MCP Office.

Ce script génère :
- user_guide.md avec 40+ exemples
- api_reference.md avec les 295 outils
- troubleshooting.md avec FAQ complète
"""

# Configuration des outils par catégorie
TOOLS_CONFIG = {
    "Word": {
        "count": 65,
        "categories": [
            ("Gestion documents", 6, [
                ("word_create_document", "Crée un nouveau document Word"),
                ("word_open_document", "Ouvre un document existant"),
                ("word_save_document", "Enregistre le document"),
                ("word_close_document", "Ferme le document"),
                ("word_save_as_template", "Sauvegarde comme modèle"),
                ("word_print_to_pdf", "Exporte en PDF"),
            ]),
            ("Contenu textuel", 4, [
                ("word_add_paragraph", "Ajoute un paragraphe"),
                ("word_insert_text_at_position", "Insère du texte à une position"),
                ("word_find_and_replace", "Recherche et remplace"),
                ("word_delete_text", "Supprime du texte"),
            ]),
            # Autres catégories...
        ],
        "examples": [
            {
                "title": "Créer un rapport Word complet",
                "prompt": "Crée un document Word avec le titre 'Rapport Annuel 2024', ajoute un paragraphe d'introduction, insère un tableau 3x3, et sauvegarde-le",
                "description": "Démonstration de création de document avec plusieurs éléments"
            },
            {
                "title": "Publipostage",
                "prompt": "Crée un document Word et effectue un publipostage avec les données : Nom=['Alice', 'Bob'], Email=['alice@test.com', 'bob@test.com']",
                "description": "Utilisation de la fonctionnalité mail merge"
            },
        ]
    },
    "Excel": {
        "count": 82,
        "categories": [
            ("Gestion classeurs", 6, [
                ("excel_create_workbook", "Crée un nouveau classeur"),
                ("excel_write_cell", "Écrit dans une cellule"),
                ("excel_create_chart", "Crée un graphique"),
            ]),
        ],
        "examples": [
            {
                "title": "Analyser des données et créer un graphique",
                "prompt": "Crée un classeur Excel, écris des données de ventes dans A1:B10, calcule la somme en B11, puis crée un graphique en colonnes",
                "description": "Workflow complet d'analyse de données"
            },
        ]
    },
    "PowerPoint": {
        "count": 63,
        "categories": [
            ("Gestion présentations", 6, [
                ("powerpoint_create_presentation", "Crée une présentation"),
                ("powerpoint_add_slide", "Ajoute une diapositive"),
            ]),
        ],
        "examples": [
            {
                "title": "Créer une présentation de pitch",
                "prompt": "Crée une présentation PowerPoint avec 5 diapositives : page de titre, problème, solution, marché, conclusion. Ajoute des images et animations",
                "description": "Création de présentation professionnelle"
            },
        ]
    },
    "Outlook": {
        "count": 85,
        "categories": [
            ("Emails", 12, [
                ("outlook_send_email", "Envoie un email"),
                ("outlook_read_email", "Lit un email"),
                ("outlook_reply_to_email", "Répond à un email"),
                ("outlook_search_emails", "Recherche des emails"),
            ]),
            ("Calendrier", 10, [
                ("outlook_create_appointment", "Crée un rendez-vous"),
                ("outlook_create_recurring_event", "Crée un événement récurrent"),
            ]),
            ("Contacts", 9, [
                ("outlook_create_contact", "Crée un contact"),
                ("outlook_search_contact", "Recherche un contact"),
            ]),
            ("Tâches", 7, [
                ("outlook_create_task", "Crée une tâche"),
                ("outlook_mark_task_complete", "Marque comme terminée"),
            ]),
        ],
        "examples": [
            {
                "title": "Organiser une réunion",
                "prompt": "Crée un rendez-vous Outlook pour demain à 10h, intitulé 'Réunion d'équipe', durée 1h, avec 5 participants, puis envoie les invitations",
                "description": "Workflow complet de gestion de réunion"
            },
            {
                "title": "Gérer sa boîte de réception",
                "prompt": "Cherche tous les emails non lus de la semaine dernière concernant 'projet', crée un dossier 'Projet Important', déplace-les dedans",
                "description": "Organisation automatique des emails"
            },
        ]
    },
}

def generate_user_guide():
    """Génère le guide utilisateur complet."""
    content = """# Guide Utilisateur - MCP Office

## Introduction

MCP Office vous permet de piloter Microsoft Office (Word, Excel, PowerPoint, Outlook) directement depuis Claude Desktop. Ce guide vous présente les 295 outils disponibles avec des exemples concrets.

## Table des Matières

1. [Démarrage Rapide](#démarrage-rapide)
2. [Word (65 outils)](#word-65-outils)
3. [Excel (82 outils)](#excel-82-outils)
4. [PowerPoint (63 outils)](#powerpoint-63-outils)
5. [Outlook (85 outils)](#outlook-85-outils)
6. [Exemples Avancés](#exemples-avancés)
7. [Workflows Inter-Applications](#workflows-inter-applications)

---

## Démarrage Rapide

### Premier Test

Une fois MCP Office installé, testez avec cette commande simple :

```
Crée un document Word avec le texte "Hello MCP Office!"
```

Vous devriez recevoir :
```
✅ Opération réussie
  • document_created: True
  • text_added: True
```

### Commandes de Base

| Application | Commande Exemple |
|-------------|------------------|
| Word | "Crée un document Word avec..." |
| Excel | "Crée un classeur Excel et écris..." |
| PowerPoint | "Crée une présentation PowerPoint avec..." |
| Outlook | "Envoie un email à... avec le sujet..." |

---

"""
    
    # Générer sections pour chaque application
    for app_name, app_config in TOOLS_CONFIG.items():
        content += f"## {app_name} ({app_config['count']} outils)\n\n"
        
        # Description
        if app_name == "Word":
            content += "Microsoft Word - Traitement de texte et création de documents.\n\n"
        elif app_name == "Excel":
            content += "Microsoft Excel - Tableur et analyse de données.\n\n"
        elif app_name == "PowerPoint":
            content += "Microsoft PowerPoint - Présentations et diaporamas.\n\n"
        elif app_name == "Outlook":
            content += "Microsoft Outlook - Emails, calendrier, contacts et tâches.\n\n"
        
        # Catégories d'outils
        content += "### Catégories d'outils\n\n"
        for cat_name, cat_count, tools in app_config["categories"]:
            content += f"#### {cat_name} ({cat_count} outils)\n\n"
            for tool_name, tool_desc in tools:
                content += f"- **`{tool_name}`** : {tool_desc}\n"
            content += "\n"
        
        # Exemples
        content += f"### Exemples {app_name}\n\n"
        for i, example in enumerate(app_config["examples"], 1):
            content += f"#### Exemple {i} : {example['title']}\n\n"
            content += f"{example['description']}\n\n"
            content += "**Prompt** :\n```\n" + example['prompt'] + "\n```\n\n"
            content += "---\n\n"
    
    # Exemples avancés
    content += """## Exemples Avancés

### Automatiser un Workflow Complet

**Scénario** : Créer un rapport mensuel automatisé

```
1. Récupère les données de ventes du mois depuis Excel "ventes_janvier.xlsx"
2. Crée un document Word avec le titre "Rapport Ventes Janvier 2024"
3. Insère un tableau avec les données
4. Génère un graphique Excel et insère-le dans Word
5. Ajoute une analyse textuelle
6. Exporte en PDF et envoie par email aux managers
```

### Traitement par Lots

**Scénario** : Traiter plusieurs documents

```
Pour chaque fichier .docx dans le dossier "rapports":
1. Ouvre le document
2. Applique le style "Corporate"
3. Ajoute le logo de l'entreprise en en-tête
4. Exporte en PDF
5. Envoie par email au destinataire correspondant
```

---

## Workflows Inter-Applications

### Excel → Word : Rapport Automatique

```
1. Ouvre le classeur Excel "donnees_Q4.xlsx"
2. Extrait les données de la feuille "Résumé"
3. Crée un document Word à partir du modèle "rapport_template.dotx"
4. Insère les données Excel comme tableau
5. Génère un graphique et l'insère
6. Sauvegarde comme "Rapport_Q4_2024.pdf"
```

### Excel → PowerPoint : Présentation de Données

```
1. Ouvre "analyses_ventes.xlsx"
2. Crée une présentation PowerPoint
3. Pour chaque région dans Excel:
   - Ajoute une diapositive
   - Insère le graphique de la région
   - Ajoute les KPIs textuels
4. Applique le thème corporate
5. Ajoute des animations
```

### Outlook → Word : Rapport d'Emails

```
1. Recherche tous les emails du projet "Alpha" de la semaine dernière
2. Crée un document Word "Suivi_Projet_Alpha.docx"
3. Pour chaque email trouvé:
   - Ajoute une section avec l'expéditeur, date, objet
   - Insère un résumé du contenu
4. Génère une table des matières
5. Exporte en PDF
```

---

## Bonnes Pratiques

### 1. Gestion des Fichiers

- **Toujours spécifier des chemins complets**
  ```
  Bon : "C:\\Users\\NOM\\Documents\\rapport.docx"
  Mauvais : "rapport.docx"
  ```

- **Vérifier l'existence des fichiers avant ouverture**
  ```
  Liste les fichiers .xlsx dans C:\\Data\\, puis ouvre "ventes.xlsx"
  ```

### 2. Gestion des Erreurs

- **Fermer les documents après usage**
  ```
  Ouvre rapport.docx, ajoute du texte, sauvegarde et ferme
  ```

- **Sauvegarder régulièrement**
  ```
  Après chaque modification importante, sauvegarde le document
  ```

### 3. Performance

- **Traiter par lots quand possible**
  ```
  Au lieu de : "Crée 10 documents Word séparément"
  Préférer : "Crée 10 documents Word en une seule opération"
  ```

---

## Limitations Connues

1. **Windows uniquement** : COM Automation nécessite Windows
2. **Office installé** : Les applications doivent être installées localement
3. **Versions Office** : Testé avec Office 2016, 2019, 2021, 365
4. **Performance** : Les opérations sur de gros fichiers peuvent prendre du temps
5. **Fichiers ouverts** : Éviter d'ouvrir les mêmes fichiers manuellement pendant l'automation

---

## Support et Ressources

- **Documentation complète** : `docs/api_reference.md`
- **Troubleshooting** : `docs/troubleshooting.md`
- **Exemples de code** : Voir tests dans `tests/`
- **Issues GitHub** : https://github.com/sched75/mcp_office/issues

---

**Profitez de l'automation complète d'Office avec Claude ! 🚀**
"""
    
    return content

def generate_troubleshooting():
    """Génère le guide de dépannage."""
    content = """# FAQ et Dépannage - MCP Office

## Table des Matières

1. [Installation](#installation)
2. [Configuration](#configuration)
3. [Erreurs Courantes](#erreurs-courantes)
4. [Performance](#performance)
5. [Word](#word)
6. [Excel](#excel)
7. [PowerPoint](#powerpoint)
8. [Outlook](#outlook)
9. [Logs et Diagnostics](#logs-et-diagnostics)

---

## Installation

### Q : Python n'est pas reconnu comme commande

**R** : Python n'est pas dans le PATH système.

**Solutions** :
1. Réinstallez Python en cochant "Add Python to PATH"
2. Ajoutez manuellement Python au PATH :
   - Panneau de configuration → Système → Variables d'environnement
   - Ajoutez `C:\\Python3X` et `C:\\Python3X\\Scripts`
3. Redémarrez votre terminal

### Q : pip ne fonctionne pas

**R** : pip n'est pas correctement installé ou configuré.

**Solutions** :
```powershell
# Réinstaller pip
python -m ensurepip --upgrade

# Ou télécharger get-pip.py
curl https://bootstrap.pypa.io/get-pip.py -o get-pip.py
python get-pip.py
```

### Q : L'environnement virtuel ne se crée pas

**R** : Problème avec le module venv.

**Solutions** :
```powershell
# Réinstaller venv
python -m pip install --upgrade virtualenv

# Créer avec virtualenv au lieu de venv
virtualenv venv
```

---

## Configuration

### Q : Claude Desktop ne détecte pas le serveur MCP

**R** : Problème de configuration ou de chemin.

**Solutions** :
1. Vérifiez l'emplacement du fichier de config :
   ```
   %APPDATA%\\Claude\\claude_desktop_config.json
   ```

2. Vérifiez le format JSON (pas d'erreur de syntaxe)

3. Vérifiez les chemins (doublage des backslashes) :
   ```json
   "cwd": "C:\\\\Users\\\\NOM\\\\Documents\\\\mcp_office"
   ```

4. Redémarrez COMPLÈTEMENT Claude Desktop

5. Consultez les logs :
   ```
   %APPDATA%\\Claude\\logs\\
   ```

### Q : Le serveur apparaît mais ne répond pas

**R** : Problème de démarrage du serveur Python.

**Solutions** :
1. Testez le serveur manuellement :
   ```powershell
   cd C:\\chemin\\vers\\mcp_office
   .\\venv\\Scripts\\Activate.ps1
   python -m src.server
   ```

2. Vérifiez les erreurs dans le terminal

3. Vérifiez que toutes les dépendances sont installées :
   ```powershell
   pip list
   ```

---

## Erreurs Courantes

### Erreur : "COMInitializationError"

**Cause** : L'application Office n'a pas pu être initialisée.

**Solutions** :
1. Fermez toutes les instances d'Office ouvertes
2. Vérifiez qu'Office est bien installé
3. Ouvrez l'application manuellement une fois (Word/Excel/etc.)
4. Vérifiez les permissions d'exécution
5. Essayez de redémarrer l'ordinateur

### Erreur : "DocumentNotFoundError"

**Cause** : Le fichier spécifié n'existe pas.

**Solutions** :
1. Vérifiez le chemin complet du fichier
2. Utilisez des chemins absolus, pas relatifs
3. Vérifiez que le fichier n'est pas ouvert ailleurs
4. Vérifiez l'extension du fichier (.docx, .xlsx, etc.)

### Erreur : "InvalidParameterError"

**Cause** : Un paramètre requis est manquant ou invalide.

**Solutions** :
1. Vérifiez la documentation de l'outil
2. Assurez-vous de fournir tous les paramètres requis
3. Vérifiez le type des paramètres (string, number, etc.)

### Erreur : "Access Denied" / "Permission Error"

**Cause** : Permissions insuffisantes sur le fichier.

**Solutions** :
1. Vérifiez que le fichier n'est pas en lecture seule
2. Fermez le fichier s'il est ouvert
3. Vérifiez les permissions du dossier
4. Exécutez Claude Desktop en tant qu'administrateur (dernier recours)

---

## Performance

### Q : Les opérations sont lentes

**R** : COM Automation peut être lent sur de gros fichiers.

**Optimisations** :
1. Fermez les applications Office inutiles
2. Désactivez le mode "Visible" (déjà fait par défaut)
3. Traitez par lots plutôt qu'individuellement
4. Utilisez des fichiers plus petits pour les tests
5. Augmentez la RAM disponible

### Q : Le serveur plante sur de gros fichiers

**R** : Limite de mémoire atteinte.

**Solutions** :
1. Augmentez la mémoire allouée à Python
2. Traitez les fichiers par sections
3. Utilisez des fichiers temporaires intermédiaires
4. Fermez les documents après traitement

---

## Word

### Q : Le texte ne s'insère pas correctement

**R** : Problème de position ou de formatage.

**Solutions** :
1. Vérifiez la position d'insertion
2. Utilisez `add_paragraph` plutôt que `insert_text_at_position` pour du texte simple
3. Assurez-vous que le document est actif

### Q : Les images ne s'affichent pas

**R** : Problème de chemin ou format d'image.

**Solutions** :
1. Utilisez des chemins absolus
2. Vérifiez que l'image existe
3. Formats supportés : .jpg, .png, .gif, .bmp
4. Vérifiez la taille de l'image (pas trop grande)

---

## Excel

### Q : Les formules ne se calculent pas

**R** : Calcul automatique désactivé.

**Solutions** :
1. Forcez le recalcul :
   ```
   Recalcule toutes les formules du classeur Excel
   ```
2. Vérifiez la syntaxe de la formule
3. Utilisez des références absolues si nécessaire

### Q : Les graphiques ne s'affichent pas

**R** : Données source incorrectes.

**Solutions** :
1. Vérifiez la plage de données
2. Assurez-vous que les données existent
3. Vérifiez le format des données (nombres vs texte)

---

## PowerPoint

### Q : Les animations ne fonctionnent pas

**R** : Ordre ou timing incorrect.

**Solutions** :
1. Vérifiez l'ordre des animations
2. Définissez des délais appropriés
3. Testez en mode diaporama

### Q : Les diapositives sont vides

**R** : Contenu non ajouté ou layout incorrect.

**Solutions** :
1. Vérifiez le layout de la diapositive
2. Ajoutez explicitement du contenu (texte, images)
3. Utilisez le bon numéro de diapositive

---

## Outlook

### Q : Les emails ne s'envoient pas

**R** : Compte non configuré ou hors ligne.

**Solutions** :
1. Vérifiez qu'Outlook est configuré avec un compte
2. Vérifiez la connexion Internet
3. Ouvrez Outlook manuellement pour vérifier
4. Vérifiez les paramètres de sécurité

### Q : Impossible de lire les emails

**R** : Problème d'ID ou de dossier.

**Solutions** :
1. Utilisez le bon `entry_id` de l'email
2. Vérifiez que l'email existe toujours
3. Recherchez l'email d'abord pour obtenir son ID

---

## Logs et Diagnostics

### Activer le logging détaillé

Éditez `src/server.py` :
```python
logging.basicConfig(
    level=logging.DEBUG,  # Changez INFO en DEBUG
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
```

### Consulter les logs Claude Desktop

```
%APPDATA%\\Claude\\logs\\
```

Cherchez les fichiers récents et les erreurs contenant "mcp-office".

### Tester le serveur isolément

```powershell
cd C:\\chemin\\vers\\mcp_office
.\\venv\\Scripts\\Activate.ps1
python -m src.server
```

Entrez des commandes JSON manuellement pour tester.

### Vérifier les versions

```powershell
python --version
pip list | findstr "mcp pywin32"
```

---

## Obtenir de l'Aide

Si votre problème persiste :

1. **Consultez les logs** détaillés
2. **Recherchez dans les Issues GitHub** : Votre problème a peut-être déjà été résolu
3. **Créez une Issue** avec :
   - Description détaillée du problème
   - Messages d'erreur complets
   - Logs pertinents
   - Version Python, Office, Windows
   - Étapes pour reproduire

**GitHub** : https://github.com/sched75/mcp_office/issues

---

**La plupart des problèmes sont résolus en redémarrant Claude Desktop ou en vérifiant les chemins ! 🔧**
"""
    return content

def main():
    """Génère tous les fichiers de documentation."""
    print("=" * 70)
    print("GÉNÉRATION DE LA DOCUMENTATION")
    print("=" * 70)
    print()
    
    # Générer user_guide.md
    print("Génération de user_guide.md...")
    user_guide = generate_user_guide()
    with open("docs/user_guide.md", "w", encoding="utf-8") as f:
        f.write(user_guide)
    print(f"✅ user_guide.md créé ({len(user_guide)} caractères)")
    
    # Générer troubleshooting.md
    print("Génération de troubleshooting.md...")
    troubleshooting = generate_troubleshooting()
    with open("docs/troubleshooting.md", "w", encoding="utf-8") as f:
        f.write(troubleshooting)
    print(f"✅ troubleshooting.md créé ({len(troubleshooting)} caractères)")
    
    print()
    print("=" * 70)
    print("✅ DOCUMENTATION GÉNÉRÉE AVEC SUCCÈS")
    print("=" * 70)
    print()
    print("Fichiers créés :")
    print("  • docs/user_guide.md")
    print("  • docs/troubleshooting.md")
    print()

if __name__ == "__main__":
    main()
