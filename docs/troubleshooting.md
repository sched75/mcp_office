# FAQ et Dépannage - MCP Office

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
   - Ajoutez `C:\Python3X` et `C:\Python3X\Scripts`
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
   %APPDATA%\Claude\claude_desktop_config.json
   ```

2. Vérifiez le format JSON (pas d'erreur de syntaxe)

3. Vérifiez les chemins (doublage des backslashes) :
   ```json
   "cwd": "C:\\Users\\NOM\\Documents\\mcp_office"
   ```

4. Redémarrez COMPLÈTEMENT Claude Desktop

5. Consultez les logs :
   ```
   %APPDATA%\Claude\logs\
   ```

### Q : Le serveur apparaît mais ne répond pas

**R** : Problème de démarrage du serveur Python.

**Solutions** :
1. Testez le serveur manuellement :
   ```powershell
   cd C:\chemin\vers\mcp_office
   .\venv\Scripts\Activate.ps1
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
%APPDATA%\Claude\logs\
```

Cherchez les fichiers récents et les erreurs contenant "mcp-office".

### Tester le serveur isolément

```powershell
cd C:\chemin\vers\mcp_office
.\venv\Scripts\Activate.ps1
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
