# Guide d'Installation - MCP Office

## Table des Matières
1. [Prérequis](#prérequis)
2. [Installation Automatique](#installation-automatique)
3. [Installation Manuelle](#installation-manuelle)
4. [Configuration Claude Desktop](#configuration-claude-desktop)
5. [Vérification](#vérification)
6. [Dépannage](#dépannage)

---

## Prérequis

### Système
- **OS** : Windows 10/11 (requis pour COM Automation)
- **Microsoft Office** : Word, Excel, PowerPoint et/ou Outlook installés
- **Python** : Version 3.8 ou supérieure
- **Claude Desktop** : Dernière version installée

### Vérification des prérequis

```powershell
# Vérifier Python
python --version
# Doit afficher Python 3.8+

# Vérifier Office (PowerShell)
Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Office\*\Word\InstallRoot
Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Office\*\Excel\InstallRoot
Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Office\*\PowerPoint\InstallRoot
Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Office\*\Outlook\InstallRoot
```

---

## Installation Automatique

### Méthode Recommandée

1. **Cloner ou télécharger le projet**
```powershell
cd C:\Users\VOTRE_NOM\Documents
git clone https://github.com/sched75/mcp_office.git
cd mcp_office
```

2. **Exécuter le script d'installation**
```powershell
.\scripts\install.ps1
```

Le script va automatiquement :
- ✅ Vérifier Python
- ✅ Créer l'environnement virtuel
- ✅ Installer les dépendances
- ✅ Détecter les applications Office
- ✅ Configurer Claude Desktop
- ✅ Vérifier l'installation

3. **Redémarrer Claude Desktop**

Fermez complètement Claude Desktop et relancez-le.

---

## Installation Manuelle

Si le script automatique ne fonctionne pas, suivez ces étapes :

### 1. Créer l'environnement virtuel

```powershell
cd C:\chemin\vers\mcp_office
python -m venv venv
.\venv\Scripts\Activate.ps1
```

### 2. Installer les dépendances

```powershell
pip install --upgrade pip
pip install -r requirements.txt
```

### 3. Configurer Claude Desktop

Ouvrez ou créez le fichier de configuration :
```
%APPDATA%\Claude\claude_desktop_config.json
```

Ajoutez cette configuration :
```json
{
  "mcpServers": {
    "mcp-office": {
      "command": "python",
      "args": [
        "-m",
        "src.server"
      ],
      "cwd": "C:\\chemin\\vers\\mcp_office",
      "env": {
        "PYTHONPATH": "C:\\chemin\\vers\\mcp_office",
        "PYTHON_UNBUFFERED": "1"
      },
      "disabled": false
    }
  }
}
```

⚠️ **Important** : Remplacez `C:\\chemin\\vers\\mcp_office` par le chemin réel vers votre projet.

### 4. Redémarrer Claude Desktop

Fermez complètement Claude Desktop et relancez-le.

---

## Configuration Claude Desktop

### Emplacement du fichier de configuration

Le fichier de configuration se trouve à :
```
%APPDATA%\Claude\claude_desktop_config.json
```

Chemin complet typique :
```
C:\Users\VOTRE_NOM\AppData\Roaming\Claude\claude_desktop_config.json
```

### Structure de configuration

```json
{
  "mcpServers": {
    "mcp-office": {
      "command": "python",
      "args": ["-m", "src.server"],
      "cwd": "C:\\chemin\\complet\\vers\\mcp_office",
      "env": {
        "PYTHONPATH": "C:\\chemin\\complet\\vers\\mcp_office",
        "PYTHON_UNBUFFERED": "1"
      },
      "disabled": false
    }
  }
}
```

### Fusion avec configuration existante

Si vous avez déjà d'autres serveurs MCP configurés :

```json
{
  "mcpServers": {
    "mon-autre-serveur": {
      "command": "...",
      "args": [...]
    },
    "mcp-office": {
      "command": "python",
      "args": ["-m", "src.server"],
      "cwd": "C:\\chemin\\vers\\mcp_office",
      "env": {
        "PYTHONPATH": "C:\\chemin\\vers\\mcp_office",
        "PYTHON_UNBUFFERED": "1"
      }
    }
  }
}
```

---

## Vérification

### 1. Vérifier la détection du serveur

Ouvrez Claude Desktop et tapez :
```
Quels serveurs MCP sont disponibles ?
```

Vous devriez voir `mcp-office` dans la liste.

### 2. Test basique Word

```
Crée un nouveau document Word et ajoute le paragraphe "Test MCP Office"
```

Si cela fonctionne, vous devriez recevoir une confirmation :
```
✅ Opération réussie
  • document_created: True
```

### 3. Test basique Excel

```
Crée un nouveau classeur Excel et écris "Hello World" dans la cellule A1
```

### 4. Test basique PowerPoint

```
Crée une nouvelle présentation PowerPoint et ajoute une diapositive avec le titre "Test MCP"
```

### 5. Test basique Outlook

```
Liste mes comptes Outlook configurés
```

---

## Dépannage

### Problème : "Python n'est pas reconnu"

**Solution** :
1. Vérifiez que Python est installé : téléchargez depuis https://www.python.org/
2. Lors de l'installation, cochez "Add Python to PATH"
3. Redémarrez votre terminal

### Problème : "Le serveur MCP ne démarre pas"

**Solutions** :
1. Vérifiez les logs Claude Desktop :
   ```
   %APPDATA%\Claude\logs\
   ```

2. Testez le serveur manuellement :
   ```powershell
   cd C:\chemin\vers\mcp_office
   .\venv\Scripts\Activate.ps1
   python -m src.server
   ```

3. Vérifiez les chemins dans la configuration

### Problème : "Erreur COM / Office non détecté"

**Solutions** :
1. Vérifiez qu'Office est bien installé
2. Essayez d'ouvrir Word/Excel/PowerPoint manuellement une fois
3. Vérifiez les permissions d'exécution

### Problème : "Le serveur apparaît mais les commandes ne fonctionnent pas"

**Solutions** :
1. Vérifiez les logs du serveur
2. Essayez de fermer toutes les applications Office en cours
3. Redémarrez Claude Desktop
4. Consultez `docs/troubleshooting.md` pour plus de détails

---

## Désinstallation

Pour désinstaller MCP Office :

1. **Supprimer la configuration Claude Desktop**
   - Ouvrir `%APPDATA%\Claude\claude_desktop_config.json`
   - Supprimer la section `"mcp-office"` du fichier

2. **Supprimer le projet**
   ```powershell
   cd C:\chemin\vers\
   Remove-Item -Recurse -Force mcp_office
   ```

3. **Redémarrer Claude Desktop**

---

## Mise à jour

Pour mettre à jour vers une nouvelle version :

```powershell
cd C:\chemin\vers\mcp_office
git pull
.\venv\Scripts\Activate.ps1
pip install --upgrade -r requirements.txt
```

Redémarrez ensuite Claude Desktop.

---

## Prochaines Étapes

Une fois l'installation réussie :

1. 📖 Consultez le [Guide Utilisateur](user_guide.md) pour découvrir les 295 outils disponibles
2. 💡 Voir des [Exemples d'Usage](user_guide.md#exemples-complets) pour des cas concrets
3. ❓ Consultez la [FAQ](troubleshooting.md) si vous rencontrez des problèmes

---

## Support

- **Documentation** : Consultez tous les fichiers dans `docs/`
- **Issues GitHub** : https://github.com/sched75/mcp_office/issues
- **Logs** : `%APPDATA%\Claude\logs\` pour les logs Claude Desktop

---

**Installation complétée avec succès ? Profitez de l'automation Office avec Claude ! 🚀**
