# 🪟 Configuration et Lancement sur Windows

Guide complet pour installer et lancer le serveur MCP Office Automation sur Windows.

## 📋 Prérequis

### 1. Microsoft Office
- **Microsoft Word**, **Excel** et/ou **PowerPoint** installés
- Office 2016 ou version ultérieure recommandée
- Office doit être activé et fonctionnel

### 2. Python
- **Python 3.10 ou supérieur** installé
- Téléchargeable depuis [python.org](https://www.python.org/downloads/)
- ⚠️ **Important**: Cocher "Add Python to PATH" lors de l'installation

### 3. Vérification de l'installation
```powershell
# Ouvrir PowerShell et vérifier Python
python --version
# Doit afficher: Python 3.10.x ou supérieur

# Vérifier pip
pip --version
```

## 🚀 Installation

### Étape 1: Cloner le projet
```powershell
cd C:\Users\VotreNom\Documents
git clone <url-du-repo>
cd mcp_office
```

### Étape 2: Créer un environnement virtuel (recommandé)
```powershell
# Créer l'environnement virtuel
python -m venv venv

# Activer l'environnement virtuel
.\venv\Scripts\Activate.ps1

# Si erreur de politique d'exécution, exécuter d'abord:
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Étape 3: Installer les dépendances
```powershell
# Mettre à jour pip
python -m pip install --upgrade pip

# Installer toutes les dépendances
pip install -r requirements.txt

# Vérifier que pywin32 est bien installé
python -c "import win32com.client; print('✅ pywin32 installé correctement')"
```

## 🎮 Lancement du Serveur

### Méthode 1: Ligne de commande
```powershell
# S'assurer que l'environnement virtuel est activé
.\venv\Scripts\Activate.ps1

# Lancer le serveur MCP
python -m src.server
```

### Méthode 2: Script de démarrage automatique
Créer un fichier `start_server.bat`:

```batch
@echo off
echo 🚀 Démarrage du serveur MCP Office Automation...
echo.

REM Activer l'environnement virtuel
call venv\Scripts\activate.bat

REM Lancer le serveur
python -m src.server

pause
```

Double-cliquer sur `start_server.bat` pour lancer.

## 🔧 Configuration MCP

### Pour Claude Desktop (Windows)

Éditer le fichier de configuration MCP:
```
%APPDATA%\Claude\claude_desktop_config.json
```

Ajouter la configuration suivante:

```json
{
  "mcpServers": {
    "office-automation": {
      "command": "python",
      "args": [
        "-m",
        "src.server"
      ],
      "cwd": "C:\\Users\\VotreNom\\Documents\\mcp_office",
      "env": {
        "PYTHONPATH": "C:\\Users\\VotreNom\\Documents\\mcp_office"
      }
    }
  }
}
```

⚠️ **Remplacer** `C:\\Users\\VotreNom\\Documents\\mcp_office` par le chemin réel du projet.

### Pour autres clients MCP

Utiliser la commande:
```
python -m src.server
```

avec le répertoire de travail: `C:\chemin\vers\mcp_office`

## 🧪 Test de Fonctionnement

### Test 1: Import des modules
```powershell
python -c "from src.word.word_service import WordService; print('✅ Word OK')"
python -c "from src.excel.excel_service import ExcelService; print('✅ Excel OK')"
python -c "from src.powerpoint.powerpoint_service import PowerPointService; print('✅ PowerPoint OK')"
```

### Test 2: Création d'un document Word
```python
# test_word.py
from src.word.word_service import WordService

service = WordService()
service.initialize()
result = service.create_document()
print(f"✅ Document créé: {result}")

service.add_paragraph("Bonjour depuis Python!")
service.save_document()
service.cleanup()
print("✅ Test Word réussi!")
```

Exécuter:
```powershell
python test_word.py
```

### Test 3: Lancer les tests unitaires
```powershell
# Tous les tests
pytest tests/ -v

# Tests spécifiques aux services (nécessite Office)
pytest tests/test_word_service.py -v
pytest tests/test_excel_service.py -v
pytest tests/test_powerpoint_service.py -v

# Avec rapport de couverture
pytest tests/ --cov=src --cov-report=html
```

## ⚙️ Outils MCP Disponibles

Une fois le serveur lancé, les outils suivants sont disponibles:

### 📝 Word (65+ outils)
- `word_create_document` - Créer un nouveau document
- `word_add_paragraph` - Ajouter un paragraphe
- `word_insert_table` - Insérer un tableau
- `word_insert_image` - Insérer une image
- Et 60+ autres outils...

### 📊 Excel (82+ outils)
- `excel_create_workbook` - Créer un classeur
- `excel_write_cell` - Écrire dans une cellule
- `excel_create_chart` - Créer un graphique
- `excel_create_pivot_table` - Créer un tableau croisé dynamique
- Et 78+ autres outils...

### 📽️ PowerPoint (63+ outils)
- `powerpoint_create_presentation` - Créer une présentation
- `powerpoint_add_slide` - Ajouter une diapositive
- `powerpoint_insert_image` - Insérer une image
- `powerpoint_add_animation` - Ajouter une animation
- Et 59+ autres outils...

## 🐛 Dépannage

### Erreur: "No module named 'win32com'"
```powershell
# Réinstaller pywin32
pip uninstall pywin32
pip install pywin32

# Post-installation pywin32
python venv\Scripts\pywin32_postinstall.py -install
```

### Erreur: "COM object initialization failed"
- Vérifier qu'Office est bien installé et activé
- Essayer de fermer tous les processus Office (Word, Excel, PowerPoint)
- Redémarrer le serveur

### Erreur: "Access is denied" ou problèmes de permissions
- Exécuter PowerShell en tant qu'administrateur
- Ou modifier la politique d'exécution:
  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
  ```

### Office n'est pas détecté
```python
# Vérifier l'installation COM d'Office
python
>>> import win32com.client
>>> word = win32com.client.Dispatch("Word.Application")
>>> print(word.Version)  # Doit afficher la version d'Office
>>> word.Quit()
```

### Processus Office restent en arrière-plan
```powershell
# Tuer tous les processus Office
taskkill /F /IM WINWORD.EXE
taskkill /F /IM EXCEL.EXE
taskkill /F /IM POWERPNT.EXE
```

## 📚 Ressources Supplémentaires

- **Documentation MCP**: [modelcontextprotocol.io](https://modelcontextprotocol.io)
- **Documentation pywin32**: [pypi.org/project/pywin32](https://pypi.org/project/pywin32/)
- **Office VBA Reference**: [docs.microsoft.com](https://docs.microsoft.com/office/vba/api/overview/)

## 🔒 Sécurité

⚠️ **Avertissements importants**:

1. **Macros et sécurité**: Le serveur peut exécuter des opérations Office - utilisez-le uniquement avec des sources de confiance
2. **Fichiers**: Ne pas ouvrir de fichiers Office non vérifiés
3. **Permissions**: Le serveur a accès complet à Office - surveillez les opérations

## 🎯 Performance

### Optimisations recommandées:

1. **Désactiver l'affichage**:
   - Les opérations sont plus rapides sans afficher l'interface Office
   - C'est le comportement par défaut du serveur

2. **Fermer les documents**:
   - Toujours appeler les méthodes de nettoyage
   - Éviter les processus Office orphelins

3. **Batch operations**:
   - Grouper les opérations pour réduire les appels COM
   - Utiliser les méthodes de bulk quand disponibles

## ✅ Checklist de Démarrage Rapide

- [ ] Python 3.10+ installé
- [ ] Office installé et activé
- [ ] Environnement virtuel créé (`python -m venv venv`)
- [ ] Dépendances installées (`pip install -r requirements.txt`)
- [ ] pywin32 vérifié (`python -c "import win32com.client"`)
- [ ] Serveur lancé (`python -m src.server`)
- [ ] Tests passent (`pytest tests/` - optionnel)

🎉 **Bon usage du serveur MCP Office Automation!**
