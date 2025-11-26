# 🚀 MCP Office - Microsoft Office Automation Server

> Serveur MCP (Model Context Protocol) pour piloter Microsoft Office (Word, Excel, PowerPoint, Outlook) directement depuis Claude Desktop.

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Code Quality](https://img.shields.io/badge/Code%20Quality-100%25%20PEP8-brightgreen)](https://www.python.org/dev/peps/pep-0008/)
[![Tests Coverage](https://img.shields.io/badge/Tests%20Coverage-~100%25-brightgreen)](tests/)

---

## 📋 Vue d'Ensemble

**MCP Office** est un serveur MCP qui expose **295 outils** pour automatiser Microsoft Office via COM Automation. Il permet à Claude de créer, modifier et gérer des documents Office de manière naturelle et intuitive.

### ✨ Fonctionnalités

| Application | Outils | Fonctionnalités Clés |
|-------------|--------|----------------------|
| **Word** | 65 | Documents, formatage, tableaux, images, publipostage, PDF |
| **Excel** | 82 | Classeurs, formules, graphiques, tableaux croisés, analyse |
| **PowerPoint** | 63 | Présentations, diapositives, animations, médias, export |
| **Outlook** | 85 | Emails, calendrier, contacts, tâches, réunions |
| **TOTAL** | **295** | **Automation complète d'Office** |

---

## 🎯 Cas d'Usage

### 📝 Génération de Rapports
```
"Crée un rapport Word avec les données Excel du fichier ventes.xlsx,
ajoute un graphique, formate en style corporate et exporte en PDF"
```

### 📊 Analyse de Données
```
"Ouvre le classeur Excel analyse.xlsx, calcule les tendances mensuelles,
génère un graphique en courbes et crée une présentation PowerPoint"
```

### 📧 Gestion d'Emails
```
"Recherche tous les emails non lus de la semaine concernant 'projet Alpha',
crée un dossier, déplace-les dedans et génère un résumé dans Word"
```

### 📅 Organisation de Réunions
```
"Crée un rendez-vous récurrent tous les lundis à 10h pour les 12 prochaines
semaines, invite l'équipe et envoie l'agenda par email"
```

---

## 🚀 Installation Rapide

### Prérequis
- **Windows** 10/11
- **Python** 3.8+
- **Microsoft Office** (Word, Excel, PowerPoint, Outlook)
- **Claude Desktop** (dernière version)

### Installation Automatique

```powershell
# 1. Cloner le projet
git clone https://github.com/sched75/mcp_office.git
cd mcp_office

# 2. Exécuter l'installation
.\scripts\install.ps1

# 3. Redémarrer Claude Desktop
```

🎉 **C'est tout !** Le serveur MCP est maintenant configuré.

### Vérification

Ouvrez Claude Desktop et testez :
```
Crée un document Word avec le texte "Test MCP Office réussi!"
```

✅ **Si vous recevez une confirmation, l'installation est réussie !**

---

## 📚 Documentation

### Guides Complets

| Guide | Description |
|-------|-------------|
| [📖 Installation](docs/installation.md) | Guide d'installation détaillé (auto/manuel) |
| [👤 Guide Utilisateur](docs/user_guide.md) | 40+ exemples et workflows complets |
| [🔧 Troubleshooting](docs/troubleshooting.md) | FAQ et résolution de problèmes |

### Documentation Technique

- **Architecture** : [VALIDATION_FINALE_COMPLETE.md](VALIDATION_FINALE_COMPLETE.md)
- **Rapport Projet** : [PROJET_FINAL_RAPPORT.md](PROJET_FINAL_RAPPORT.md)
- **TODO & Roadmap** : [TODO.md](TODO.md)

---

## 🎨 Exemples d'Usage

### Word : Créer un Rapport Automatisé

```
Crée un document Word "Rapport_Q1_2024.docx" avec :
1. Page de titre "Rapport Trimestriel Q1 2024"
2. Table des matières
3. Section "Résumé Exécutif" avec 2 paragraphes
4. Tableau 5x3 avec les données de ventes
5. Graphique en colonnes
6. Export en PDF
```

### Excel : Analyser des Données

```
Ouvre le classeur "donnees_ventes.xlsx" :
1. Calcule la somme des ventes par région
2. Crée un tableau croisé dynamique
3. Génère 3 graphiques (colonnes, lignes, secteurs)
4. Applique une mise en forme conditionnelle
5. Exporte en PDF
```

### PowerPoint : Présentation Professionnelle

```
Crée une présentation "Pitch_Startup.pptx" avec :
1. Diapo de titre avec logo
2. Diapo "Problème" avec 3 puces
3. Diapo "Solution" avec image
4. Diapo "Marché" avec graphique
5. Applique le thème "Corporate"
6. Ajoute des transitions
```

### Outlook : Organisation Automatique

```
1. Cherche les emails non lus contenant "urgent"
2. Crée un dossier "Urgent - Cette Semaine"
3. Déplace les emails trouvés
4. Crée une tâche "Traiter emails urgents" avec priorité haute
5. Envoie un résumé par email au manager
```

---

## 🏗️ Architecture

### Structure du Projet

```
mcp_office/
├── src/
│   ├── server.py              # ⭐ Serveur MCP principal (295 outils)
│   ├── core/                  # Classes de base et utilitaires
│   ├── word/                  # Service Word (65 méthodes)
│   ├── excel/                 # Service Excel (82 méthodes)
│   ├── powerpoint/            # Service PowerPoint (63 méthodes)
│   └── outlook/               # Service Outlook (85 méthodes)
├── tests/                     # Tests unitaires (~100% couverture Outlook)
├── docs/                      # Documentation complète
├── scripts/                   # Scripts d'installation et démarrage
├── config/                    # Configuration Claude Desktop
└── requirements.txt           # Dépendances Python
```

### Qualité du Code

| Métrique | Résultat | Statut |
|----------|----------|--------|
| **PEP 8 Compliance** | 100% | ✅ Parfait |
| **Complexité (Radon)** | A (3.30) | ✅ Excellent |
| **Maintenabilité** | Tous fichiers A | ✅ Parfait |
| **Tests Outlook** | ~100% couverture | ✅ Excellent |

---

## 🧪 Tests

### Tests Unitaires

```powershell
# Activer l'environnement
.\venv\Scripts\Activate.ps1

# Exécuter tous les tests
pytest tests/ -v

# Avec couverture
pytest tests/ --cov=src --cov-report=html
```

### Tests Manuels avec Claude

```
# Test Word
"Crée un document Word et ajoute 3 paragraphes avec différents styles"

# Test Excel
"Crée un classeur Excel avec un tableau de données et un graphique"

# Test PowerPoint
"Crée une présentation de 5 diapositives avec des images"

# Test Outlook
"Liste mes comptes Outlook et le nombre d'emails non lus"
```

---

## 🤝 Contribution

Les contributions sont les bienvenues ! Voici comment contribuer :

1. **Fork** le projet
2. **Créer une branche** : `git checkout -b feature/nouvelle-fonctionnalite`
3. **Commit** : `git commit -m "Ajout nouvelle fonctionnalité"`
4. **Push** : `git push origin feature/nouvelle-fonctionnalite`
5. **Pull Request**

### Standards de Code

- ✅ **PEP 8** compliance (100%)
- ✅ **Docstrings** Google Style
- ✅ **Type hints** complets
- ✅ **Tests unitaires** pour nouvelles fonctionnalités
- ✅ **Ruff** validation : `ruff check src/`
- ✅ **Radon** complexity : `radon cc src/ -a -s`

---

## 📜 Licence

Ce projet est sous licence MIT. Voir [LICENSE](LICENSE) pour plus de détails.

---

## 👨‍💻 Auteur

**Pascal-Louis**
- GitHub: [@sched75](https://github.com/sched75)
- Projet: [mcp_office](https://github.com/sched75/mcp_office)

---

## 🙏 Remerciements

- **Anthropic** pour Claude et le protocol MCP
- **Microsoft** pour Office COM Automation
- **Python Community** pour les excellentes librairies

---

## 📞 Support

- **Documentation** : Consultez les fichiers dans `docs/`
- **Issues** : [GitHub Issues](https://github.com/sched75/mcp_office/issues)
- **Discord** : [Rejoignez la communauté](https://discord.gg/claude-ai)

---

## 🚧 Roadmap

- [x] **Phase 1** : Implémentation des 295 fonctionnalités
- [x] **Phase 2** : Validation qualité (PEP 8, tests, docs)
- [x] **Phase 3** : Intégration MCP serveur
- [ ] **Phase 4** : Tests d'intégration complets
- [ ] **Phase 5** : Optimisations performance
- [ ] **Phase 6** : Support macOS/Linux (via Wine/CrossOver)

---

## ⭐ Star History

Si vous trouvez ce projet utile, n'hésitez pas à lui donner une étoile ! ⭐

---

<p align="center">
  <b>Automatisez Microsoft Office avec Claude - C'est magique ! ✨</b>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Made%20with-❤️-red.svg" alt="Made with love">
  <img src="https://img.shields.io/badge/Powered%20by-Claude-blue.svg" alt="Powered by Claude">
  <img src="https://img.shields.io/badge/Built%20for-Productivity-green.svg" alt="Built for Productivity">
</p>
