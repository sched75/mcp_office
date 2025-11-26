# Guide Utilisateur - MCP Office

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

## Word (65 outils)

Microsoft Word - Traitement de texte et création de documents.

### Catégories d'outils

#### Gestion documents (6 outils)

- **`word_create_document`** : Crée un nouveau document Word
- **`word_open_document`** : Ouvre un document existant
- **`word_save_document`** : Enregistre le document
- **`word_close_document`** : Ferme le document
- **`word_save_as_template`** : Sauvegarde comme modèle
- **`word_print_to_pdf`** : Exporte en PDF

#### Contenu textuel (4 outils)

- **`word_add_paragraph`** : Ajoute un paragraphe
- **`word_insert_text_at_position`** : Insère du texte à une position
- **`word_find_and_replace`** : Recherche et remplace
- **`word_delete_text`** : Supprime du texte

### Exemples Word

#### Exemple 1 : Créer un rapport Word complet

Démonstration de création de document avec plusieurs éléments

**Prompt** :
```
Crée un document Word avec le titre 'Rapport Annuel 2024', ajoute un paragraphe d'introduction, insère un tableau 3x3, et sauvegarde-le
```

---

#### Exemple 2 : Publipostage

Utilisation de la fonctionnalité mail merge

**Prompt** :
```
Crée un document Word et effectue un publipostage avec les données : Nom=['Alice', 'Bob'], Email=['alice@test.com', 'bob@test.com']
```

---

## Excel (82 outils)

Microsoft Excel - Tableur et analyse de données.

### Catégories d'outils

#### Gestion classeurs (6 outils)

- **`excel_create_workbook`** : Crée un nouveau classeur
- **`excel_write_cell`** : Écrit dans une cellule
- **`excel_create_chart`** : Crée un graphique

### Exemples Excel

#### Exemple 1 : Analyser des données et créer un graphique

Workflow complet d'analyse de données

**Prompt** :
```
Crée un classeur Excel, écris des données de ventes dans A1:B10, calcule la somme en B11, puis crée un graphique en colonnes
```

---

## PowerPoint (63 outils)

Microsoft PowerPoint - Présentations et diaporamas.

### Catégories d'outils

#### Gestion présentations (6 outils)

- **`powerpoint_create_presentation`** : Crée une présentation
- **`powerpoint_add_slide`** : Ajoute une diapositive

### Exemples PowerPoint

#### Exemple 1 : Créer une présentation de pitch

Création de présentation professionnelle

**Prompt** :
```
Crée une présentation PowerPoint avec 5 diapositives : page de titre, problème, solution, marché, conclusion. Ajoute des images et animations
```

---

## Outlook (85 outils)

Microsoft Outlook - Emails, calendrier, contacts et tâches.

### Catégories d'outils

#### Emails (12 outils)

- **`outlook_send_email`** : Envoie un email
- **`outlook_read_email`** : Lit un email
- **`outlook_reply_to_email`** : Répond à un email
- **`outlook_search_emails`** : Recherche des emails

#### Calendrier (10 outils)

- **`outlook_create_appointment`** : Crée un rendez-vous
- **`outlook_create_recurring_event`** : Crée un événement récurrent

#### Contacts (9 outils)

- **`outlook_create_contact`** : Crée un contact
- **`outlook_search_contact`** : Recherche un contact

#### Tâches (7 outils)

- **`outlook_create_task`** : Crée une tâche
- **`outlook_mark_task_complete`** : Marque comme terminée

### Exemples Outlook

#### Exemple 1 : Organiser une réunion

Workflow complet de gestion de réunion

**Prompt** :
```
Crée un rendez-vous Outlook pour demain à 10h, intitulé 'Réunion d'équipe', durée 1h, avec 5 participants, puis envoie les invitations
```

---

#### Exemple 2 : Gérer sa boîte de réception

Organisation automatique des emails

**Prompt** :
```
Cherche tous les emails non lus de la semaine dernière concernant 'projet', crée un dossier 'Projet Important', déplace-les dedans
```

---

## Exemples Avancés

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
  Bon : "C:\Users\NOM\Documents\rapport.docx"
  Mauvais : "rapport.docx"
  ```

- **Vérifier l'existence des fichiers avant ouverture**
  ```
  Liste les fichiers .xlsx dans C:\Data\, puis ouvre "ventes.xlsx"
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
