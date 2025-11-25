# Outlook Service - Documentation

## Vue d'ensemble

Le service Outlook fournit une automatisation complète de Microsoft Outlook avec 85 fonctionnalités couvrant tous les aspects de la gestion des emails, calendriers, contacts, et tâches.

## Installation

Le service est déjà intégré au projet `mcp_office`. Aucune installation supplémentaire n'est nécessaire.

## Utilisation rapide

```python
from src.outlook import OutlookService

# Créer et initialiser le service
outlook = OutlookService()
outlook.initialize()

# Envoyer un email
result = outlook.send_email(
    to="recipient@example.com",
    subject="Hello from MCP Office",
    body="This is a test email"
)
print(result['success'])  # True

# Créer un rendez-vous
result = outlook.create_appointment(
    subject="Team Meeting",
    start_time="2024-01-15T10:00:00",
    end_time="2024-01-15T11:00:00",
    location="Conference Room A"
)

# Nettoyer à la fin
outlook.cleanup()
```

## Catégories de fonctionnalités

### 📧 Gestion des emails (12 méthodes)
- Créer, envoyer, répondre, transférer des emails
- Rechercher, marquer, supprimer, déplacer des emails
- Gérer les flags et les statuts de lecture

### 📎 Pièces jointes (5 méthodes)
- Ajouter, lister, sauvegarder, supprimer des pièces jointes
- Envoyer des emails avec plusieurs pièces jointes

### 📁 Gestion des dossiers (7 méthodes)
- Créer, supprimer, renommer, déplacer des dossiers
- Lister les dossiers et obtenir des statistiques

### 📅 Calendrier (10 méthodes)
- Créer, modifier, supprimer des rendez-vous
- Gérer les événements récurrents
- Rechercher des rendez-vous par date
- Définir des rappels et statuts occupés
- Exporter en format ICS

### 👥 Réunions (8 méthodes)
- Créer des demandes de réunion
- Inviter des participants
- Accepter/refuser/proposer de nouveaux horaires
- Vérifier la disponibilité

### 👤 Contacts (9 méthodes)
- Créer, modifier, supprimer, rechercher des contacts
- Gérer des groupes de contacts
- Importer/exporter des contacts

### ✅ Tâches (7 méthodes)
- Créer, modifier, supprimer des tâches
- Définir des priorités et dates d'échéance
- Marquer des tâches comme terminées

### 🔧 Opérations avancées (27 méthodes)
- Formatage des emails (HTML, importance, sensibilité)
- Catégories et organisation
- Règles et automatisation
- Signatures
- Gestion des comptes

## Exemples détaillés

### Envoyer un email avec pièce jointe

```python
result = outlook.send_with_attachments(
    to="boss@company.com",
    subject="Monthly Report",
    body="Please find attached the monthly report.",
    attachments=["report.pdf", "charts.xlsx"],
    cc="team@company.com",
    importance=2  # High importance
)
```

### Créer un événement récurrent

```python
result = outlook.create_recurring_event(
    subject="Weekly Team Sync",
    start_time="2024-01-15T10:00:00",
    end_time="2024-01-15T11:00:00",
    recurrence_type=1,  # Weekly
    interval=1,
    occurrences=52,  # Every week for a year
    location="Virtual - Teams"
)
```

### Rechercher des emails

```python
result = outlook.search_emails(
    folder_name="Inbox",
    subject="project alpha",
    sender="john@company.com",
    unread_only=True,
    max_results=20
)

for email in result['results']:
    print(f"{email['subject']} - {email['received_time']}")
```

### Créer un contact complet

```python
result = outlook.create_contact(
    first_name="Jane",
    last_name="Smith",
    email="jane.smith@example.com",
    phone="+1234567890",
    company="ABC Corporation",
    job_title="Project Manager"
)
```

### Gérer des tâches

```python
# Créer une tâche
result = outlook.create_task(
    subject="Finish quarterly report",
    body="Complete analysis and charts",
    due_date="2024-01-31T17:00:00",
    priority=2  # High priority
)

task_id = result['entry_id']

# Marquer comme terminée
result = outlook.mark_task_complete(task_id)
```

## Gestion des erreurs

Le service utilise des exceptions spécifiques pour différents types d'erreurs :

```python
from src.core.exceptions import (
    OutlookItemNotFoundError,
    InvalidRecipientError,
    AttachmentError,
    CalendarOperationError,
)

try:
    result = outlook.read_email("invalid_id")
except OutlookItemNotFoundError as e:
    print(f"Email not found: {e}")
except Exception as e:
    print(f"Unexpected error: {e}")
```

## Architecture

Le service utilise le pattern Mixin pour organiser les fonctionnalités :

```
OutlookService
├── BaseOfficeService (classe de base)
├── MailOperationsMixin (12 méthodes)
├── AttachmentOperationsMixin (5 méthodes)
├── FolderOperationsMixin (7 méthodes)
├── CalendarOperationsMixin (10 méthodes)
├── MeetingOperationsMixin (8 méthodes)
├── ContactOperationsMixin (9 méthodes)
├── TaskOperationsMixin (7 méthodes)
└── AdvancedOperationsMixin (27 méthodes)
```

## Tests

Pour exécuter les tests :

```bash
pytest tests/test_outlook_service.py -v
```

## Standards de qualité

- ✅ SOLID principles
- ✅ PEP 8 compliance
- ✅ Type hints complets
- ✅ Docstrings détaillées
- ✅ Gestion d'erreurs robuste
- ✅ Tests unitaires complets

## Prérequis

- Microsoft Outlook installé et configuré
- Python 3.8+
- `pywin32` pour l'automation COM

## Limitations

- Nécessite Windows avec Outlook installé
- L'application Outlook doit être configurée avec au moins un compte
- Certaines fonctionnalités avancées peuvent nécessiter des permissions spécifiques

## Support

Pour toute question ou problème :
1. Consulter la documentation dans les docstrings
2. Vérifier les tests pour des exemples d'utilisation
3. Consulter les exceptions pour la gestion d'erreurs

## Licence

Fait partie du projet MCP Office.
