"""
MCP Office Server - Complete integration with all 295+ tools.

This server exposes all Office automation capabilities:
- Word: 65 tools
- Excel: 82 tools
- PowerPoint: 63 tools
- Outlook: 85+ tools
Total: 295+ tools

Author: Pascal-Louis
Version: 2.0.0 - Complete Outlook Integration
"""

import logging
from typing import Any, Dict, Optional

from mcp.server import Server
from mcp.types import Tool

from src.core.exceptions import (
    InvalidParameterError,
)
from src.excel.excel_service import ExcelService
from src.outlook.outlook_service import OutlookService
from src.powerpoint.powerpoint_service import PowerPointService
from src.word.word_service import WordService

# Configuration du logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger("mcp_office")

# Initialisation du serveur MCP
app = Server("mcp-office-server")

# Services Office
word_service: Optional[WordService] = None
excel_service: Optional[ExcelService] = None
powerpoint_service: Optional[PowerPointService] = None
outlook_service: Optional[OutlookService] = None


def format_result(result: Dict[str, Any]) -> str:
    """Formate un résultat pour l'affichage."""
    if not result.get("success", False):
        error_msg = result.get("error", "Unknown error")
        return f"❌ Erreur: {error_msg}"

    lines = ["✅ Opération réussie"]
    for key, value in result.items():
        if key not in ["success", "error"] and value is not None:
            if isinstance(value, (list, dict)) and len(str(value)) > 200:
                lines.append(f"  • {key}: [données volumineuses - {len(value)} éléments]")
            else:
                lines.append(f"  • {key}: {value}")

    return "\n".join(lines)


def validate_parameters(params: Dict[str, Any], required: list[str]) -> None:
    """Valide la présence des paramètres requis."""
    missing = [p for p in required if p not in params or params[p] is None]
    if missing:
        raise InvalidParameterError(f"Paramètres manquants: {', '.join(missing)}")


# ============================================================================
# TOOL DEFINITIONS - OUTLOOK (85 TOOLS)
# ============================================================================


def get_outlook_tools() -> list[Tool]:
    """Retourne tous les outils Outlook (85 outils)."""
    tools = []

    # ========================================================================
    # MAIL OPERATIONS (12 outils)
    # ========================================================================

    tools.extend(
        [
            Tool(
                name="outlook_send_email",
                description="Envoie un email via Outlook",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "to": {"type": "string", "description": "Destinataire(s), séparés par ;"},
                        "subject": {"type": "string", "description": "Objet de l'email"},
                        "body": {"type": "string", "description": "Corps du message"},
                        "cc": {"type": "string", "description": "Copie à (CC)"},
                        "bcc": {"type": "string", "description": "Copie cachée (BCC)"},
                        "importance": {
                            "type": "integer",
                            "description": "Importance: 0=basse, 1=normale, 2=haute",
                            "default": 1,
                        },
                    },
                    "required": ["to", "subject", "body"],
                },
            ),
            Tool(
                name="outlook_send_with_attachments",
                description="Envoie un email avec pièces jointes",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "to": {"type": "string"},
                        "subject": {"type": "string"},
                        "body": {"type": "string"},
                        "attachments": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": "Liste des chemins de fichiers",
                        },
                        "cc": {"type": "string"},
                        "bcc": {"type": "string"},
                        "importance": {"type": "integer"},
                    },
                    "required": ["to", "subject", "body", "attachments"],
                },
            ),
            Tool(
                name="outlook_read_email",
                description="Lit les détails d'un email",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "email_entry_id": {
                            "type": "string",
                            "description": "ID d'entrée de l'email",
                        },
                    },
                    "required": ["email_entry_id"],
                },
            ),
            Tool(
                name="outlook_reply_to_email",
                description="Répond à un email",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "email_entry_id": {"type": "string"},
                        "body": {"type": "string", "description": "Corps de la réponse"},
                        "send_immediately": {
                            "type": "boolean",
                            "description": "Envoyer immédiatement (sinon brouillon)",
                            "default": True,
                        },
                    },
                    "required": ["email_entry_id", "body"],
                },
            ),
            Tool(
                name="outlook_reply_all_to_email",
                description="Répond à tous les destinataires d'un email",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "email_entry_id": {"type": "string"},
                        "body": {"type": "string"},
                        "send_immediately": {"type": "boolean", "default": True},
                    },
                    "required": ["email_entry_id", "body"],
                },
            ),
            Tool(
                name="outlook_forward_email",
                description="Transfère un email",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "email_entry_id": {"type": "string"},
                        "to": {"type": "string", "description": "Destinataires"},
                        "body": {"type": "string", "description": "Message additionnel"},
                        "send_immediately": {"type": "boolean", "default": True},
                    },
                    "required": ["email_entry_id", "to"],
                },
            ),
            Tool(
                name="outlook_mark_as_read",
                description="Marque un email comme lu",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "email_entry_id": {"type": "string"},
                    },
                    "required": ["email_entry_id"],
                },
            ),
            Tool(
                name="outlook_mark_as_unread",
                description="Marque un email comme non lu",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "email_entry_id": {"type": "string"},
                    },
                    "required": ["email_entry_id"],
                },
            ),
            Tool(
                name="outlook_flag_email",
                description="Ajoute/modifie un drapeau sur un email",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "email_entry_id": {"type": "string"},
                        "flag_status": {
                            "type": "integer",
                            "description": "0=non marqué, 1=terminé, 2=marqué",
                            "default": 2,
                        },
                    },
                    "required": ["email_entry_id"],
                },
            ),
            Tool(
                name="outlook_delete_email",
                description="Supprime un email (déplace vers Éléments supprimés)",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "email_entry_id": {"type": "string"},
                    },
                    "required": ["email_entry_id"],
                },
            ),
            Tool(
                name="outlook_move_email_to_folder",
                description="Déplace un email vers un dossier",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "email_entry_id": {"type": "string"},
                        "folder_path": {
                            "type": "string",
                            "description": "Chemin du dossier (ex: 'Inbox/Archive')",
                        },
                    },
                    "required": ["email_entry_id", "folder_path"],
                },
            ),
            Tool(
                name="outlook_search_emails",
                description="Recherche des emails dans Outlook",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "folder_name": {
                            "type": "string",
                            "description": "Nom du dossier",
                            "default": "Inbox",
                        },
                        "subject": {"type": "string", "description": "Filtrer par objet"},
                        "sender": {"type": "string", "description": "Filtrer par expéditeur"},
                        "body_contains": {"type": "string", "description": "Filtrer par contenu"},
                        "start_date": {"type": "string", "description": "Date de début (ISO)"},
                        "end_date": {"type": "string", "description": "Date de fin (ISO)"},
                        "unread_only": {"type": "boolean", "description": "Seulement non lus"},
                        "max_results": {
                            "type": "integer",
                            "description": "Nombre max de résultats",
                            "default": 50,
                        },
                    },
                },
            ),
        ]
    )

    # ========================================================================
    # ATTACHMENT OPERATIONS (5 outils)
    # ========================================================================

    tools.extend(
        [
            Tool(
                name="outlook_add_attachment",
                description="Ajoute une pièce jointe à un email",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "email_entry_id": {"type": "string"},
                        "file_path": {"type": "string", "description": "Chemin du fichier"},
                        "display_name": {
                            "type": "string",
                            "description": "Nom d'affichage (optionnel)",
                        },
                    },
                    "required": ["email_entry_id", "file_path"],
                },
            ),
            Tool(
                name="outlook_list_attachments",
                description="Liste les pièces jointes d'un email",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "email_entry_id": {"type": "string"},
                    },
                    "required": ["email_entry_id"],
                },
            ),
            Tool(
                name="outlook_save_attachment",
                description="Sauvegarde une pièce jointe",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "email_entry_id": {"type": "string"},
                        "attachment_index": {
                            "type": "integer",
                            "description": "Index de la pièce jointe (1-based)",
                        },
                        "save_path": {"type": "string", "description": "Chemin de destination"},
                    },
                    "required": ["email_entry_id", "attachment_index", "save_path"],
                },
            ),
            Tool(
                name="outlook_remove_attachment",
                description="Supprime une pièce jointe d'un email",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "email_entry_id": {"type": "string"},
                        "attachment_index": {"type": "integer", "description": "Index (1-based)"},
                    },
                    "required": ["email_entry_id", "attachment_index"],
                },
            ),
            Tool(
                name="outlook_create_new_message",
                description="Crée un nouveau brouillon d'email",
                inputSchema={
                    "type": "object",
                    "properties": {},
                },
            ),
        ]
    )

    # ========================================================================
    # FOLDER OPERATIONS (7 outils)
    # ========================================================================

    tools.extend(
        [
            Tool(
                name="outlook_create_folder",
                description="Crée un nouveau dossier Outlook",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "folder_name": {"type": "string"},
                        "parent_folder": {
                            "type": "string",
                            "description": "Dossier parent",
                            "default": "Inbox",
                        },
                    },
                    "required": ["folder_name"],
                },
            ),
            Tool(
                name="outlook_delete_folder",
                description="Supprime un dossier Outlook",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "folder_path": {
                            "type": "string",
                            "description": "Chemin du dossier (ex: 'Inbox/Archive')",
                        },
                    },
                    "required": ["folder_path"],
                },
            ),
            Tool(
                name="outlook_rename_folder",
                description="Renomme un dossier Outlook",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "folder_path": {"type": "string"},
                        "new_name": {"type": "string"},
                    },
                    "required": ["folder_path", "new_name"],
                },
            ),
            Tool(
                name="outlook_move_folder",
                description="Déplace un dossier vers un autre emplacement",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "folder_path": {"type": "string"},
                        "destination_path": {"type": "string"},
                    },
                    "required": ["folder_path", "destination_path"],
                },
            ),
            Tool(
                name="outlook_list_folders",
                description="Liste les dossiers Outlook",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "parent_folder": {
                            "type": "string",
                            "description": "Dossier parent",
                            "default": "Inbox",
                        },
                        "recursive": {
                            "type": "boolean",
                            "description": "Inclure les sous-dossiers",
                            "default": False,
                        },
                    },
                },
            ),
            Tool(
                name="outlook_get_folder_item_count",
                description="Obtient le nombre d'éléments dans un dossier",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "folder_path": {"type": "string"},
                    },
                    "required": ["folder_path"],
                },
            ),
            Tool(
                name="outlook_get_unread_count",
                description="Obtient le nombre de messages non lus dans un dossier",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "folder_path": {"type": "string", "default": "Inbox"},
                    },
                },
            ),
        ]
    )

    # ========================================================================
    # CALENDAR OPERATIONS (12 outils)
    # ========================================================================

    tools.extend(
        [
            Tool(
                name="outlook_create_appointment",
                description="Crée un rendez-vous dans le calendrier",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "subject": {"type": "string"},
                        "start_time": {"type": "string", "description": "Format ISO"},
                        "end_time": {"type": "string", "description": "Format ISO"},
                        "location": {"type": "string"},
                        "body": {"type": "string"},
                        "reminder_minutes": {
                            "type": "integer",
                            "description": "Minutes avant rappel",
                            "default": 15,
                        },
                        "busy_status": {
                            "type": "integer",
                            "description": "0=libre, 1=tentative, 2=occupé, 3=absent",
                            "default": 2,
                        },
                    },
                    "required": ["subject", "start_time", "end_time"],
                },
            ),
            Tool(
                name="outlook_create_recurring_event",
                description="Crée un événement récurrent",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "subject": {"type": "string"},
                        "start_time": {"type": "string"},
                        "end_time": {"type": "string"},
                        "recurrence_type": {
                            "type": "integer",
                            "description": "0=daily, 1=weekly, 2=monthly, 3=yearly",
                        },
                        "interval": {"type": "integer", "description": "Intervalle de récurrence"},
                        "occurrences": {
                            "type": "integer",
                            "description": "Nombre d'occurrences (optionnel)",
                        },
                        "end_date": {
                            "type": "string",
                            "description": "Date de fin de récurrence (optionnel)",
                        },
                        "location": {"type": "string"},
                        "body": {"type": "string"},
                    },
                    "required": ["subject", "start_time", "end_time", "recurrence_type"],
                },
            ),
            Tool(
                name="outlook_read_appointment",
                description="Lit les détails d'un rendez-vous",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "appointment_entry_id": {"type": "string"},
                    },
                    "required": ["appointment_entry_id"],
                },
            ),
            Tool(
                name="outlook_modify_appointment",
                description="Modifie un rendez-vous existant",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "appointment_entry_id": {"type": "string"},
                        "subject": {"type": "string"},
                        "start_time": {"type": "string"},
                        "end_time": {"type": "string"},
                        "location": {"type": "string"},
                        "body": {"type": "string"},
                    },
                    "required": ["appointment_entry_id"],
                },
            ),
            Tool(
                name="outlook_delete_appointment",
                description="Supprime un rendez-vous",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "appointment_entry_id": {"type": "string"},
                    },
                    "required": ["appointment_entry_id"],
                },
            ),
            Tool(
                name="outlook_search_appointments",
                description="Recherche des rendez-vous",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "subject": {"type": "string"},
                        "location": {"type": "string"},
                        "start_date": {"type": "string"},
                        "end_date": {"type": "string"},
                        "max_results": {"type": "integer", "default": 50},
                    },
                },
            ),
            Tool(
                name="outlook_get_appointments_by_date",
                description="Obtient les rendez-vous pour une période",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "start_date": {"type": "string"},
                        "end_date": {"type": "string"},
                    },
                    "required": ["start_date", "end_date"],
                },
            ),
            Tool(
                name="outlook_set_reminder",
                description="Définit un rappel pour un rendez-vous",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "appointment_entry_id": {"type": "string"},
                        "reminder_minutes": {"type": "integer"},
                    },
                    "required": ["appointment_entry_id", "reminder_minutes"],
                },
            ),
            Tool(
                name="outlook_set_busy_status",
                description="Définit le statut occupé d'un rendez-vous",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "appointment_entry_id": {"type": "string"},
                        "busy_status": {
                            "type": "integer",
                            "description": "0=libre, 1=tentative, 2=occupé, 3=absent",
                        },
                    },
                    "required": ["appointment_entry_id", "busy_status"],
                },
            ),
            Tool(
                name="outlook_export_appointment_ics",
                description="Exporte un rendez-vous vers un fichier ICS",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "appointment_entry_id": {"type": "string"},
                        "output_path": {"type": "string"},
                    },
                    "required": ["appointment_entry_id", "output_path"],
                },
            ),
            Tool(
                name="outlook_get_calendar_count",
                description="Obtient le nombre de rendez-vous dans le calendrier",
                inputSchema={
                    "type": "object",
                    "properties": {},
                },
            ),
            Tool(
                name="outlook_export_to_pdf",
                description="Exporte le calendrier vers PDF",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "output_path": {"type": "string"},
                    },
                    "required": ["output_path"],
                },
            ),
        ]
    )

    # ========================================================================
    # MEETING OPERATIONS (8 outils)
    # ========================================================================

    tools.extend(
        [
            Tool(
                name="outlook_create_meeting_request",
                description="Crée une demande de réunion",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "subject": {"type": "string"},
                        "start_time": {"type": "string"},
                        "end_time": {"type": "string"},
                        "required_attendees": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": "Liste des participants requis",
                        },
                        "optional_attendees": {
                            "type": "array",
                            "items": {"type": "string"},
                            "description": "Liste des participants optionnels",
                        },
                        "location": {"type": "string"},
                        "body": {"type": "string"},
                    },
                    "required": ["subject", "start_time", "end_time", "required_attendees"],
                },
            ),
            Tool(
                name="outlook_invite_participants",
                description="Invite des participants à une réunion",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "meeting_entry_id": {"type": "string"},
                        "attendees": {
                            "type": "array",
                            "items": {"type": "string"},
                        },
                        "required": {
                            "type": "boolean",
                            "description": "Participants requis (vs optionnels)",
                            "default": True,
                        },
                    },
                    "required": ["meeting_entry_id", "attendees"],
                },
            ),
            Tool(
                name="outlook_accept_meeting",
                description="Accepte une invitation à une réunion",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "meeting_entry_id": {"type": "string"},
                    },
                    "required": ["meeting_entry_id"],
                },
            ),
            Tool(
                name="outlook_decline_meeting",
                description="Refuse une invitation à une réunion",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "meeting_entry_id": {"type": "string"},
                    },
                    "required": ["meeting_entry_id"],
                },
            ),
            Tool(
                name="outlook_propose_new_time",
                description="Propose un nouveau créneau pour une réunion",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "meeting_entry_id": {"type": "string"},
                        "new_start": {"type": "string"},
                        "new_end": {"type": "string"},
                    },
                    "required": ["meeting_entry_id", "new_start", "new_end"],
                },
            ),
            Tool(
                name="outlook_cancel_meeting",
                description="Annule une réunion",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "meeting_entry_id": {"type": "string"},
                    },
                    "required": ["meeting_entry_id"],
                },
            ),
            Tool(
                name="outlook_update_meeting",
                description="Met à jour une réunion existante",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "meeting_entry_id": {"type": "string"},
                        "subject": {"type": "string"},
                        "start_time": {"type": "string"},
                        "end_time": {"type": "string"},
                        "location": {"type": "string"},
                    },
                    "required": ["meeting_entry_id"],
                },
            ),
            Tool(
                name="outlook_check_availability",
                description="Vérifie la disponibilité des participants",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "attendees": {
                            "type": "array",
                            "items": {"type": "string"},
                        },
                        "start_time": {"type": "string"},
                        "end_time": {"type": "string"},
                        "duration_minutes": {"type": "integer"},
                    },
                    "required": ["attendees", "start_time", "end_time"],
                },
            ),
        ]
    )

    # ========================================================================
    # CONTACT OPERATIONS (9 outils)
    # ========================================================================

    tools.extend(
        [
            Tool(
                name="outlook_create_contact",
                description="Crée un nouveau contact",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "first_name": {"type": "string"},
                        "last_name": {"type": "string"},
                        "email": {"type": "string"},
                        "phone": {"type": "string"},
                        "company": {"type": "string"},
                        "job_title": {"type": "string"},
                    },
                    "required": ["first_name", "last_name"],
                },
            ),
            Tool(
                name="outlook_modify_contact",
                description="Modifie un contact existant",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "contact_entry_id": {"type": "string"},
                        "first_name": {"type": "string"},
                        "last_name": {"type": "string"},
                        "email": {"type": "string"},
                        "phone": {"type": "string"},
                    },
                    "required": ["contact_entry_id"],
                },
            ),
            Tool(
                name="outlook_delete_contact",
                description="Supprime un contact",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "contact_entry_id": {"type": "string"},
                    },
                    "required": ["contact_entry_id"],
                },
            ),
            Tool(
                name="outlook_search_contact",
                description="Recherche un contact",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "search_term": {
                            "type": "string",
                            "description": "Terme de recherche (nom, email, etc.)",
                        },
                    },
                    "required": ["search_term"],
                },
            ),
            Tool(
                name="outlook_list_all_contacts",
                description="Liste tous les contacts",
                inputSchema={
                    "type": "object",
                    "properties": {},
                },
            ),
            Tool(
                name="outlook_create_contact_group",
                description="Crée un groupe de contacts (liste de distribution)",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "group_name": {"type": "string"},
                    },
                    "required": ["group_name"],
                },
            ),
            Tool(
                name="outlook_add_to_contact_group",
                description="Ajoute un contact à un groupe",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "group_entry_id": {"type": "string"},
                        "contact_email": {"type": "string"},
                    },
                    "required": ["group_entry_id", "contact_email"],
                },
            ),
            Tool(
                name="outlook_export_contacts_vcf",
                description="Exporte les contacts vers VCF",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "output_path": {"type": "string"},
                    },
                    "required": ["output_path"],
                },
            ),
            Tool(
                name="outlook_import_contacts",
                description="Importe des contacts depuis un fichier",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {"type": "string"},
                    },
                    "required": ["file_path"],
                },
            ),
        ]
    )

    # ========================================================================
    # TASK OPERATIONS (7 outils)
    # ========================================================================

    tools.extend(
        [
            Tool(
                name="outlook_create_task",
                description="Crée une nouvelle tâche",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "subject": {"type": "string"},
                        "body": {"type": "string"},
                        "due_date": {"type": "string", "description": "Date d'échéance (ISO)"},
                        "priority": {
                            "type": "integer",
                            "description": "0=basse, 1=normale, 2=haute",
                            "default": 1,
                        },
                    },
                    "required": ["subject"],
                },
            ),
            Tool(
                name="outlook_modify_task",
                description="Modifie une tâche existante",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "task_entry_id": {"type": "string"},
                        "subject": {"type": "string"},
                        "body": {"type": "string"},
                        "due_date": {"type": "string"},
                    },
                    "required": ["task_entry_id"],
                },
            ),
            Tool(
                name="outlook_delete_task",
                description="Supprime une tâche",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "task_entry_id": {"type": "string"},
                    },
                    "required": ["task_entry_id"],
                },
            ),
            Tool(
                name="outlook_mark_task_complete",
                description="Marque une tâche comme terminée",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "task_entry_id": {"type": "string"},
                    },
                    "required": ["task_entry_id"],
                },
            ),
            Tool(
                name="outlook_set_task_priority",
                description="Définit la priorité d'une tâche",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "task_entry_id": {"type": "string"},
                        "priority": {
                            "type": "integer",
                            "description": "0=basse, 1=normale, 2=haute",
                        },
                    },
                    "required": ["task_entry_id", "priority"],
                },
            ),
            Tool(
                name="outlook_set_task_due_date",
                description="Définit la date d'échéance d'une tâche",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "task_entry_id": {"type": "string"},
                        "due_date": {"type": "string"},
                    },
                    "required": ["task_entry_id", "due_date"],
                },
            ),
            Tool(
                name="outlook_list_tasks",
                description="Liste les tâches",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "completed": {
                            "type": "boolean",
                            "description": "Inclure les tâches terminées",
                            "default": False,
                        },
                    },
                },
            ),
        ]
    )

    # ========================================================================
    # ADVANCED OPERATIONS (10 outils)
    # ========================================================================

    tools.extend(
        [
            Tool(
                name="outlook_list_accounts",
                description="Liste tous les comptes Outlook configurés",
                inputSchema={"type": "object", "properties": {}},
            ),
            Tool(
                name="outlook_get_default_account",
                description="Obtient le compte par défaut",
                inputSchema={"type": "object", "properties": {}},
            ),
            Tool(
                name="outlook_get_inbox_count",
                description="Obtient le nombre de messages dans la boîte de réception",
                inputSchema={"type": "object", "properties": {}},
            ),
            Tool(
                name="outlook_create_category",
                description="Crée une catégorie Outlook",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "name": {"type": "string"},
                        "color": {
                            "type": "integer",
                            "description": "Index de couleur (0-24)",
                            "default": 0,
                        },
                    },
                    "required": ["name"],
                },
            ),
            Tool(
                name="outlook_list_categories",
                description="Liste toutes les catégories",
                inputSchema={"type": "object", "properties": {}},
            ),
            Tool(
                name="outlook_apply_category",
                description="Applique une catégorie à un élément",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "item_entry_id": {"type": "string"},
                        "category": {"type": "string", "description": "Nom de la catégorie"},
                    },
                    "required": ["item_entry_id", "category"],
                },
            ),
            Tool(
                name="outlook_com_operation",
                description="Exécute une opération COM personnalisée (avancé)",
                inputSchema={
                    "type": "object",
                    "properties": {
                        "operation_name": {"type": "string"},
                    },
                    "required": ["operation_name"],
                },
            ),
        ]
    )

    logger.info(f"Generated {len(tools)} Outlook tools")
    return tools


# ============================================================================
# CONTINUATION FILE MARKER - FILE IS TOO LARGE
# This will be split into server.py (part 1) and server_handlers.py (part 2)
# ============================================================================
