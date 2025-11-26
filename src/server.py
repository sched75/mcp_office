"""
MCP Office Server - Complete Integration (295+ tools).

PRODUCTION-READY server with ALL Office services integrated.

Architecture:
- Word: 59 tools
- Excel: 82 tools
- PowerPoint: 63 tools
- Outlook: 67 tools
Total: 271 tools

Author: Pascal-Louis
Version: 3.0.0 - Complete Integration
"""

import asyncio
import logging
from typing import Any, Dict, Optional, Callable

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

from src.word.word_service import WordService
from src.excel.excel_service import ExcelService
from src.powerpoint.powerpoint_service import PowerPointService
from src.outlook.outlook_service import OutlookService
from src.core.exceptions import (
    COMInitializationError,
    DocumentNotFoundError,
    InvalidParameterError,
)

# Import des configurations d'outils
from tools_configs import (
    WORD_TOOLS_CONFIG,
    EXCEL_TOOLS_CONFIG,
    POWERPOINT_TOOLS_CONFIG,
)

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

# =============================================================================
# OUTLOOK TOOLS CONFIGURATION (67 tools)
# =============================================================================

OUTLOOK_TOOLS_CONFIG = {
    "send_email": {
        "required": ["to", "subject", "body"],
        "optional": ["cc", "bcc", "importance"],
        "desc": "Envoie un email via Outlook"
    },
    "send_with_attachments": {
        "required": ["to", "subject", "body", "attachments"],
        "optional": ["cc", "bcc", "importance"],
        "desc": "Envoie un email avec pièces jointes"
    },
    "read_email": {
        "required": ["email_entry_id"],
        "optional": [],
        "desc": "Lit les détails d'un email"
    },
    "reply_to_email": {
        "required": ["email_entry_id", "body"],
        "optional": ["send_immediately"],
        "desc": "Répond à un email"
    },
    "reply_all_to_email": {
        "required": ["email_entry_id", "body"],
        "optional": ["send_immediately"],
        "desc": "Répond à tous les destinataires"
    },
    "forward_email": {
        "required": ["email_entry_id", "to"],
        "optional": ["body", "send_immediately"],
        "desc": "Transfère un email"
    },
    "mark_as_read": {
        "required": ["email_entry_id"],
        "optional": [],
        "desc": "Marque un email comme lu"
    },
    "mark_as_unread": {
        "required": ["email_entry_id"],
        "optional": [],
        "desc": "Marque un email comme non lu"
    },
    "flag_email": {
        "required": ["email_entry_id"],
        "optional": ["flag_status"],
        "desc": "Ajoute un drapeau sur un email"
    },
    "delete_email": {
        "required": ["email_entry_id"],
        "optional": [],
        "desc": "Supprime un email"
    },
    "move_email_to_folder": {
        "required": ["email_entry_id", "folder_path"],
        "optional": [],
        "desc": "Déplace un email vers un dossier"
    },
    "search_emails": {
        "required": [],
        "optional": ["folder_name", "subject", "sender", "body_contains", 
                    "start_date", "end_date", "unread_only", "max_results"],
        "desc": "Recherche des emails"
    },
    "add_attachment": {
        "required": ["email_entry_id", "file_path"],
        "optional": ["display_name"],
        "desc": "Ajoute une pièce jointe"
    },
    "list_attachments": {
        "required": ["email_entry_id"],
        "optional": [],
        "desc": "Liste les pièces jointes"
    },
    "save_attachment": {
        "required": ["email_entry_id", "attachment_index", "save_path"],
        "optional": [],
        "desc": "Sauvegarde une pièce jointe"
    },
    "remove_attachment": {
        "required": ["email_entry_id", "attachment_index"],
        "optional": [],
        "desc": "Supprime une pièce jointe"
    },
    "create_new_message": {
        "required": [],
        "optional": [],
        "desc": "Crée un nouveau brouillon"
    },
    "create_folder": {
        "required": ["folder_name"],
        "optional": ["parent_folder"],
        "desc": "Crée un dossier"
    },
    "delete_folder": {
        "required": ["folder_path"],
        "optional": [],
        "desc": "Supprime un dossier"
    },
    "rename_folder": {
        "required": ["folder_path", "new_name"],
        "optional": [],
        "desc": "Renomme un dossier"
    },
    "move_folder": {
        "required": ["folder_path", "destination_path"],
        "optional": [],
        "desc": "Déplace un dossier"
    },
    "list_folders": {
        "required": [],
        "optional": ["parent_folder", "recursive"],
        "desc": "Liste les dossiers"
    },
    "get_folder_item_count": {
        "required": ["folder_path"],
        "optional": [],
        "desc": "Compte les éléments d'un dossier"
    },
    "get_unread_count": {
        "required": [],
        "optional": ["folder_path"],
        "desc": "Compte les messages non lus"
    },
    "create_appointment": {
        "required": ["subject", "start_time", "end_time"],
        "optional": ["location", "body", "reminder_minutes", "busy_status"],
        "desc": "Crée un rendez-vous"
    },
    "create_recurring_event": {
        "required": ["subject", "start_time", "end_time", "recurrence_type"],
        "optional": ["interval", "occurrences", "end_date", "location", "body"],
        "desc": "Crée un événement récurrent"
    },
    "read_appointment": {
        "required": ["appointment_entry_id"],
        "optional": [],
        "desc": "Lit un rendez-vous"
    },
    "modify_appointment": {
        "required": ["appointment_entry_id"],
        "optional": ["subject", "start_time", "end_time", "location", "body"],
        "desc": "Modifie un rendez-vous"
    },
    "delete_appointment": {
        "required": ["appointment_entry_id"],
        "optional": [],
        "desc": "Supprime un rendez-vous"
    },
    "search_appointments": {
        "required": [],
        "optional": ["subject", "location", "start_date", "end_date", "max_results"],
        "desc": "Recherche des rendez-vous"
    },
    "get_appointments_by_date": {
        "required": ["start_date", "end_date"],
        "optional": [],
        "desc": "Obtient les rendez-vous par date"
    },
    "set_reminder": {
        "required": ["appointment_entry_id", "reminder_minutes"],
        "optional": [],
        "desc": "Définit un rappel"
    },
    "set_busy_status": {
        "required": ["appointment_entry_id", "busy_status"],
        "optional": [],
        "desc": "Définit le statut occupé"
    },
    "export_appointment_ics": {
        "required": ["appointment_entry_id", "output_path"],
        "optional": [],
        "desc": "Exporte en ICS"
    },
    "get_calendar_count": {
        "required": [],
        "optional": [],
        "desc": "Compte les rendez-vous"
    },
    "export_to_pdf": {
        "required": ["output_path"],
        "optional": [],
        "desc": "Exporte le calendrier en PDF"
    },
    "create_meeting_request": {
        "required": ["subject", "start_time", "end_time", "required_attendees"],
        "optional": ["optional_attendees", "location", "body"],
        "desc": "Crée une demande de réunion"
    },
    "invite_participants": {
        "required": ["meeting_entry_id", "attendees"],
        "optional": ["required"],
        "desc": "Invite des participants"
    },
    "accept_meeting": {
        "required": ["meeting_entry_id"],
        "optional": [],
        "desc": "Accepte une réunion"
    },
    "decline_meeting": {
        "required": ["meeting_entry_id"],
        "optional": [],
        "desc": "Refuse une réunion"
    },
    "propose_new_time": {
        "required": ["meeting_entry_id", "new_start", "new_end"],
        "optional": [],
        "desc": "Propose un nouveau créneau"
    },
    "cancel_meeting": {
        "required": ["meeting_entry_id"],
        "optional": [],
        "desc": "Annule une réunion"
    },
    "update_meeting": {
        "required": ["meeting_entry_id"],
        "optional": ["subject", "start_time", "end_time", "location"],
        "desc": "Met à jour une réunion"
    },
    "check_availability": {
        "required": ["attendees", "start_time", "end_time"],
        "optional": ["duration_minutes"],
        "desc": "Vérifie la disponibilité"
    },
    "create_contact": {
        "required": ["first_name", "last_name"],
        "optional": ["email", "phone", "company", "job_title"],
        "desc": "Crée un contact"
    },
    "modify_contact": {
        "required": ["contact_entry_id"],
        "optional": ["first_name", "last_name", "email", "phone"],
        "desc": "Modifie un contact"
    },
    "delete_contact": {
        "required": ["contact_entry_id"],
        "optional": [],
        "desc": "Supprime un contact"
    },
    "search_contact": {
        "required": ["search_term"],
        "optional": [],
        "desc": "Recherche un contact"
    },
    "list_all_contacts": {
        "required": [],
        "optional": [],
        "desc": "Liste tous les contacts"
    },
    "create_contact_group": {
        "required": ["group_name"],
        "optional": [],
        "desc": "Crée un groupe de contacts"
    },
    "add_to_contact_group": {
        "required": ["group_entry_id", "contact_email"],
        "optional": [],
        "desc": "Ajoute à un groupe"
    },
    "export_contacts_vcf": {
        "required": ["output_path"],
        "optional": [],
        "desc": "Exporte les contacts en VCF"
    },
    "import_contacts": {
        "required": ["file_path"],
        "optional": [],
        "desc": "Importe des contacts"
    },
    "create_task": {
        "required": ["subject"],
        "optional": ["body", "due_date", "priority"],
        "desc": "Crée une tâche"
    },
    "modify_task": {
        "required": ["task_entry_id"],
        "optional": ["subject", "body", "due_date"],
        "desc": "Modifie une tâche"
    },
    "delete_task": {
        "required": ["task_entry_id"],
        "optional": [],
        "desc": "Supprime une tâche"
    },
    "mark_task_complete": {
        "required": ["task_entry_id"],
        "optional": [],
        "desc": "Marque une tâche terminée"
    },
    "set_task_priority": {
        "required": ["task_entry_id", "priority"],
        "optional": [],
        "desc": "Définit la priorité"
    },
    "set_task_due_date": {
        "required": ["task_entry_id", "due_date"],
        "optional": [],
        "desc": "Définit l'échéance"
    },
    "list_tasks": {
        "required": [],
        "optional": ["completed"],
        "desc": "Liste les tâches"
    },
    "list_accounts": {
        "required": [],
        "optional": [],
        "desc": "Liste les comptes"
    },
    "get_default_account": {
        "required": [],
        "optional": [],
        "desc": "Obtient le compte par défaut"
    },
    "get_inbox_count": {
        "required": [],
        "optional": [],
        "desc": "Compte les messages inbox"
    },
    "create_category": {
        "required": ["name"],
        "optional": ["color"],
        "desc": "Crée une catégorie"
    },
    "list_categories": {
        "required": [],
        "optional": [],
        "desc": "Liste les catégories"
    },
    "apply_category": {
        "required": ["item_entry_id", "category"],
        "optional": [],
        "desc": "Applique une catégorie"
    },
    "com_operation": {
        "required": ["operation_name"],
        "optional": [],
        "desc": "Opération COM personnalisée"
    }
}

# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

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


def generate_tool(service_prefix: str, name: str, config: dict) -> Tool:
    """Génère un outil MCP à partir de sa configuration."""
    properties = {}
    
    # Ajouter tous les paramètres (requis et optionnels)
    for param in config.get("required", []) + config.get("optional", []):
        properties[param] = {"type": "string"}  # Type par défaut
    
    return Tool(
        name=f"{service_prefix}_{name}",
        description=config["desc"],
        inputSchema={
            "type": "object",
            "properties": properties,
            "required": config.get("required", []),
        },
    )


def build_handlers(service: Any, service_config: Dict, service_prefix: str) -> Dict[str, Callable]:
    """Construit le mapping des handlers de manière dynamique."""
    
    def create_handler(method_name: str, config: dict):
        """Crée un handler dynamique pour une méthode."""
        method = getattr(service, method_name)
        
        def handler(args: dict):
            # Extraire tous les arguments (requis + optionnels)
            kwargs = {}
            for param in config.get("required", []) + config.get("optional", []):
                if param in args:
                    kwargs[param] = args[param]
            return method(**kwargs)
        
        return handler
    
    # Générer tous les handlers automatiquement
    handlers = {}
    for method_name, config in service_config.items():
        tool_name = f"{service_prefix}_{method_name}"
        handlers[tool_name] = create_handler(method_name, config)
    
    return handlers


# =============================================================================
# MCP SERVER HANDLERS
# =============================================================================

@app.list_tools()
async def list_tools() -> list[Tool]:
    """Liste tous les outils disponibles (271 outils)."""
    tools = []
    
    # Générer les outils Word
    for method_name, config in WORD_TOOLS_CONFIG.items():
        tools.append(generate_tool("word", method_name, config))
    
    # Générer les outils Excel
    for method_name, config in EXCEL_TOOLS_CONFIG.items():
        tools.append(generate_tool("excel", method_name, config))
    
    # Générer les outils PowerPoint
    for method_name, config in POWERPOINT_TOOLS_CONFIG.items():
        tools.append(generate_tool("powerpoint", method_name, config))
    
    # Générer les outils Outlook
    for method_name, config in OUTLOOK_TOOLS_CONFIG.items():
        tools.append(generate_tool("outlook", method_name, config))
    
    logger.info(f"Loaded {len(tools)} tools total")
    logger.info(f"  - Word: {len(WORD_TOOLS_CONFIG)} tools")
    logger.info(f"  - Excel: {len(EXCEL_TOOLS_CONFIG)} tools")
    logger.info(f"  - PowerPoint: {len(POWERPOINT_TOOLS_CONFIG)} tools")
    logger.info(f"  - Outlook: {len(OUTLOOK_TOOLS_CONFIG)} tools")
    
    return tools


@app.call_tool()
async def call_tool(name: str, arguments: Any) -> list[TextContent]:
    """Exécute un outil MCP avec routing automatique."""
    logger.info(f"Calling tool: {name}")
    
    try:
        # Convertir arguments en dictionnaire
        if not isinstance(arguments, dict):
            arguments = {}
        
        result = None
        
        # === WORD TOOLS ===
        if name.startswith("word_"):
            if word_service is None:
                raise COMInitializationError("Word service not initialized")
            
            handlers = build_handlers(word_service, WORD_TOOLS_CONFIG, "word")
            
            if name in handlers:
                result = handlers[name](arguments)
            else:
                return [TextContent(
                    type="text",
                    text=f"❌ Outil Word non implémenté: {name}"
                )]
        
        # === EXCEL TOOLS ===
        elif name.startswith("excel_"):
            if excel_service is None:
                raise COMInitializationError("Excel service not initialized")
            
            handlers = build_handlers(excel_service, EXCEL_TOOLS_CONFIG, "excel")
            
            if name in handlers:
                result = handlers[name](arguments)
            else:
                return [TextContent(
                    type="text",
                    text=f"❌ Outil Excel non implémenté: {name}"
                )]
        
        # === POWERPOINT TOOLS ===
        elif name.startswith("powerpoint_"):
            if powerpoint_service is None:
                raise COMInitializationError("PowerPoint service not initialized")
            
            handlers = build_handlers(powerpoint_service, POWERPOINT_TOOLS_CONFIG, "powerpoint")
            
            if name in handlers:
                result = handlers[name](arguments)
            else:
                return [TextContent(
                    type="text",
                    text=f"❌ Outil PowerPoint non implémenté: {name}"
                )]
        
        # === OUTLOOK TOOLS ===
        elif name.startswith("outlook_"):
            if outlook_service is None:
                raise COMInitializationError("Outlook service not initialized")
            
            handlers = build_handlers(outlook_service, OUTLOOK_TOOLS_CONFIG, "outlook")
            
            if name in handlers:
                result = handlers[name](arguments)
            else:
                return [TextContent(
                    type="text",
                    text=f"❌ Outil Outlook non implémenté: {name}"
                )]
        
        else:
            return [TextContent(type="text", text=f"❌ Outil inconnu: {name}")]
        
        # Formater et retourner le résultat
        if result is None:
            return [TextContent(type="text", text=f"❌ Aucun résultat retourné")]
        
        formatted = format_result(result)
        return [TextContent(type="text", text=formatted)]
    
    except InvalidParameterError as e:
        logger.error(f"Invalid parameters for {name}: {e}")
        return [TextContent(type="text", text=f"❌ Paramètres invalides: {e}")]
    
    except DocumentNotFoundError as e:
        logger.error(f"Document not found: {e}")
        return [TextContent(type="text", text=f"❌ Document non trouvé: {e}")]
    
    except COMInitializationError as e:
        logger.error(f"COM initialization error: {e}")
        return [TextContent(type="text", text=f"❌ Erreur d'initialisation: {e}")]
    
    except Exception as e:
        logger.exception(f"Error calling tool {name}")
        return [TextContent(type="text", text=f"❌ Erreur inattendue: {str(e)}")]


# =============================================================================
# LIFECYCLE MANAGEMENT
# =============================================================================

async def initialize_services():
    """Initialise tous les services Office."""
    global word_service, excel_service, powerpoint_service, outlook_service
    
    logger.info("Initializing Office services...")
    
    try:
        # Initialiser Word
        word_service = WordService()
        word_service.initialize()
        logger.info(f"✅ Word service initialized ({len(WORD_TOOLS_CONFIG)} tools)")
        
        # Initialiser Excel
        excel_service = ExcelService()
        excel_service.initialize()
        logger.info(f"✅ Excel service initialized ({len(EXCEL_TOOLS_CONFIG)} tools)")
        
        # Initialiser PowerPoint
        powerpoint_service = PowerPointService()
        powerpoint_service.initialize()
        logger.info(f"✅ PowerPoint service initialized ({len(POWERPOINT_TOOLS_CONFIG)} tools)")
        
        # Initialiser Outlook
        outlook_service = OutlookService()
        outlook_service.initialize()
        logger.info(f"✅ Outlook service initialized ({len(OUTLOOK_TOOLS_CONFIG)} tools)")
        
        total_tools = (len(WORD_TOOLS_CONFIG) + len(EXCEL_TOOLS_CONFIG) + 
                      len(POWERPOINT_TOOLS_CONFIG) + len(OUTLOOK_TOOLS_CONFIG))
        
        logger.info("🚀 All Office services ready!")
        logger.info(f"📊 Total tools available: {total_tools}")
    
    except Exception as e:
        logger.error(f"Failed to initialize services: {e}")
        raise


async def cleanup_services():
    """Nettoie tous les services Office."""
    logger.info("Cleaning up Office services...")
    
    try:
        if word_service:
            word_service.cleanup()
        if excel_service:
            excel_service.cleanup()
        if powerpoint_service:
            powerpoint_service.cleanup()
        if outlook_service:
            outlook_service.cleanup()
        
        logger.info("✅ All services cleaned up")
    except Exception as e:
        logger.error(f"Error during cleanup: {e}")


async def main():
    """Point d'entrée principal du serveur MCP."""
    logger.info("Starting MCP Office Server...")
    logger.info("=" * 80)
    logger.info("MCP Office - Complete Office Automation Server")
    logger.info("=" * 80)
    
    try:
        # Initialiser les services
        await initialize_services()
        
        # Démarrer le serveur MCP
        async with stdio_server() as (read_stream, write_stream):
            await app.run(
                read_stream,
                write_stream,
                app.create_initialization_options(),
            )
    
    except KeyboardInterrupt:
        logger.info("Server interrupted by user")
    except Exception as e:
        logger.exception("Fatal error in server")
        raise
    finally:
        # Nettoyer les services
        await cleanup_services()
        logger.info("Server stopped")


if __name__ == "__main__":
    asyncio.run(main())
