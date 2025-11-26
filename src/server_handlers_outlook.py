"""
MCP Office Server - Complete Outlook Integration (295+ tools).

Modular architecture with handler mapping for maintainability.

Author: Pascal-Louis
Version: 2.0.0 - Complete Outlook Integration with 85+ handlers
"""

import logging
from typing import Any, Callable, Dict, Optional

from mcp.server import Server

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


# ===========================================================================
# OUTLOOK HANDLER MAPPING
# ===========================================================================
# This mapping provides a clean, maintainable way to route tool calls
# to the appropriate OutlookService methods
# ===========================================================================


def build_outlook_handler_map(service: OutlookService) -> Dict[str, Callable]:
    """Construit le mapping des handlers Outlook."""
    return {
        # Mail Operations (12)
        "outlook_send_email": lambda args: service.send_email(
            to=args["to"],
            subject=args["subject"],
            body=args["body"],
            cc=args.get("cc"),
            bcc=args.get("bcc"),
            importance=args.get("importance", 1),
        ),
        "outlook_send_with_attachments": lambda args: service.send_with_attachments(
            to=args["to"],
            subject=args["subject"],
            body=args["body"],
            attachments=args["attachments"],
            cc=args.get("cc"),
            bcc=args.get("bcc"),
            importance=args.get("importance", 1),
        ),
        "outlook_read_email": lambda args: service.read_email(args["email_entry_id"]),
        "outlook_reply_to_email": lambda args: service.reply_to_email(
            email_entry_id=args["email_entry_id"],
            body=args["body"],
            send_immediately=args.get("send_immediately", True),
        ),
        "outlook_reply_all_to_email": lambda args: service.reply_all_to_email(
            email_entry_id=args["email_entry_id"],
            body=args["body"],
            send_immediately=args.get("send_immediately", True),
        ),
        "outlook_forward_email": lambda args: service.forward_email(
            email_entry_id=args["email_entry_id"],
            to=args["to"],
            body=args.get("body"),
            send_immediately=args.get("send_immediately", True),
        ),
        "outlook_mark_as_read": lambda args: service.mark_as_read(args["email_entry_id"]),
        "outlook_mark_as_unread": lambda args: service.mark_as_unread(args["email_entry_id"]),
        "outlook_flag_email": lambda args: service.flag_email(
            email_entry_id=args["email_entry_id"],
            flag_status=args.get("flag_status", 2),
        ),
        "outlook_delete_email": lambda args: service.delete_email(args["email_entry_id"]),
        "outlook_move_email_to_folder": lambda args: service.move_email_to_folder(
            email_entry_id=args["email_entry_id"],
            folder_path=args["folder_path"],
        ),
        "outlook_search_emails": lambda args: service.search_emails(
            folder_name=args.get("folder_name", "Inbox"),
            subject=args.get("subject"),
            sender=args.get("sender"),
            body_contains=args.get("body_contains"),
            start_date=args.get("start_date"),
            end_date=args.get("end_date"),
            unread_only=args.get("unread_only", False),
            max_results=args.get("max_results", 50),
        ),
        # Attachment Operations (5)
        "outlook_add_attachment": lambda args: service.add_attachment(
            email_entry_id=args["email_entry_id"],
            file_path=args["file_path"],
            display_name=args.get("display_name"),
        ),
        "outlook_list_attachments": lambda args: service.list_attachments(args["email_entry_id"]),
        "outlook_save_attachment": lambda args: service.save_attachment(
            email_entry_id=args["email_entry_id"],
            attachment_index=args["attachment_index"],
            save_path=args["save_path"],
        ),
        "outlook_remove_attachment": lambda args: service.remove_attachment(
            email_entry_id=args["email_entry_id"],
            attachment_index=args["attachment_index"],
        ),
        "outlook_create_new_message": lambda args: service.create_new_message(),
        # Folder Operations (7)
        "outlook_create_folder": lambda args: service.create_folder(
            folder_name=args["folder_name"],
            parent_folder=args.get("parent_folder", "Inbox"),
        ),
        "outlook_delete_folder": lambda args: service.delete_folder(args["folder_path"]),
        "outlook_rename_folder": lambda args: service.rename_folder(
            folder_path=args["folder_path"],
            new_name=args["new_name"],
        ),
        "outlook_move_folder": lambda args: service.move_folder(
            folder_path=args["folder_path"],
            destination_path=args["destination_path"],
        ),
        "outlook_list_folders": lambda args: service.list_folders(
            parent_folder=args.get("parent_folder", "Inbox"),
            recursive=args.get("recursive", False),
        ),
        "outlook_get_folder_item_count": lambda args: service.get_folder_item_count(
            args["folder_path"]
        ),
        "outlook_get_unread_count": lambda args: service.get_unread_count(
            args.get("folder_path", "Inbox")
        ),
        # Calendar Operations (12)
        "outlook_create_appointment": lambda args: service.create_appointment(
            subject=args["subject"],
            start_time=args["start_time"],
            end_time=args["end_time"],
            location=args.get("location"),
            body=args.get("body"),
            reminder_minutes=args.get("reminder_minutes", 15),
            busy_status=args.get("busy_status", 2),
        ),
        "outlook_create_recurring_event": lambda args: service.create_recurring_event(
            subject=args["subject"],
            start_time=args["start_time"],
            end_time=args["end_time"],
            recurrence_type=args["recurrence_type"],
            interval=args.get("interval", 1),
            occurrences=args.get("occurrences"),
            end_date=args.get("end_date"),
            location=args.get("location"),
            body=args.get("body"),
        ),
        "outlook_read_appointment": lambda args: service.read_appointment(
            args["appointment_entry_id"]
        ),
        "outlook_modify_appointment": lambda args: service.modify_appointment(
            appointment_entry_id=args["appointment_entry_id"],
            subject=args.get("subject"),
            start_time=args.get("start_time"),
            end_time=args.get("end_time"),
            location=args.get("location"),
            body=args.get("body"),
        ),
        "outlook_delete_appointment": lambda args: service.delete_appointment(
            args["appointment_entry_id"]
        ),
        "outlook_search_appointments": lambda args: service.search_appointments(
            subject=args.get("subject"),
            location=args.get("location"),
            start_date=args.get("start_date"),
            end_date=args.get("end_date"),
            max_results=args.get("max_results", 50),
        ),
        "outlook_get_appointments_by_date": lambda args: service.get_appointments_by_date(
            start_date=args["start_date"],
            end_date=args["end_date"],
        ),
        "outlook_set_reminder": lambda args: service.set_reminder(
            appointment_entry_id=args["appointment_entry_id"],
            reminder_minutes=args["reminder_minutes"],
        ),
        "outlook_set_busy_status": lambda args: service.set_busy_status(
            appointment_entry_id=args["appointment_entry_id"],
            busy_status=args["busy_status"],
        ),
        "outlook_export_appointment_ics": lambda args: service.export_appointment_ics(
            appointment_entry_id=args["appointment_entry_id"],
            output_path=args["output_path"],
        ),
        "outlook_get_calendar_count": lambda args: service.get_calendar_count(),
        "outlook_export_to_pdf": lambda args: service.export_to_pdf(args["output_path"]),
        # Meeting Operations (8)
        "outlook_create_meeting_request": lambda args: service.create_meeting_request(
            subject=args["subject"],
            start_time=args["start_time"],
            end_time=args["end_time"],
            required_attendees=args["required_attendees"],
            optional_attendees=args.get("optional_attendees"),
            location=args.get("location"),
            body=args.get("body"),
        ),
        "outlook_invite_participants": lambda args: service.invite_participants(
            meeting_entry_id=args["meeting_entry_id"],
            attendees=args["attendees"],
            required=args.get("required", True),
        ),
        "outlook_accept_meeting": lambda args: service.accept_meeting(args["meeting_entry_id"]),
        "outlook_decline_meeting": lambda args: service.decline_meeting(args["meeting_entry_id"]),
        "outlook_propose_new_time": lambda args: service.propose_new_time(
            meeting_entry_id=args["meeting_entry_id"],
            new_start=args["new_start"],
            new_end=args["new_end"],
        ),
        "outlook_cancel_meeting": lambda args: service.cancel_meeting(args["meeting_entry_id"]),
        "outlook_update_meeting": lambda args: service.update_meeting(
            meeting_entry_id=args["meeting_entry_id"],
            subject=args.get("subject"),
            start_time=args.get("start_time"),
            end_time=args.get("end_time"),
            location=args.get("location"),
        ),
        "outlook_check_availability": lambda args: service.check_availability(
            attendees=args["attendees"],
            start_time=args["start_time"],
            end_time=args["end_time"],
            duration_minutes=args.get("duration_minutes"),
        ),
        # Contact Operations (9)
        "outlook_create_contact": lambda args: service.create_contact(
            first_name=args["first_name"],
            last_name=args["last_name"],
            email=args.get("email"),
            phone=args.get("phone"),
            company=args.get("company"),
            job_title=args.get("job_title"),
        ),
        "outlook_modify_contact": lambda args: service.modify_contact(
            contact_entry_id=args["contact_entry_id"],
            first_name=args.get("first_name"),
            last_name=args.get("last_name"),
            email=args.get("email"),
            phone=args.get("phone"),
        ),
        "outlook_delete_contact": lambda args: service.delete_contact(args["contact_entry_id"]),
        "outlook_search_contact": lambda args: service.search_contact(args["search_term"]),
        "outlook_list_all_contacts": lambda args: service.list_all_contacts(),
        "outlook_create_contact_group": lambda args: service.create_contact_group(
            args["group_name"]
        ),
        "outlook_add_to_contact_group": lambda args: service.add_to_contact_group(
            group_entry_id=args["group_entry_id"],
            contact_email=args["contact_email"],
        ),
        "outlook_export_contacts_vcf": lambda args: service.export_contacts_vcf(
            args["output_path"]
        ),
        "outlook_import_contacts": lambda args: service.import_contacts(args["file_path"]),
        # Task Operations (7)
        "outlook_create_task": lambda args: service.create_task(
            subject=args["subject"],
            body=args.get("body"),
            due_date=args.get("due_date"),
            priority=args.get("priority", 1),
        ),
        "outlook_modify_task": lambda args: service.modify_task(
            task_entry_id=args["task_entry_id"],
            subject=args.get("subject"),
            body=args.get("body"),
            due_date=args.get("due_date"),
        ),
        "outlook_delete_task": lambda args: service.delete_task(args["task_entry_id"]),
        "outlook_mark_task_complete": lambda args: service.mark_task_complete(
            args["task_entry_id"]
        ),
        "outlook_set_task_priority": lambda args: service.set_task_priority(
            task_entry_id=args["task_entry_id"],
            priority=args["priority"],
        ),
        "outlook_set_task_due_date": lambda args: service.set_task_due_date(
            task_entry_id=args["task_entry_id"],
            due_date=args["due_date"],
        ),
        "outlook_list_tasks": lambda args: service.list_tasks(
            completed=args.get("completed", False)
        ),
        # Advanced Operations (10)
        "outlook_list_accounts": lambda args: service.list_accounts(),
        "outlook_get_default_account": lambda args: service.get_default_account(),
        "outlook_get_inbox_count": lambda args: service.get_inbox_count(),
        "outlook_create_category": lambda args: service.create_category(
            name=args["name"],
            color=args.get("color", 0),
        ),
        "outlook_list_categories": lambda args: service.list_categories(),
        "outlook_apply_category": lambda args: service.apply_category(
            item_entry_id=args["item_entry_id"],
            category=args["category"],
        ),
        "outlook_com_operation": lambda args: service.com_operation(args["operation_name"]),
    }


# IMPORTANT: Ce fichier sera trop long. Suite dans server_tools.py
# pour les définitions de tools et server_lifecycle.py pour le lifecycle
