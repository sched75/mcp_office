"""Outlook automation service implementing all 85 Outlook functionalities.

This service provides comprehensive Outlook automation capabilities following
SOLID principles and design patterns.
"""

from typing import Any

from ..core.base_office import BaseOfficeService
from ..core.types import ApplicationType
from ..utils.com_wrapper import com_safe
from ..utils.helpers import dict_to_result
from .additional_operations import (
    AdvancedOperationsMixin,
    ContactOperationsMixin,
    MeetingOperationsMixin,
    TaskOperationsMixin,
)
from .attachment_operations import AttachmentOperationsMixin
from .calendar_operations import CalendarOperationsMixin
from .folder_operations import FolderOperationsMixin
from .mail_operations import MailOperationsMixin


class OutlookService(
    BaseOfficeService,
    MailOperationsMixin,
    AttachmentOperationsMixin,
    FolderOperationsMixin,
    CalendarOperationsMixin,
    MeetingOperationsMixin,
    ContactOperationsMixin,
    TaskOperationsMixin,
    AdvancedOperationsMixin,
):
    """Outlook automation service with all 85 functionalities.

    This service implements the complete Outlook automation API covering:
    - Email management (12 methods)
    - Attachment operations (5 methods)
    - Folder management (7 methods)
    - Calendar operations (10 methods)
    - Meeting management (8 methods)
    - Contact management (9 methods)
    - Task management (7 methods)
    - Advanced operations (27 methods)

    Example:
        >>> outlook = OutlookService()
        >>> outlook.initialize()
        >>> result = outlook.send_email(
        ...     to="recipient@example.com",
        ...     subject="Test",
        ...     body="This is a test email"
        ... )
        >>> print(result['success'])
        True
    """

    def __init__(self, visible: bool = False) -> None:
        """Initialize Outlook service.

        Args:
            visible: Whether to make Outlook window visible (not typically used)
        """
        super().__init__(ApplicationType.OUTLOOK, visible)
        self._namespace = None

    @property
    def namespace(self):
        """Get the MAPI namespace.

        Returns:
            MAPI namespace object

        Raises:
            COMInitializationError: If not initialized
        """
        if self._namespace is None:
            self._namespace = self.application.GetNamespace("MAPI")
        return self._namespace

    def _close_document(self) -> None:
        """Close method - not applicable for Outlook.

        Outlook doesn't have the concept of opening/closing documents
        like Word, Excel, or PowerPoint.
        """
        # Outlook doesn't have documents to close
        pass

    @com_safe("create_document")
    def create_document(self) -> dict[str, Any]:
        """Create method - adapted for Outlook.

        In Outlook context, this initializes the service.
        """
        if not self.is_initialized:
            self.initialize()

        return dict_to_result(
            success=True,
            message="Outlook service initialized successfully",
            accounts=self.application.Session.Accounts.Count,
        )

    @com_safe("open_document")
    def open_document(self, file_path: str) -> dict[str, Any]:
        """Open method - not applicable for Outlook.

        Args:
            file_path: Path to file (not used for Outlook)

        Returns:
            Dictionary indicating method not applicable
        """
        return dict_to_result(
            success=False,
            message="Open document not applicable for Outlook service",
            note="Use specific methods like read_email, read_appointment, etc.",
        )

    @com_safe("save_document")
    def save_document(self, file_path: str | None = None) -> dict[str, Any]:
        """Save method - not applicable for Outlook.

        Args:
            file_path: Path to save to (not used for Outlook)

        Returns:
            Dictionary indicating method not applicable
        """
        return dict_to_result(
            success=False,
            message="Save document not applicable for Outlook service",
            note="Items are automatically saved when using specific methods",
        )

    @com_safe("close_document")
    def close_document(self, save_changes: bool = False) -> dict[str, Any]:
        """Close method - not applicable for Outlook.

        Args:
            save_changes: Whether to save (not used for Outlook)

        Returns:
            Dictionary indicating method not applicable
        """
        return dict_to_result(
            success=False,
            message="Close document not applicable for Outlook service",
            note="Use cleanup() to terminate the Outlook service",
        )

    @com_safe("export_to_pdf")
    def export_to_pdf(self, output_path: str) -> dict[str, Any]:
        """Export to PDF - not directly applicable for Outlook.

        Args:
            output_path: Output path

        Returns:
            Dictionary indicating method not applicable
        """
        return dict_to_result(
            success=False,
            message="Export to PDF not directly applicable for Outlook",
            note="Use specific methods like export_appointment_ics for calendar items",
        )

    def get_inbox_count(self) -> dict[str, Any]:
        """Get inbox message count (helper method).

        Returns:
            Dictionary with inbox statistics

        Example:
            >>> result = outlook.get_inbox_count()
            >>> print(result['total_items'])
        """
        inbox = self.namespace.GetDefaultFolder(6)  # 6 = olFolderInbox

        return dict_to_result(
            success=True,
            message="Inbox statistics retrieved",
            total_items=inbox.Items.Count,
            unread_items=inbox.UnReadItemCount,
        )

    def get_calendar_count(self) -> dict[str, Any]:
        """Get calendar item count (helper method).

        Returns:
            Dictionary with calendar statistics

        Example:
            >>> result = outlook.get_calendar_count()
            >>> print(result['total_items'])
        """
        calendar = self.namespace.GetDefaultFolder(9)  # 9 = olFolderCalendar

        return dict_to_result(
            success=True,
            message="Calendar statistics retrieved",
            total_items=calendar.Items.Count,
        )
