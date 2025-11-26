"""Custom exceptions for Office automation.

This module defines a hierarchy of exceptions following Python best practices
and providing clear, actionable error messages.
"""


class OfficeAutomationError(Exception):
    """Base exception for all Office automation errors.

    All custom exceptions in this module inherit from this base class,
    following the Liskov Substitution Principle (SOLID).
    """

    def __init__(self, message: str, details: str | None = None) -> None:
        """Initialize the exception with message and optional details.

        Args:
            message: Primary error message
            details: Additional error details or context
        """
        self.message = message
        self.details = details
        super().__init__(self._format_message())

    def _format_message(self) -> str:
        """Format the complete error message."""
        if self.details:
            return f"{self.message}\nDetails: {self.details}"
        return self.message


class COMInitializationError(OfficeAutomationError):
    """Raised when COM initialization fails."""

    def __init__(self, app_type: str, details: str | None = None) -> None:
        """Initialize COM initialization error.

        Args:
            app_type: Type of Office application (Word, Excel, PowerPoint)
            details: Additional error details
        """
        message = f"Failed to initialize {app_type} COM object"
        super().__init__(message, details)


class DocumentNotFoundError(OfficeAutomationError):
    """Raised when a document/file is not found."""

    def __init__(self, file_path: str) -> None:
        """Initialize document not found error.

        Args:
            file_path: Path to the missing document
        """
        message = f"Document not found: {file_path}"
        super().__init__(message)


class DocumentNotOpenError(OfficeAutomationError):
    """Raised when attempting operations on a closed document."""

    def __init__(self, operation: str) -> None:
        """Initialize document not open error.

        Args:
            operation: The operation that was attempted
        """
        message = f"Cannot perform '{operation}': No document is currently open"
        super().__init__(message)


class InvalidParameterError(OfficeAutomationError):
    """Raised when invalid parameters are provided."""

    def __init__(self, param_name: str, param_value: object, reason: str) -> None:
        """Initialize invalid parameter error.

        Args:
            param_name: Name of the invalid parameter
            param_value: Value that was provided
            reason: Explanation of why the parameter is invalid
        """
        message = f"Invalid parameter '{param_name}': {param_value}"
        super().__init__(message, reason)


class COMOperationError(OfficeAutomationError):
    """Raised when a COM operation fails."""

    def __init__(self, operation: str, com_error: Exception) -> None:
        """Initialize COM operation error.

        Args:
            operation: Description of the operation that failed
            com_error: The underlying COM exception
        """
        message = f"COM operation failed: {operation}"
        details = f"{type(com_error).__name__}: {com_error!s}"
        super().__init__(message, details)


class FileOperationError(OfficeAutomationError):
    """Raised when file operations fail."""

    def __init__(self, operation: str, file_path: str, reason: str) -> None:
        """Initialize file operation error.

        Args:
            operation: Type of file operation (save, open, close, etc.)
            file_path: Path to the file
            reason: Reason for failure
        """
        message = f"File operation '{operation}' failed for: {file_path}"
        super().__init__(message, reason)


class RangeError(OfficeAutomationError):
    """Raised when cell/range operations fail (Excel specific)."""

    def __init__(self, range_address: str, reason: str) -> None:
        """Initialize range error.

        Args:
            range_address: Address of the range (e.g., "A1:B10")
            reason: Reason for failure
        """
        message = f"Range operation failed for: {range_address}"
        super().__init__(message, reason)


class TemplateError(OfficeAutomationError):
    """Raised when template operations fail."""

    def __init__(self, template_path: str, reason: str) -> None:
        """Initialize template error.

        Args:
            template_path: Path to the template
            reason: Reason for failure
        """
        message = f"Template operation failed: {template_path}"
        super().__init__(message, reason)


class ProtectionError(OfficeAutomationError):
    """Raised when document protection operations fail."""

    def __init__(self, operation: str, reason: str) -> None:
        """Initialize protection error.

        Args:
            operation: Protection operation that failed
            reason: Reason for failure
        """
        message = f"Protection operation '{operation}' failed"
        super().__init__(message, reason)


class ResourceCleanupError(OfficeAutomationError):
    """Raised when resource cleanup fails."""

    def __init__(self, resource: str, details: str | None = None) -> None:
        """Initialize resource cleanup error.

        Args:
            resource: Type of resource that failed to clean up
            details: Additional details about the failure
        """
        message = f"Failed to cleanup resource: {resource}"
        super().__init__(message, details)


class OutlookItemNotFoundError(OfficeAutomationError):
    """Raised when an Outlook item is not found."""

    def __init__(self, item_type: str, identifier: str) -> None:
        """Initialize Outlook item not found error.

        Args:
            item_type: Type of Outlook item (email, appointment, etc.)
            identifier: Identifier used to search for the item
        """
        message = f"Outlook {item_type} not found: {identifier}"
        super().__init__(message)


class InvalidRecipientError(OfficeAutomationError):
    """Raised when an email recipient is invalid."""

    def __init__(self, recipient: str, reason: str) -> None:
        """Initialize invalid recipient error.

        Args:
            recipient: Email address that is invalid
            reason: Reason why the recipient is invalid
        """
        message = f"Invalid recipient: {recipient}"
        super().__init__(message, reason)


class FolderOperationError(OfficeAutomationError):
    """Raised when folder operations fail."""

    def __init__(self, operation: str, folder_name: str, reason: str) -> None:
        """Initialize folder operation error.

        Args:
            operation: Type of folder operation that failed
            folder_name: Name of the folder
            reason: Reason for failure
        """
        message = f"Folder operation '{operation}' failed for: {folder_name}"
        super().__init__(message, reason)


class AttachmentError(OfficeAutomationError):
    """Raised when attachment operations fail."""

    def __init__(self, operation: str, filename: str, reason: str) -> None:
        """Initialize attachment error.

        Args:
            operation: Type of attachment operation that failed
            filename: Name of the attachment file
            reason: Reason for failure
        """
        message = f"Attachment operation '{operation}' failed for: {filename}"
        super().__init__(message, reason)


class CalendarOperationError(OfficeAutomationError):
    """Raised when calendar operations fail."""

    def __init__(self, operation: str, reason: str) -> None:
        """Initialize calendar operation error.

        Args:
            operation: Type of calendar operation that failed
            reason: Reason for failure
        """
        message = f"Calendar operation '{operation}' failed"
        super().__init__(message, reason)


class MeetingOperationError(OfficeAutomationError):
    """Raised when meeting operations fail."""

    def __init__(self, operation: str, reason: str) -> None:
        """Initialize meeting operation error.

        Args:
            operation: Type of meeting operation that failed
            reason: Reason for failure
        """
        message = f"Meeting operation '{operation}' failed"
        super().__init__(message, reason)
