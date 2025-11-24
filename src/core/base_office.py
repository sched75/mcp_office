"""Abstract base class for Office automation services.

This module implements the base class for all Office services following SOLID principles:
- Single Responsibility: Each service handles one Office application
- Open/Closed: Extensible through inheritance, closed for modification
- Liskov Substitution: All services can be used interchangeably through this interface
- Interface Segregation: Clean, focused interface
- Dependency Inversion: Depends on abstractions (COM interface)
"""

import atexit
from abc import ABC, abstractmethod
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Generic, TypeVar

import pythoncom
import win32com.client

from .exceptions import (
    COMInitializationError,
    COMOperationError,
    DocumentNotOpenError,
    ResourceCleanupError,
)
from .types import ApplicationType

# Type variable for the COM application object
TApp = TypeVar("TApp")


class BaseOfficeService(ABC, Generic[TApp]):
    """Abstract base class for Office automation services.

    This class provides common functionality for all Office applications
    while enforcing a consistent interface through abstract methods.
    """

    def __init__(self, application_type: ApplicationType, visible: bool = False) -> None:
        """Initialize the Office service.

        Args:
            application_type: Type of Office application
            visible: Whether to make the application window visible

        Raises:
            COMInitializationError: If COM initialization fails
        """
        self._app_type = application_type
        self._visible = visible
        self._app: TApp | None = None
        self._current_document: Any | None = None
        self._is_initialized = False

        # Register cleanup on exit
        atexit.register(self.cleanup)

    @property
    def application(self) -> TApp:
        """Get the COM application object.

        Returns:
            The COM application object

        Raises:
            COMInitializationError: If application is not initialized
        """
        if not self._is_initialized or self._app is None:
            msg = "Application not initialized. Call initialize() first."
            raise COMInitializationError(self._app_type.value, msg)
        return self._app

    @property
    def current_document(self) -> Any:
        """Get the currently active document.

        Returns:
            The current document object

        Raises:
            DocumentNotOpenError: If no document is open
        """
        if self._current_document is None:
            raise DocumentNotOpenError("access current document")
        return self._current_document

    @property
    def is_initialized(self) -> bool:
        """Check if the service is initialized."""
        return self._is_initialized

    def initialize(self) -> None:
        """Initialize the COM application.

        This method sets up the COM environment and creates the application object.

        Raises:
            COMInitializationError: If initialization fails
        """
        if self._is_initialized:
            return

        try:
            # Initialize COM for this thread
            pythoncom.CoInitialize()

            # Create the COM application
            self._app = win32com.client.Dispatch(self._app_type.value)
            self._app.Visible = self._visible
            self._app.DisplayAlerts = False  # Prevent popup dialogs

            self._is_initialized = True

        except Exception as e:
            raise COMInitializationError(self._app_type.value, str(e)) from e

    @contextmanager
    def com_operation(self, operation_name: str):
        """Context manager for COM operations with error handling.

        Args:
            operation_name: Name of the operation for error messages

        Yields:
            None

        Raises:
            COMOperationError: If the COM operation fails
        """
        try:
            yield
        except Exception as e:
            raise COMOperationError(operation_name, e) from e

    def cleanup(self) -> None:
        """Clean up COM resources.

        This method should be called when done with the service.
        It's automatically called on program exit via atexit.
        """
        errors = []

        try:
            # Close current document if any
            if self._current_document is not None:
                try:
                    self._close_document()
                except Exception as e:
                    errors.append(f"Document cleanup: {e}")

            # Quit application
            if self._app is not None:
                try:
                    self._app.Quit()
                except Exception as e:
                    errors.append(f"Application quit: {e}")

            # Uninitialize COM
            try:
                pythoncom.CoUninitialize()
            except Exception as e:
                errors.append(f"COM uninitialize: {e}")

        except Exception as e:
            errors.append(f"General cleanup: {e}")
        finally:
            self._app = None
            self._current_document = None
            self._is_initialized = False

        if errors:
            raise ResourceCleanupError(
                self._app_type.value, "; ".join(errors)
            )

    def _validate_file_path(self, file_path: str | Path) -> Path:
        """Validate and normalize a file path.

        Args:
            file_path: Path to validate

        Returns:
            Normalized Path object
        """
        path = Path(file_path)
        return path.resolve()

    @abstractmethod
    def _close_document(self) -> None:
        """Close the current document.

        This is application-specific and must be implemented by subclasses.
        """

    @abstractmethod
    def create_document(self) -> dict[str, Any]:
        """Create a new document.

        Returns:
            Dictionary with creation result and metadata
        """

    @abstractmethod
    def open_document(self, file_path: str) -> dict[str, Any]:
        """Open an existing document.

        Args:
            file_path: Path to the document

        Returns:
            Dictionary with open result and metadata
        """

    @abstractmethod
    def save_document(self, file_path: str | None = None) -> dict[str, Any]:
        """Save the current document.

        Args:
            file_path: Optional path to save to (Save As)

        Returns:
            Dictionary with save result and metadata
        """

    @abstractmethod
    def close_document(self, save_changes: bool = False) -> dict[str, Any]:
        """Close the current document.

        Args:
            save_changes: Whether to save changes before closing

        Returns:
            Dictionary with close result
        """

    @abstractmethod
    def export_to_pdf(self, output_path: str) -> dict[str, Any]:
        """Export the current document to PDF.

        Args:
            output_path: Path for the PDF file

        Returns:
            Dictionary with export result
        """


class DocumentOperationMixin:
    """Mixin providing common document operations.

    This mixin follows the Interface Segregation Principle by providing
    optional functionality that can be added to services as needed.
    """

    def get_document_info(self, document: Any) -> dict[str, Any]:
        """Get information about a document.

        Args:
            document: The document object

        Returns:
            Dictionary with document metadata
        """
        try:
            return {
                "name": getattr(document, "Name", "Unknown"),
                "path": getattr(document, "FullName", ""),
                "saved": getattr(document, "Saved", True),
            }
        except Exception:
            return {"name": "Unknown", "path": "", "saved": True}
