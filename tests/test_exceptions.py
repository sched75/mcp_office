"""Unit tests for exceptions module."""

import pytest

from src.core.exceptions import (
    COMInitializationError,
    COMOperationError,
    DocumentNotFoundError,
    DocumentNotOpenError,
    FileOperationError,
    InvalidParameterError,
    OfficeAutomationError,
    ProtectionError,
    RangeError,
    ResourceCleanupError,
    TemplateError,
)


class TestOfficeAutomationError:
    """Tests for base OfficeAutomationError class."""

    def test_simple_message(self) -> None:
        """Test exception with simple message."""
        error = OfficeAutomationError("Simple error message")
        assert str(error) == "Simple error message"

    def test_message_with_details(self) -> None:
        """Test exception with message and details."""
        error = OfficeAutomationError("Error message", "Additional details")
        assert "Error message" in str(error)
        assert "Additional details" in str(error)

    def test_can_be_raised(self) -> None:
        """Test exception can be raised."""
        with pytest.raises(OfficeAutomationError):
            raise OfficeAutomationError("Test error")


class TestCOMInitializationError:
    """Tests for COMInitializationError."""

    def test_with_application_name(self) -> None:
        """Test error with application name."""
        error = COMInitializationError(app_type="Word")
        assert "Word" in str(error)

    def test_with_details(self) -> None:
        """Test error with details."""
        error = COMInitializationError(app_type="Excel", details="Not installed")
        assert "Excel" in str(error)
        assert "Not installed" in str(error)

    def test_inheritance(self) -> None:
        """Test error inherits from OfficeAutomationError."""
        error = COMInitializationError(app_type="PowerPoint")
        assert isinstance(error, OfficeAutomationError)


class TestDocumentNotFoundError:
    """Tests for DocumentNotFoundError."""

    def test_with_file_path(self) -> None:
        """Test error with file path."""
        error = DocumentNotFoundError(file_path="/path/to/document.docx")
        assert "/path/to/document.docx" in str(error)

    def test_inheritance(self) -> None:
        """Test error inherits from OfficeAutomationError."""
        error = DocumentNotFoundError(file_path="/test.docx")
        assert isinstance(error, OfficeAutomationError)


class TestDocumentNotOpenError:
    """Tests for DocumentNotOpenError."""

    def test_with_operation(self) -> None:
        """Test error with operation name."""
        error = DocumentNotOpenError(operation="save_document")
        assert "save_document" in str(error)

    def test_inheritance(self) -> None:
        """Test error inherits from OfficeAutomationError."""
        error = DocumentNotOpenError(operation="test")
        assert isinstance(error, OfficeAutomationError)


class TestInvalidParameterError:
    """Tests for InvalidParameterError."""

    def test_with_parameter_and_reason(self) -> None:
        """Test error with parameter name and reason."""
        error = InvalidParameterError(param_name="rows", param_value=0, reason="must be positive")
        assert "rows" in str(error)
        assert "must be positive" in str(error)

    def test_with_value(self) -> None:
        """Test error with value."""
        error = InvalidParameterError(param_name="color", param_value="xyz", reason="invalid")
        assert "color" in str(error)
        assert "xyz" in str(error)

    def test_inheritance(self) -> None:
        """Test error inherits from OfficeAutomationError."""
        error = InvalidParameterError(param_name="test", param_value="bad", reason="test")
        assert isinstance(error, OfficeAutomationError)


class TestCOMOperationError:
    """Tests for COMOperationError."""

    def test_with_operation_name(self) -> None:
        """Test error with operation name."""
        original_error = ValueError("Original error")
        error = COMOperationError(operation="create_document", com_error=original_error)
        assert "create_document" in str(error)

    def test_with_original_error(self) -> None:
        """Test error includes original error message."""
        original_error = ValueError("Original error message")
        error = COMOperationError(operation="test", com_error=original_error)
        assert "Original error message" in str(error)

    def test_inheritance(self) -> None:
        """Test error inherits from OfficeAutomationError."""
        error = COMOperationError(operation="test", com_error=Exception())
        assert isinstance(error, OfficeAutomationError)


class TestFileOperationError:
    """Tests for FileOperationError."""

    def test_with_file_path_and_operation(self) -> None:
        """Test error with file path and operation."""
        error = FileOperationError(
            file_path="/test.docx", operation="save", reason="permission denied"
        )
        assert "/test.docx" in str(error)
        assert "save" in str(error)
        assert "permission denied" in str(error)

    def test_inheritance(self) -> None:
        """Test error inherits from OfficeAutomationError."""
        error = FileOperationError(file_path="/test", operation="read", reason="test")
        assert isinstance(error, OfficeAutomationError)


class TestRangeError:
    """Tests for RangeError."""

    def test_with_range_address(self) -> None:
        """Test error with range address."""
        error = RangeError(range_address="A1:B10", reason="invalid range")
        assert "A1:B10" in str(error)
        assert "invalid range" in str(error)

    def test_inheritance(self) -> None:
        """Test error inherits from OfficeAutomationError."""
        error = RangeError(range_address="A1:B10", reason="test")
        assert isinstance(error, OfficeAutomationError)


class TestTemplateError:
    """Tests for TemplateError."""

    def test_with_template_name(self) -> None:
        """Test error with template name."""
        error = TemplateError(template_path="MyTemplate", reason="not found")
        assert "MyTemplate" in str(error)
        assert "not found" in str(error)

    def test_inheritance(self) -> None:
        """Test error inherits from OfficeAutomationError."""
        error = TemplateError(template_path="test", reason="test")
        assert isinstance(error, OfficeAutomationError)


class TestProtectionError:
    """Tests for ProtectionError."""

    def test_with_operation(self) -> None:
        """Test error with operation."""
        error = ProtectionError(operation="modify_cell", reason="worksheet is protected")
        assert "modify_cell" in str(error)
        assert "protected" in str(error)

    def test_inheritance(self) -> None:
        """Test error inherits from OfficeAutomationError."""
        error = ProtectionError(operation="test", reason="test")
        assert isinstance(error, OfficeAutomationError)


class TestResourceCleanupError:
    """Tests for ResourceCleanupError."""

    def test_with_resource_name(self) -> None:
        """Test error with resource name."""
        error = ResourceCleanupError(resource="COM Application", details="failed to release")
        assert "COM Application" in str(error)
        assert "failed to release" in str(error)

    def test_inheritance(self) -> None:
        """Test error inherits from OfficeAutomationError."""
        error = ResourceCleanupError(resource="test", details="test")
        assert isinstance(error, OfficeAutomationError)


class TestExceptionChaining:
    """Tests for exception chaining."""

    def test_can_chain_exceptions(self) -> None:
        """Test exceptions can be chained."""
        try:
            try:
                raise ValueError("Original error")
            except ValueError as e:
                raise COMOperationError(operation="test", com_error=e) from e
        except COMOperationError as error:
            assert error.__cause__ is not None
            assert isinstance(error.__cause__, ValueError)

    def test_exception_context_preserved(self) -> None:
        """Test exception context is preserved."""
        try:
            raise InvalidParameterError(param_name="test", param_value="bad", reason="invalid")
        except InvalidParameterError as error:
            assert isinstance(error, OfficeAutomationError)
            assert "test" in str(error)
