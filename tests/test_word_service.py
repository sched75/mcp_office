"""Unit tests for WordService with mocked COM objects."""

from pathlib import Path
from typing import Any
from unittest.mock import MagicMock, Mock, patch

import pytest

from src.core.exceptions import DocumentNotOpenError
from src.word.word_service import WordService


@pytest.fixture
def word_service(mock_pythoncom: Any, mocker: Any) -> WordService:
    """Create a WordService instance with mocked COM."""
    with patch("win32com.client.Dispatch") as mock_dispatch:
        mock_app = MagicMock()
        mock_dispatch.return_value = mock_app
        service = WordService()
        service.initialize()
        return service


@pytest.fixture
def word_service_with_doc(word_service: WordService, mock_word_document: MagicMock) -> WordService:
    """Create a WordService with an open document."""
    word_service._current_document = mock_word_document
    return word_service


class TestWordServiceInitialization:
    """Tests for WordService initialization."""

    def test_service_creates_successfully(self, mock_pythoncom: Any) -> None:
        """Test WordService can be instantiated."""
        with patch("win32com.client.Dispatch"):
            service = WordService()
            assert service is not None
            assert not service.is_initialized

    def test_service_initializes(self, word_service: WordService) -> None:
        """Test WordService initializes COM application."""
        assert word_service.is_initialized
        assert word_service.application is not None

    def test_service_cleanup(self, word_service: WordService) -> None:
        """Test WordService cleanup."""
        word_service.cleanup()
        # Should not raise exception


class TestDocumentCreation:
    """Tests for document creation methods."""

    def test_create_document(self, word_service: WordService) -> None:
        """Test creating a new Word document."""
        mock_doc = MagicMock()
        word_service.application.Documents.Add.return_value = mock_doc

        result = word_service.create_document()

        assert result["success"] is True
        assert "created" in result["message"].lower()
        word_service.application.Documents.Add.assert_called_once()

    def test_open_document(self, word_service: WordService, sample_docx_path: Path) -> None:
        """Test opening an existing Word document."""
        sample_docx_path.touch()
        mock_doc = MagicMock()
        mock_doc.FullName = str(sample_docx_path)
        word_service.application.Documents.Open.return_value = mock_doc

        result = word_service.open_document(str(sample_docx_path))

        assert result["success"] is True
        word_service.application.Documents.Open.assert_called_once()

    def test_create_from_template(self, word_service: WordService, tmp_path: Path) -> None:
        """Test creating document from template."""
        template_path = tmp_path / "template.dotx"
        template_path.touch()

        mock_doc = MagicMock()
        word_service.application.Documents.Add.return_value = mock_doc

        result = word_service.create_from_template(str(template_path))

        assert result["success"] is True
        word_service.application.Documents.Add.assert_called_once()


class TestDocumentOperations:
    """Tests for document operations."""

    def test_save_document(self, word_service_with_doc: WordService) -> None:
        """Test saving a document."""
        result = word_service_with_doc.save_document()

        assert result["success"] is True
        word_service_with_doc.current_document.Save.assert_called_once()

    def test_save_document_no_doc_open(self, word_service: WordService) -> None:
        """Test saving when no document is open."""
        with pytest.raises(DocumentNotOpenError):
            word_service.save_document()

    def test_close_document(self, word_service_with_doc: WordService) -> None:
        """Test closing a document."""
        result = word_service_with_doc.close_document()

        assert result["success"] is True
        word_service_with_doc.current_document.Close.assert_called_once()

    def test_export_to_pdf(self, word_service_with_doc: WordService, tmp_path: Path) -> None:
        """Test exporting document to PDF."""
        pdf_path = tmp_path / "output.pdf"

        result = word_service_with_doc.export_to_pdf(str(pdf_path))

        assert result["success"] is True
        word_service_with_doc.current_document.ExportAsFixedFormat.assert_called_once()


class TestTextOperations:
    """Tests for text manipulation methods."""

    def test_add_paragraph(self, word_service_with_doc: WordService) -> None:
        """Test adding a paragraph."""
        mock_range = MagicMock()
        word_service_with_doc.current_document.Content = mock_range

        result = word_service_with_doc.add_paragraph("Test paragraph")

        assert result["success"] is True
        assert "Test paragraph" in str(mock_range.InsertAfter.call_args)

    def test_insert_text_at_position(self, word_service_with_doc: WordService) -> None:
        """Test inserting text at specific position."""
        mock_range = MagicMock()
        word_service_with_doc.current_document.Range.return_value = mock_range

        result = word_service_with_doc.insert_text_at_position(10, "Inserted text")

        assert result["success"] is True
        word_service_with_doc.current_document.Range.assert_called_with(10, 10)

    def test_find_and_replace(self, word_service_with_doc: WordService) -> None:
        """Test find and replace operation."""
        mock_find = MagicMock()
        mock_range = MagicMock()
        mock_range.Find = mock_find
        mock_find.Execute.return_value = True
        word_service_with_doc.current_document.Content = mock_range

        result = word_service_with_doc.find_and_replace("old", "new")

        assert result["success"] is True
        mock_find.Execute.assert_called()


class TestFormatting:
    """Tests for formatting methods."""

    def test_set_paragraph_alignment(self, word_service_with_doc: WordService) -> None:
        """Test setting paragraph alignment."""
        mock_para = MagicMock()
        word_service_with_doc.current_document.Paragraphs.return_value = mock_para

        result = word_service_with_doc.set_paragraph_alignment(1, "center")

        assert result["success"] is True

    def test_apply_style(self, word_service_with_doc: WordService) -> None:
        """Test applying a style."""
        mock_range = MagicMock()
        word_service_with_doc.current_document.Range.return_value = mock_range

        result = word_service_with_doc.apply_style(0, 10, "Heading 1")

        assert result["success"] is True

    def test_set_line_spacing(self, word_service_with_doc: WordService) -> None:
        """Test setting line spacing."""
        mock_para = MagicMock()
        word_service_with_doc.current_document.Paragraphs.return_value = mock_para

        result = word_service_with_doc.set_line_spacing(1, 1.5)

        assert result["success"] is True


class TestTableOperations:
    """Tests for table operations."""

    def test_insert_table(self, word_service_with_doc: WordService) -> None:
        """Test inserting a table."""
        mock_content = MagicMock()
        mock_tables = MagicMock()
        word_service_with_doc.current_document.Content = mock_content
        word_service_with_doc.current_document.Tables = mock_tables

        result = word_service_with_doc.insert_table(3, 4)

        assert result["success"] is True
        mock_tables.Add.assert_called_once()

    def test_set_table_cell_text(self, word_service_with_doc: WordService) -> None:
        """Test setting table cell text."""
        mock_table = MagicMock()
        mock_cell = MagicMock()
        mock_table.Cell.return_value = mock_cell
        word_service_with_doc.current_document.Tables.return_value = mock_table

        result = word_service_with_doc.set_table_cell_text(1, 1, 1, "Cell text")

        assert result["success"] is True
        mock_table.Cell.assert_called_with(1, 1)


class TestImageOperations:
    """Tests for image operations."""

    def test_insert_image(self, word_service_with_doc: WordService, sample_image_path: Path) -> None:
        """Test inserting an image."""
        mock_shapes = MagicMock()
        mock_range = MagicMock()
        word_service_with_doc.current_document.Content = mock_range
        word_service_with_doc.current_document.InlineShapes = mock_shapes

        result = word_service_with_doc.insert_image(str(sample_image_path))

        assert result["success"] is True
        mock_shapes.AddPicture.assert_called_once()

    def test_resize_image(self, word_service_with_doc: WordService) -> None:
        """Test resizing an image."""
        mock_shape = MagicMock()
        word_service_with_doc.current_document.InlineShapes.return_value = mock_shape

        result = word_service_with_doc.resize_image(1, 200, 150)

        assert result["success"] is True
        assert mock_shape.Width == 200
        assert mock_shape.Height == 150


class TestDocumentStructure:
    """Tests for document structure methods."""

    def test_insert_page_break(self, word_service_with_doc: WordService) -> None:
        """Test inserting a page break."""
        mock_range = MagicMock()
        word_service_with_doc.current_document.Content = mock_range

        result = word_service_with_doc.insert_page_break()

        assert result["success"] is True
        mock_range.InsertBreak.assert_called_once()

    def test_add_header(self, word_service_with_doc: WordService) -> None:
        """Test adding a header."""
        mock_section = MagicMock()
        mock_header = MagicMock()
        mock_section.Headers.return_value = mock_header
        word_service_with_doc.current_document.Sections.return_value = mock_section

        result = word_service_with_doc.add_header("Header text")

        assert result["success"] is True

    def test_add_footer(self, word_service_with_doc: WordService) -> None:
        """Test adding a footer."""
        mock_section = MagicMock()
        mock_footer = MagicMock()
        mock_section.Footers.return_value = mock_footer
        word_service_with_doc.current_document.Sections.return_value = mock_section

        result = word_service_with_doc.add_footer("Footer text")

        assert result["success"] is True


class TestDocumentProtection:
    """Tests for document protection methods."""

    def test_protect_document(self, word_service_with_doc: WordService) -> None:
        """Test protecting a document."""
        result = word_service_with_doc.protect_document("password")

        assert result["success"] is True
        word_service_with_doc.current_document.Protect.assert_called_once()

    def test_set_password(self, word_service_with_doc: WordService) -> None:
        """Test setting document password."""
        result = word_service_with_doc.set_password("mypassword")

        assert result["success"] is True

    def test_unprotect_document(self, word_service_with_doc: WordService) -> None:
        """Test unprotecting a document."""
        result = word_service_with_doc.unprotect_document("password")

        assert result["success"] is True
        word_service_with_doc.current_document.Unprotect.assert_called_once()


class TestDocumentProperties:
    """Tests for document properties methods."""

    def test_get_document_properties(self, word_service_with_doc: WordService) -> None:
        """Test getting document properties."""
        mock_props = MagicMock()
        mock_props.Item.return_value.Value = "Test Value"
        word_service_with_doc.current_document.BuiltInDocumentProperties = mock_props

        result = word_service_with_doc.get_document_properties()

        assert result["success"] is True
        assert "properties" in result

    def test_set_document_properties(self, word_service_with_doc: WordService) -> None:
        """Test setting document properties."""
        mock_props = MagicMock()
        word_service_with_doc.current_document.BuiltInDocumentProperties = mock_props

        result = word_service_with_doc.set_document_properties(
            title="Test Title",
            author="Test Author"
        )

        assert result["success"] is True
