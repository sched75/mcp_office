"""Unit tests for PowerPointService with mocked COM objects."""

from pathlib import Path
from typing import Any
from unittest.mock import MagicMock, patch

import pytest

from src.core.exceptions import DocumentNotOpenError
from src.powerpoint.powerpoint_service import PowerPointService


@pytest.fixture
def ppt_service(mock_pythoncom: Any, mocker: Any) -> PowerPointService:
    """Create a PowerPointService instance with mocked COM."""
    with patch("win32com.client.Dispatch") as mock_dispatch:
        mock_app = MagicMock()
        mock_dispatch.return_value = mock_app
        service = PowerPointService()
        service.initialize()
        return service


@pytest.fixture
def ppt_service_with_pres(
    ppt_service: PowerPointService, mock_powerpoint_presentation: MagicMock
) -> PowerPointService:
    """Create a PowerPointService with an open presentation."""
    ppt_service._current_document = mock_powerpoint_presentation
    return ppt_service


class TestPowerPointServiceInitialization:
    """Tests for PowerPointService initialization."""

    def test_service_creates_successfully(self, mock_pythoncom: Any) -> None:
        """Test PowerPointService can be instantiated."""
        with patch("win32com.client.Dispatch"):
            service = PowerPointService()
            assert service is not None
            assert not service.is_initialized

    def test_service_initializes(self, ppt_service: PowerPointService) -> None:
        """Test PowerPointService initializes COM application."""
        assert ppt_service.is_initialized
        assert ppt_service.application is not None

    def test_service_cleanup(self, ppt_service: PowerPointService) -> None:
        """Test PowerPointService cleanup."""
        ppt_service.cleanup()
        # Should not raise exception


class TestPresentationCreation:
    """Tests for presentation creation methods."""

    def test_create_presentation(self, ppt_service: PowerPointService) -> None:
        """Test creating a new PowerPoint presentation."""
        mock_pres = MagicMock()
        ppt_service.application.Presentations.Add.return_value = mock_pres

        result = ppt_service.create_presentation()

        assert result["success"] is True
        ppt_service.application.Presentations.Add.assert_called_once()

    def test_open_presentation(
        self, ppt_service: PowerPointService, sample_pptx_path: Path
    ) -> None:
        """Test opening an existing PowerPoint presentation."""
        sample_pptx_path.touch()
        mock_pres = MagicMock()
        mock_pres.FullName = str(sample_pptx_path)
        ppt_service.application.Presentations.Open.return_value = mock_pres

        result = ppt_service.open_presentation(str(sample_pptx_path))

        assert result["success"] is True
        ppt_service.application.Presentations.Open.assert_called_once()

    def test_create_from_template(self, ppt_service: PowerPointService, tmp_path: Path) -> None:
        """Test creating presentation from template."""
        template_path = tmp_path / "template.potx"
        template_path.touch()

        mock_pres = MagicMock()
        ppt_service.application.Presentations.Open.return_value = mock_pres

        result = ppt_service.create_from_template(str(template_path))

        assert result["success"] is True


class TestPresentationOperations:
    """Tests for presentation operations."""

    def test_save_presentation(self, ppt_service_with_pres: PowerPointService) -> None:
        """Test saving a presentation."""
        result = ppt_service_with_pres.save_presentation()

        assert result["success"] is True
        ppt_service_with_pres.current_document.Save.assert_called_once()

    def test_save_presentation_no_doc_open(self, ppt_service: PowerPointService) -> None:
        """Test saving when no presentation is open."""
        with pytest.raises(DocumentNotOpenError):
            ppt_service.save_presentation()

    def test_close_presentation(self, ppt_service_with_pres: PowerPointService) -> None:
        """Test closing a presentation."""
        result = ppt_service_with_pres.close_presentation()

        assert result["success"] is True
        ppt_service_with_pres.current_document.Close.assert_called_once()

    def test_export_to_pdf(self, ppt_service_with_pres: PowerPointService, tmp_path: Path) -> None:
        """Test exporting presentation to PDF."""
        pdf_path = tmp_path / "output.pdf"

        result = ppt_service_with_pres.export_to_pdf(str(pdf_path))

        assert result["success"] is True
        ppt_service_with_pres.current_document.ExportAsFixedFormat.assert_called_once()


class TestSlideOperations:
    """Tests for slide operations."""

    def test_add_slide(self, ppt_service_with_pres: PowerPointService) -> None:
        """Test adding a slide."""
        mock_slide = MagicMock()
        mock_slide.SlideIndex = 1
        ppt_service_with_pres.current_document.Slides.Add.return_value = mock_slide

        result = ppt_service_with_pres.add_slide(layout=2)

        assert result["success"] is True
        assert result["slide_index"] == 1
        ppt_service_with_pres.current_document.Slides.Add.assert_called_once()

    def test_delete_slide(self, ppt_service_with_pres: PowerPointService) -> None:
        """Test deleting a slide."""
        mock_slide = MagicMock()
        ppt_service_with_pres.current_document.Slides.return_value = mock_slide

        result = ppt_service_with_pres.delete_slide(1)

        assert result["success"] is True
        mock_slide.Delete.assert_called_once()

    def test_duplicate_slide(self, ppt_service_with_pres: PowerPointService) -> None:
        """Test duplicating a slide."""
        mock_slide = MagicMock()
        mock_duplicate = MagicMock()
        mock_duplicate.SlideIndex = 2
        mock_slide.Duplicate.return_value = MagicMock()
        mock_slide.Duplicate.return_value.Item.return_value = mock_duplicate
        ppt_service_with_pres.current_document.Slides.return_value = mock_slide

        result = ppt_service_with_pres.duplicate_slide(1)

        assert result["success"] is True

    def test_move_slide(self, ppt_service_with_pres: PowerPointService) -> None:
        """Test moving a slide."""
        mock_slide = MagicMock()
        ppt_service_with_pres.current_document.Slides.return_value = mock_slide

        result = ppt_service_with_pres.move_slide(1, 3)

        assert result["success"] is True
        mock_slide.MoveTo.assert_called_once_with(3)


class TestTextOperations:
    """Tests for text operations."""

    def test_add_textbox(self, ppt_service_with_pres: PowerPointService) -> None:
        """Test adding a textbox to a slide."""
        mock_slide = MagicMock()
        mock_shapes = MagicMock()
        mock_shape = MagicMock()
        mock_shapes.AddTextbox.return_value = mock_shape
        mock_slide.Shapes = mock_shapes
        ppt_service_with_pres.current_document.Slides.return_value = mock_slide

        result = ppt_service_with_pres.add_textbox(1, "Test text", 100, 100, 200, 50)

        assert result["success"] is True
        mock_shapes.AddTextbox.assert_called_once()

    def test_modify_title(self, ppt_service_with_pres: PowerPointService) -> None:
        """Test modifying slide title."""
        mock_slide = MagicMock()
        mock_title = MagicMock()
        mock_text_frame = MagicMock()
        mock_text_range = MagicMock()
        mock_title.TextFrame = mock_text_frame
        mock_text_frame.TextRange = mock_text_range
        mock_slide.Shapes.Title = mock_title
        ppt_service_with_pres.current_document.Slides.return_value = mock_slide

        result = ppt_service_with_pres.modify_title(1, "New Title")

        assert result["success"] is True
        assert mock_text_range.Text == "New Title"

    def test_add_bullets(self, ppt_service_with_pres: PowerPointService) -> None:
        """Test adding bullet points."""
        mock_slide = MagicMock()
        mock_shape = MagicMock()
        mock_text_frame = MagicMock()
        mock_text_range = MagicMock()
        mock_shape.TextFrame = mock_text_frame
        mock_text_frame.TextRange = mock_text_range
        mock_slide.Shapes.return_value = mock_shape
        ppt_service_with_pres.current_document.Slides.return_value = mock_slide

        result = ppt_service_with_pres.add_bullets(1, 1, ["Point 1", "Point 2", "Point 3"])

        assert result["success"] is True


class TestImageOperations:
    """Tests for image operations."""

    def test_insert_image(
        self, ppt_service_with_pres: PowerPointService, sample_image_path: Path
    ) -> None:
        """Test inserting an image."""
        mock_slide = MagicMock()
        mock_shapes = MagicMock()
        mock_slide.Shapes = mock_shapes
        ppt_service_with_pres.current_document.Slides.return_value = mock_slide

        result = ppt_service_with_pres.insert_image(1, str(sample_image_path), 100, 100, 200, 150)

        assert result["success"] is True
        mock_shapes.AddPicture.assert_called_once()

    def test_resize_image(self, ppt_service_with_pres: PowerPointService) -> None:
        """Test resizing an image."""
        mock_slide = MagicMock()
        mock_shape = MagicMock()
        mock_slide.Shapes.return_value = mock_shape
        ppt_service_with_pres.current_document.Slides.return_value = mock_slide

        result = ppt_service_with_pres.resize_image(1, 1, 300, 200)

        assert result["success"] is True
        assert mock_shape.Width == 300
        assert mock_shape.Height == 200

    def test_reposition_image(self, ppt_service_with_pres: PowerPointService) -> None:
        """Test repositioning an image."""
        mock_slide = MagicMock()
        mock_shape = MagicMock()
        mock_slide.Shapes.return_value = mock_shape
        ppt_service_with_pres.current_document.Slides.return_value = mock_slide

        result = ppt_service_with_pres.reposition_image(1, 1, 150, 100)

        assert result["success"] is True
        assert mock_shape.Left == 150
        assert mock_shape.Top == 100


class TestShapeOperations:
    """Tests for shape operations."""

    def test_insert_shape(self, ppt_service_with_pres: PowerPointService) -> None:
        """Test inserting a shape."""
        mock_slide = MagicMock()
        mock_shapes = MagicMock()
        mock_slide.Shapes = mock_shapes
        ppt_service_with_pres.current_document.Slides.return_value = mock_slide

        result = ppt_service_with_pres.insert_shape(1, "rectangle", 100, 100, 200, 150)

        assert result["success"] is True
        mock_shapes.AddShape.assert_called_once()

    def test_modify_fill_color(self, ppt_service_with_pres: PowerPointService) -> None:
        """Test modifying shape fill color."""
        mock_slide = MagicMock()
        mock_shape = MagicMock()
        mock_slide.Shapes.return_value = mock_shape
        ppt_service_with_pres.current_document.Slides.return_value = mock_slide

        result = ppt_service_with_pres.modify_fill_color(1, 1, 255, 0, 0)

        assert result["success"] is True

    def test_group_shapes(self, ppt_service_with_pres: PowerPointService) -> None:
        """Test grouping shapes."""
        mock_slide = MagicMock()
        mock_shapes = MagicMock()
        mock_shape_range = MagicMock()
        mock_shapes.Range.return_value = mock_shape_range
        mock_slide.Shapes = mock_shapes
        ppt_service_with_pres.current_document.Slides.return_value = mock_slide

        result = ppt_service_with_pres.group_shapes(1, [1, 2, 3])

        assert result["success"] is True
        mock_shape_range.Group.assert_called_once()


class TestTableOperations:
    """Tests for table operations."""

    def test_insert_table(self, ppt_service_with_pres: PowerPointService) -> None:
        """Test inserting a table."""
        mock_slide = MagicMock()
        mock_shapes = MagicMock()
        mock_slide.Shapes = mock_shapes
        ppt_service_with_pres.current_document.Slides.return_value = mock_slide

        result = ppt_service_with_pres.insert_table(1, 3, 4, 100, 100, 400, 200)

        assert result["success"] is True
        mock_shapes.AddTable.assert_called_once()

    def test_fill_table_cell(self, ppt_service_with_pres: PowerPointService) -> None:
        """Test filling a table cell."""
        mock_slide = MagicMock()
        mock_shape = MagicMock()
        mock_table = MagicMock()
        mock_cell = MagicMock()
        mock_shape.Table = mock_table
        mock_table.Cell.return_value = mock_cell
        mock_slide.Shapes.return_value = mock_shape
        ppt_service_with_pres.current_document.Slides.return_value = mock_slide

        result = ppt_service_with_pres.fill_table_cell(1, 1, 1, 1, "Cell text")

        assert result["success"] is True


class TestChartOperations:
    """Tests for chart operations."""

    def test_insert_chart(self, ppt_service_with_pres: PowerPointService) -> None:
        """Test inserting a chart."""
        mock_slide = MagicMock()
        mock_shapes = MagicMock()
        mock_slide.Shapes = mock_shapes
        ppt_service_with_pres.current_document.Slides.return_value = mock_slide

        result = ppt_service_with_pres.insert_chart(1, "column", 100, 100, 400, 300)

        assert result["success"] is True
        mock_shapes.AddChart2.assert_called_once()


class TestAnimationOperations:
    """Tests for animation operations."""

    def test_add_entrance_animation(self, ppt_service_with_pres: PowerPointService) -> None:
        """Test adding entrance animation."""
        mock_slide = MagicMock()
        mock_shape = MagicMock()
        mock_timeline = MagicMock()
        mock_slide.Shapes.return_value = mock_shape
        mock_slide.TimeLine = mock_timeline
        ppt_service_with_pres.current_document.Slides.return_value = mock_slide

        result = ppt_service_with_pres.add_entrance_animation(1, 1, "fade")

        assert result["success"] is True

    def test_add_exit_animation(self, ppt_service_with_pres: PowerPointService) -> None:
        """Test adding exit animation."""
        mock_slide = MagicMock()
        mock_shape = MagicMock()
        mock_timeline = MagicMock()
        mock_slide.Shapes.return_value = mock_shape
        mock_slide.TimeLine = mock_timeline
        ppt_service_with_pres.current_document.Slides.return_value = mock_slide

        result = ppt_service_with_pres.add_exit_animation(1, 1, "dissolve")

        assert result["success"] is True

    def test_set_animation_order(self, ppt_service_with_pres: PowerPointService) -> None:
        """Test setting animation order."""
        mock_slide = MagicMock()
        mock_timeline = MagicMock()
        mock_effect = MagicMock()
        mock_timeline.MainSequence.return_value = mock_effect
        mock_slide.TimeLine = mock_timeline
        ppt_service_with_pres.current_document.Slides.return_value = mock_slide

        result = ppt_service_with_pres.set_animation_order(1, 1, 2)

        assert result["success"] is True


class TestTransitionOperations:
    """Tests for transition operations."""

    def test_apply_transition(self, ppt_service_with_pres: PowerPointService) -> None:
        """Test applying a transition."""
        mock_slide = MagicMock()
        ppt_service_with_pres.current_document.Slides.return_value = mock_slide

        result = ppt_service_with_pres.apply_transition(1, "fade")

        assert result["success"] is True

    def test_set_transition_duration(self, ppt_service_with_pres: PowerPointService) -> None:
        """Test setting transition duration."""
        mock_slide = MagicMock()
        ppt_service_with_pres.current_document.Slides.return_value = mock_slide

        result = ppt_service_with_pres.set_transition_duration(1, 2.0)

        assert result["success"] is True

    def test_apply_transition_to_all(self, ppt_service_with_pres: PowerPointService) -> None:
        """Test applying transition to all slides."""
        mock_slide1 = MagicMock()
        mock_slide2 = MagicMock()
        ppt_service_with_pres.current_document.Slides = [mock_slide1, mock_slide2]

        result = ppt_service_with_pres.apply_transition_to_all("push")

        assert result["success"] is True


class TestThemeOperations:
    """Tests for theme operations."""

    def test_apply_theme(self, ppt_service_with_pres: PowerPointService, tmp_path: Path) -> None:
        """Test applying a theme."""
        theme_path = tmp_path / "theme.thmx"
        theme_path.touch()

        result = ppt_service_with_pres.apply_theme(str(theme_path))

        assert result["success"] is True

    def test_set_background(self, ppt_service_with_pres: PowerPointService) -> None:
        """Test setting slide background."""
        mock_slide = MagicMock()
        ppt_service_with_pres.current_document.Slides.return_value = mock_slide

        result = ppt_service_with_pres.set_background(1, 255, 255, 255)

        assert result["success"] is True


class TestAdvancedFeatures:
    """Tests for advanced features."""

    def test_add_speaker_notes(self, ppt_service_with_pres: PowerPointService) -> None:
        """Test adding speaker notes."""
        mock_slide = MagicMock()
        mock_notes_page = MagicMock()
        mock_shapes = MagicMock()
        mock_placeholder = MagicMock()
        mock_text_frame = MagicMock()
        mock_text_range = MagicMock()

        mock_slide.NotesPage = mock_notes_page
        mock_notes_page.Shapes = mock_shapes
        mock_shapes.Placeholders.return_value = mock_placeholder
        mock_placeholder.TextFrame = mock_text_frame
        mock_text_frame.TextRange = mock_text_range
        ppt_service_with_pres.current_document.Slides.return_value = mock_slide

        result = ppt_service_with_pres.add_speaker_notes(1, "Speaker notes")

        assert result["success"] is True

    def test_insert_hyperlink(self, ppt_service_with_pres: PowerPointService) -> None:
        """Test inserting a hyperlink."""
        mock_slide = MagicMock()
        mock_shape = MagicMock()
        mock_text_frame = MagicMock()
        mock_text_range = MagicMock()
        mock_shape.TextFrame = mock_text_frame
        mock_text_frame.TextRange = mock_text_range
        mock_slide.Shapes.return_value = mock_shape
        ppt_service_with_pres.current_document.Slides.return_value = mock_slide

        result = ppt_service_with_pres.insert_hyperlink(1, 1, "https://example.com", "Link text")

        assert result["success"] is True
