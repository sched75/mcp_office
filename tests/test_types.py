"""Unit tests for types module."""

from src.core.types import (
    AnimationType,
    ApplicationType,
    BorderStyle,
    ChartType,
    DocumentFormat,
    FontStyle,
    ImagePosition,
    ProtectionType,
    SlideLayout,
    TextAlignment,
    TransitionType,
    VerticalAlignment,
)


class TestApplicationType:
    """Tests for ApplicationType enum."""

    def test_word_value(self) -> None:
        """Test Word application type value."""
        assert ApplicationType.WORD.value == "Word.Application"

    def test_excel_value(self) -> None:
        """Test Excel application type value."""
        assert ApplicationType.EXCEL.value == "Excel.Application"

    def test_powerpoint_value(self) -> None:
        """Test PowerPoint application type value."""
        assert ApplicationType.POWERPOINT.value == "PowerPoint.Application"

    def test_enum_members(self) -> None:
        """Test enum has expected members."""
        assert len(ApplicationType) == 3
        assert ApplicationType.WORD in ApplicationType
        assert ApplicationType.EXCEL in ApplicationType
        assert ApplicationType.POWERPOINT in ApplicationType


class TestDocumentFormat:
    """Tests for DocumentFormat enum."""

    def test_docx_value(self) -> None:
        """Test DOCX format value."""
        assert DocumentFormat.DOCX.value == "docx"

    def test_pdf_value(self) -> None:
        """Test PDF format value."""
        assert DocumentFormat.PDF.value == "pdf"

    def test_has_office_formats(self) -> None:
        """Test enum has Office formats."""
        formats = [f.value for f in DocumentFormat]
        assert "docx" in formats
        assert "xlsx" in formats
        assert "pptx" in formats

    def test_has_legacy_formats(self) -> None:
        """Test enum has legacy formats."""
        formats = [f.value for f in DocumentFormat]
        assert "doc" in formats
        assert "xls" in formats
        assert "ppt" in formats


class TestTextAlignment:
    """Tests for TextAlignment enum."""

    def test_alignment_values(self) -> None:
        """Test text alignment values."""
        assert TextAlignment.LEFT.value == "left"
        assert TextAlignment.CENTER.value == "center"
        assert TextAlignment.RIGHT.value == "right"
        assert TextAlignment.JUSTIFY.value == "justify"

    def test_has_all_alignments(self) -> None:
        """Test enum has all standard alignments."""
        assert len(TextAlignment) == 4


class TestVerticalAlignment:
    """Tests for VerticalAlignment enum."""

    def test_alignment_values(self) -> None:
        """Test vertical alignment values."""
        assert VerticalAlignment.TOP.value == "top"
        assert VerticalAlignment.MIDDLE.value == "middle"
        assert VerticalAlignment.BOTTOM.value == "bottom"

    def test_has_all_alignments(self) -> None:
        """Test enum has all vertical alignments."""
        assert len(VerticalAlignment) == 3


class TestFontStyle:
    """Tests for FontStyle enum."""

    def test_font_styles(self) -> None:
        """Test font style values."""
        assert FontStyle.BOLD.value == "bold"
        assert FontStyle.ITALIC.value == "italic"
        assert FontStyle.UNDERLINE.value == "underline"

    def test_has_all_styles(self) -> None:
        """Test enum has all font styles."""
        assert len(FontStyle) == 3


class TestBorderStyle:
    """Tests for BorderStyle enum."""

    def test_border_styles(self) -> None:
        """Test border style values."""
        assert BorderStyle.NONE.value == "none"
        assert BorderStyle.THIN.value == "thin"
        assert BorderStyle.MEDIUM.value == "medium"
        assert BorderStyle.THICK.value == "thick"

    def test_has_all_styles(self) -> None:
        """Test enum has all border styles."""
        assert len(BorderStyle) >= 4


class TestImagePosition:
    """Tests for ImagePosition enum."""

    def test_position_values(self) -> None:
        """Test image position values."""
        assert ImagePosition.INLINE.value == "inline"
        assert ImagePosition.ABSOLUTE.value == "absolute"
        assert ImagePosition.RELATIVE.value == "relative"

    def test_has_all_positions(self) -> None:
        """Test enum has all position types."""
        assert len(ImagePosition) == 3


class TestChartType:
    """Tests for ChartType enum."""

    def test_basic_chart_types(self) -> None:
        """Test basic chart type values."""
        assert ChartType.COLUMN.value == "column"
        assert ChartType.BAR.value == "bar"
        assert ChartType.LINE.value == "line"
        assert ChartType.PIE.value == "pie"

    def test_has_common_charts(self) -> None:
        """Test enum has common chart types."""
        chart_values = [c.value for c in ChartType]
        assert "column" in chart_values
        assert "bar" in chart_values
        assert "line" in chart_values
        assert "pie" in chart_values
        assert "area" in chart_values


class TestProtectionType:
    """Tests for ProtectionType enum."""

    def test_protection_types(self) -> None:
        """Test protection type values."""
        assert ProtectionType.NONE.value == "none"
        assert ProtectionType.READ_ONLY.value == "read_only"
        assert ProtectionType.FORMS.value == "forms"

    def test_has_all_types(self) -> None:
        """Test enum has protection types."""
        assert len(ProtectionType) >= 3


class TestSlideLayout:
    """Tests for SlideLayout enum."""

    def test_slide_layouts(self) -> None:
        """Test slide layout values."""
        assert SlideLayout.TITLE.value == 1
        assert SlideLayout.TITLE_AND_CONTENT.value == 2
        assert SlideLayout.BLANK.value == 7

    def test_has_common_layouts(self) -> None:
        """Test enum has common slide layouts."""
        layouts = [layout for layout in SlideLayout]
        assert SlideLayout.TITLE in layouts
        assert SlideLayout.TITLE_AND_CONTENT in layouts
        assert SlideLayout.BLANK in layouts


class TestAnimationType:
    """Tests for AnimationType enum."""

    def test_animation_types(self) -> None:
        """Test animation type values."""
        assert AnimationType.ENTRANCE.value == "entrance"
        assert AnimationType.EXIT.value == "exit"
        assert AnimationType.EMPHASIS.value == "emphasis"

    def test_has_all_types(self) -> None:
        """Test enum has animation types."""
        assert len(AnimationType) >= 3


class TestTransitionType:
    """Tests for TransitionType enum."""

    def test_transition_types(self) -> None:
        """Test transition type values."""
        assert TransitionType.NONE.value == "none"
        assert TransitionType.FADE.value == "fade"
        assert TransitionType.PUSH.value == "push"

    def test_has_common_transitions(self) -> None:
        """Test enum has common transitions."""
        transitions = [t.value for t in TransitionType]
        assert "fade" in transitions
        assert "push" in transitions
        assert "wipe" in transitions
