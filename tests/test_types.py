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
        # Values are auto-generated integers
        assert isinstance(DocumentFormat.DOCX.value, int)

    def test_pdf_value(self) -> None:
        """Test PDF format value."""
        # Values are auto-generated integers
        assert isinstance(DocumentFormat.PDF.value, int)

    def test_has_office_formats(self) -> None:
        """Test enum has Office formats."""
        format_names = [f.name for f in DocumentFormat]
        assert "DOCX" in format_names
        assert "XLSX" in format_names
        assert "PPTX" in format_names

    def test_has_legacy_formats(self) -> None:
        """Test enum has legacy formats."""
        format_names = [f.name for f in DocumentFormat]
        assert "DOC" in format_names
        assert "XLS" in format_names
        assert "PPT" in format_names


class TestTextAlignment:
    """Tests for TextAlignment enum."""

    def test_alignment_values(self) -> None:
        """Test text alignment values."""
        # Values are auto-generated integers
        assert isinstance(TextAlignment.LEFT.value, int)
        assert isinstance(TextAlignment.CENTER.value, int)
        assert isinstance(TextAlignment.RIGHT.value, int)
        assert isinstance(TextAlignment.JUSTIFY.value, int)

    def test_has_all_alignments(self) -> None:
        """Test enum has all standard alignments."""
        assert len(TextAlignment) == 4


class TestVerticalAlignment:
    """Tests for VerticalAlignment enum."""

    def test_alignment_values(self) -> None:
        """Test vertical alignment values."""
        # Values are auto-generated integers
        assert isinstance(VerticalAlignment.TOP.value, int)
        assert isinstance(VerticalAlignment.MIDDLE.value, int)
        assert isinstance(VerticalAlignment.BOTTOM.value, int)

    def test_has_all_alignments(self) -> None:
        """Test enum has all vertical alignments."""
        assert len(VerticalAlignment) == 3


class TestFontStyle:
    """Tests for FontStyle enum."""

    def test_font_styles(self) -> None:
        """Test font style values."""
        # Values are auto-generated integers
        assert isinstance(FontStyle.BOLD.value, int)
        assert isinstance(FontStyle.ITALIC.value, int)
        assert isinstance(FontStyle.UNDERLINE.value, int)

    def test_has_all_styles(self) -> None:
        """Test enum has all font styles."""
        assert len(FontStyle) == 4  # BOLD, ITALIC, UNDERLINE, STRIKETHROUGH


class TestBorderStyle:
    """Tests for BorderStyle enum."""

    def test_border_styles(self) -> None:
        """Test border style values."""
        # Values are auto-generated integers
        assert isinstance(BorderStyle.NONE.value, int)
        assert isinstance(BorderStyle.SINGLE.value, int)

    def test_has_all_styles(self) -> None:
        """Test enum has all border styles."""
        assert len(BorderStyle) >= 4


class TestImagePosition:
    """Tests for ImagePosition enum."""

    def test_position_values(self) -> None:
        """Test image position values."""
        # Values are auto-generated integers
        assert isinstance(ImagePosition.INLINE.value, int)
        assert isinstance(ImagePosition.FLOAT.value, int)
        assert isinstance(ImagePosition.ANCHOR.value, int)

    def test_has_all_positions(self) -> None:
        """Test enum has all position types."""
        assert len(ImagePosition) == 3


class TestChartType:
    """Tests for ChartType enum."""

    def test_basic_chart_types(self) -> None:
        """Test basic chart type values."""
        # Values are auto-generated integers
        assert isinstance(ChartType.COLUMN.value, int)
        assert isinstance(ChartType.BAR.value, int)
        assert isinstance(ChartType.LINE.value, int)
        assert isinstance(ChartType.PIE.value, int)

    def test_has_common_charts(self) -> None:
        """Test enum has common chart types."""
        chart_names = [c.name for c in ChartType]
        assert "COLUMN" in chart_names
        assert "BAR" in chart_names
        assert "LINE" in chart_names
        assert "PIE" in chart_names
        assert "AREA" in chart_names


class TestProtectionType:
    """Tests for ProtectionType enum."""

    def test_protection_types(self) -> None:
        """Test protection type values."""
        # Values are auto-generated integers, NONE doesn't exist
        assert isinstance(ProtectionType.READ_ONLY.value, int)
        assert isinstance(ProtectionType.FORMS.value, int)

    def test_has_all_types(self) -> None:
        """Test enum has protection types."""
        assert len(ProtectionType) >= 3


class TestSlideLayout:
    """Tests for SlideLayout enum."""

    def test_slide_layouts(self) -> None:
        """Test slide layout values."""
        # Values are auto-generated integers, names are different
        assert isinstance(SlideLayout.TITLE_SLIDE.value, int)
        assert isinstance(SlideLayout.TITLE_AND_CONTENT.value, int)
        assert isinstance(SlideLayout.BLANK.value, int)

    def test_has_common_layouts(self) -> None:
        """Test enum has common slide layouts."""
        layouts = list(SlideLayout)
        assert SlideLayout.TITLE_SLIDE in layouts
        assert SlideLayout.TITLE_AND_CONTENT in layouts
        assert SlideLayout.BLANK in layouts


class TestAnimationType:
    """Tests for AnimationType enum."""

    def test_animation_types(self) -> None:
        """Test animation type values."""
        # Values are auto-generated integers
        assert isinstance(AnimationType.ENTRANCE.value, int)
        assert isinstance(AnimationType.EXIT.value, int)
        assert isinstance(AnimationType.EMPHASIS.value, int)

    def test_has_all_types(self) -> None:
        """Test enum has animation types."""
        assert len(AnimationType) >= 3


class TestTransitionType:
    """Tests for TransitionType enum."""

    def test_transition_types(self) -> None:
        """Test transition type values."""
        # Values are auto-generated integers
        assert isinstance(TransitionType.NONE.value, int)
        assert isinstance(TransitionType.FADE.value, int)
        assert isinstance(TransitionType.PUSH.value, int)

    def test_has_common_transitions(self) -> None:
        """Test enum has common transitions."""
        transition_names = [t.name for t in TransitionType]
        assert "FADE" in transition_names
        assert "PUSH" in transition_names
        assert "WIPE" in transition_names
