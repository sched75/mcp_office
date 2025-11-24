"""Unit tests for COM wrapper module."""

from unittest.mock import Mock

import pytest

from src.core.exceptions import COMOperationError
from src.utils.com_wrapper import (
    COMConstants,
    com_safe,
    office_color_to_rgb,
    rgb_to_office_color,
)


class TestComSafeDecorator:
    """Tests for com_safe decorator."""

    def test_com_safe_success(self) -> None:
        """Test decorator with successful operation."""

        @com_safe("test_operation")
        def successful_func() -> str:
            return "success"

        result = successful_func()
        assert result == "success"

    def test_com_safe_with_exception(self) -> None:
        """Test decorator catches and wraps exceptions."""

        @com_safe("test_operation")
        def failing_func() -> None:
            raise ValueError("Test error")

        with pytest.raises(COMOperationError) as exc_info:
            failing_func()

        assert "test_operation" in str(exc_info.value)

    def test_com_safe_preserves_function_name(self) -> None:
        """Test decorator preserves function metadata."""

        @com_safe("test_operation")
        def my_function() -> None:
            """My docstring."""
            pass

        assert my_function.__name__ == "my_function"
        assert my_function.__doc__ == "My docstring."

    def test_com_safe_with_args_kwargs(self) -> None:
        """Test decorator works with args and kwargs."""

        @com_safe("test_operation")
        def func_with_params(a: int, b: int, c: int = 0) -> int:
            return a + b + c

        result = func_with_params(1, 2, c=3)
        assert result == 6


class TestCOMConstants:
    """Tests for COMConstants class."""

    def test_word_constants_exist(self) -> None:
        """Test Word constants are defined."""
        assert hasattr(COMConstants, "WD_ALIGN_PARAGRAPH_LEFT")
        assert hasattr(COMConstants, "WD_ALIGN_PARAGRAPH_CENTER")
        assert hasattr(COMConstants, "WD_ALIGN_PARAGRAPH_RIGHT")

    def test_excel_constants_exist(self) -> None:
        """Test Excel constants are defined."""
        assert hasattr(COMConstants, "XL_HALIGN_CENTER")
        assert hasattr(COMConstants, "XL_HALIGN_LEFT")
        assert hasattr(COMConstants, "XL_HALIGN_RIGHT")

    def test_get_word_alignment_valid(self) -> None:
        """Test get_word_alignment with valid input."""
        assert COMConstants.get_word_alignment("left") == COMConstants.WD_ALIGN_PARAGRAPH_LEFT
        assert COMConstants.get_word_alignment("center") == COMConstants.WD_ALIGN_PARAGRAPH_CENTER
        assert COMConstants.get_word_alignment("right") == COMConstants.WD_ALIGN_PARAGRAPH_RIGHT
        assert COMConstants.get_word_alignment("justify") == COMConstants.WD_ALIGN_PARAGRAPH_JUSTIFY

    def test_get_word_alignment_case_insensitive(self) -> None:
        """Test get_word_alignment is case insensitive."""
        assert COMConstants.get_word_alignment("LEFT") == COMConstants.WD_ALIGN_PARAGRAPH_LEFT
        assert COMConstants.get_word_alignment("Center") == COMConstants.WD_ALIGN_PARAGRAPH_CENTER

    def test_get_word_alignment_invalid(self) -> None:
        """Test get_word_alignment with invalid input."""
        with pytest.raises(ValueError, match="Invalid alignment"):
            COMConstants.get_word_alignment("invalid")

    def test_get_excel_halignment_valid(self) -> None:
        """Test get_excel_halignment with valid input."""
        assert COMConstants.get_excel_halignment("left") == COMConstants.XL_HALIGN_LEFT
        assert COMConstants.get_excel_halignment("center") == COMConstants.XL_HALIGN_CENTER
        assert COMConstants.get_excel_halignment("right") == COMConstants.XL_HALIGN_RIGHT

    def test_get_excel_halignment_invalid(self) -> None:
        """Test get_excel_halignment with invalid input."""
        with pytest.raises(ValueError, match="Invalid horizontal alignment"):
            COMConstants.get_excel_halignment("invalid")

    def test_get_excel_valignment_valid(self) -> None:
        """Test get_excel_valignment with valid input."""
        assert COMConstants.get_excel_valignment("top") == COMConstants.XL_VALIGN_TOP
        assert COMConstants.get_excel_valignment("center") == COMConstants.XL_VALIGN_CENTER
        assert COMConstants.get_excel_valignment("bottom") == COMConstants.XL_VALIGN_BOTTOM

    def test_get_excel_valignment_invalid(self) -> None:
        """Test get_excel_valignment with invalid input."""
        with pytest.raises(ValueError, match="Invalid vertical alignment"):
            COMConstants.get_excel_valignment("invalid")

    def test_get_chart_type_valid(self) -> None:
        """Test get_chart_type with valid input."""
        assert COMConstants.get_chart_type("column") == COMConstants.XL_COLUMN_CLUSTERED
        assert COMConstants.get_chart_type("bar") == COMConstants.XL_BAR_CLUSTERED
        assert COMConstants.get_chart_type("line") == COMConstants.XL_LINE
        assert COMConstants.get_chart_type("pie") == COMConstants.XL_PIE

    def test_get_chart_type_invalid(self) -> None:
        """Test get_chart_type with invalid input."""
        with pytest.raises(ValueError, match="Invalid chart type"):
            COMConstants.get_chart_type("invalid")


class TestRGBToOfficeColor:
    """Tests for rgb_to_office_color function."""

    def test_rgb_to_office_color_black(self) -> None:
        """Test conversion of black color."""
        result = rgb_to_office_color(0, 0, 0)
        assert result == 0

    def test_rgb_to_office_color_white(self) -> None:
        """Test conversion of white color."""
        result = rgb_to_office_color(255, 255, 255)
        assert result == 0xFFFFFF

    def test_rgb_to_office_color_red(self) -> None:
        """Test conversion of red color."""
        result = rgb_to_office_color(255, 0, 0)
        assert result == 0x0000FF  # BGR format

    def test_rgb_to_office_color_green(self) -> None:
        """Test conversion of green color."""
        result = rgb_to_office_color(0, 255, 0)
        assert result == 0x00FF00

    def test_rgb_to_office_color_blue(self) -> None:
        """Test conversion of blue color."""
        result = rgb_to_office_color(0, 0, 255)
        assert result == 0xFF0000  # BGR format

    def test_rgb_to_office_color_custom(self) -> None:
        """Test conversion of custom color."""
        result = rgb_to_office_color(128, 64, 192)
        assert result == 0xC04080  # BGR: 192, 64, 128


class TestOfficeColorToRGB:
    """Tests for office_color_to_rgb function."""

    def test_office_color_to_rgb_black(self) -> None:
        """Test conversion of black color."""
        r, g, b = office_color_to_rgb(0)
        assert (r, g, b) == (0, 0, 0)

    def test_office_color_to_rgb_white(self) -> None:
        """Test conversion of white color."""
        r, g, b = office_color_to_rgb(0xFFFFFF)
        assert (r, g, b) == (255, 255, 255)

    def test_office_color_to_rgb_red(self) -> None:
        """Test conversion of red color (BGR 0x0000FF)."""
        r, g, b = office_color_to_rgb(0x0000FF)
        assert (r, g, b) == (255, 0, 0)

    def test_office_color_to_rgb_green(self) -> None:
        """Test conversion of green color."""
        r, g, b = office_color_to_rgb(0x00FF00)
        assert (r, g, b) == (0, 255, 0)

    def test_office_color_to_rgb_blue(self) -> None:
        """Test conversion of blue color (BGR 0xFF0000)."""
        r, g, b = office_color_to_rgb(0xFF0000)
        assert (r, g, b) == (0, 0, 255)

    def test_office_color_to_rgb_custom(self) -> None:
        """Test conversion of custom color."""
        r, g, b = office_color_to_rgb(0xC04080)
        assert (r, g, b) == (128, 64, 192)


class TestColorConversionRoundTrip:
    """Tests for round-trip color conversions."""

    def test_rgb_office_rgb_round_trip(self) -> None:
        """Test round trip RGB -> Office -> RGB."""
        original_rgb = (100, 150, 200)
        office_color = rgb_to_office_color(*original_rgb)
        result_rgb = office_color_to_rgb(office_color)
        assert result_rgb == original_rgb

    def test_office_rgb_office_round_trip(self) -> None:
        """Test round trip Office -> RGB -> Office."""
        original_office = 0xABCDEF
        rgb = office_color_to_rgb(original_office)
        result_office = rgb_to_office_color(*rgb)
        assert result_office == original_office
