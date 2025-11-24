"""Unit tests for validators module."""

from pathlib import Path

import pytest

from src.core.exceptions import InvalidParameterError
from src.utils.validators import (
    validate_cell_address,
    validate_choice,
    validate_dimensions,
    validate_file_path,
    validate_percentage,
    validate_positive_number,
    validate_range_address,
    validate_rgb_color,
    validate_string_not_empty,
)


class TestValidateFilePath:
    """Tests for validate_file_path function."""

    def test_validate_file_path_with_string(self, tmp_path: Path) -> None:
        """Test validation with string path."""
        test_file = tmp_path / "test.txt"
        test_file.touch()
        result = validate_file_path(str(test_file), must_exist=True)
        assert result == test_file

    def test_validate_file_path_with_path_object(self, tmp_path: Path) -> None:
        """Test validation with Path object."""
        test_file = tmp_path / "test.txt"
        test_file.touch()
        result = validate_file_path(test_file, must_exist=True)
        assert result == test_file

    def test_validate_file_path_nonexistent_no_requirement(self, tmp_path: Path) -> None:
        """Test validation when file doesn't exist but not required."""
        test_file = tmp_path / "nonexistent.txt"
        result = validate_file_path(test_file, must_exist=False)
        assert result == test_file

    def test_validate_file_path_nonexistent_required(self, tmp_path: Path) -> None:
        """Test validation fails when file must exist but doesn't."""
        test_file = tmp_path / "nonexistent.txt"
        with pytest.raises(InvalidParameterError, match="does not exist"):
            validate_file_path(test_file, must_exist=True)

    def test_validate_file_path_with_valid_extension(self, tmp_path: Path) -> None:
        """Test validation with valid file extension."""
        test_file = tmp_path / "test.docx"
        test_file.touch()
        result = validate_file_path(test_file, must_exist=True, extensions=[".docx", ".doc"])
        assert result == test_file

    def test_validate_file_path_with_invalid_extension(self, tmp_path: Path) -> None:
        """Test validation fails with invalid extension."""
        test_file = tmp_path / "test.txt"
        test_file.touch()
        with pytest.raises(InvalidParameterError, match="must have one of these extensions"):
            validate_file_path(test_file, must_exist=True, extensions=[".docx", ".doc"])

    def test_validate_file_path_empty_string(self) -> None:
        """Test validation fails with empty string."""
        with pytest.raises(InvalidParameterError, match="empty"):
            validate_file_path("")


class TestValidateCellAddress:
    """Tests for validate_cell_address function."""

    def test_validate_cell_address_valid_single_letter(self) -> None:
        """Test validation with valid single letter column."""
        assert validate_cell_address("A1") == "A1"
        assert validate_cell_address("Z99") == "Z99"

    def test_validate_cell_address_valid_double_letter(self) -> None:
        """Test validation with valid double letter column."""
        assert validate_cell_address("AA1") == "AA1"
        assert validate_cell_address("ZZ999") == "ZZ999"

    def test_validate_cell_address_lowercase(self) -> None:
        """Test validation converts lowercase to uppercase."""
        assert validate_cell_address("a1") == "A1"
        assert validate_cell_address("aa100") == "AA100"

    def test_validate_cell_address_invalid_format(self) -> None:
        """Test validation fails with invalid format."""
        with pytest.raises(InvalidParameterError, match="Invalid cell address"):
            validate_cell_address("1A")
        with pytest.raises(InvalidParameterError):
            validate_cell_address("ABC")
        with pytest.raises(InvalidParameterError):
            validate_cell_address("123")

    def test_validate_cell_address_empty(self) -> None:
        """Test validation fails with empty string."""
        with pytest.raises(InvalidParameterError):
            validate_cell_address("")


class TestValidateRangeAddress:
    """Tests for validate_range_address function."""

    def test_validate_range_address_valid(self) -> None:
        """Test validation with valid range."""
        assert validate_range_address("A1:B10") == "A1:B10"
        assert validate_range_address("AA1:ZZ999") == "AA1:ZZ999"

    def test_validate_range_address_lowercase(self) -> None:
        """Test validation converts lowercase to uppercase."""
        assert validate_range_address("a1:b10") == "A1:B10"

    def test_validate_range_address_invalid_format(self) -> None:
        """Test validation fails with invalid format."""
        with pytest.raises(InvalidParameterError, match="Invalid range address"):
            validate_range_address("A1")
        with pytest.raises(InvalidParameterError):
            validate_range_address("A1:B")
        with pytest.raises(InvalidParameterError):
            validate_range_address("1:10")

    def test_validate_range_address_empty(self) -> None:
        """Test validation fails with empty string."""
        with pytest.raises(InvalidParameterError):
            validate_range_address("")


class TestValidateRGBColor:
    """Tests for validate_rgb_color function."""

    def test_validate_rgb_color_valid(self) -> None:
        """Test validation with valid RGB values."""
        assert validate_rgb_color(0, 0, 0) == (0, 0, 0)
        assert validate_rgb_color(255, 255, 255) == (255, 255, 255)
        assert validate_rgb_color(128, 64, 192) == (128, 64, 192)

    def test_validate_rgb_color_negative(self) -> None:
        """Test validation fails with negative values."""
        with pytest.raises(InvalidParameterError, match="between 0 and 255"):
            validate_rgb_color(-1, 0, 0)
        with pytest.raises(InvalidParameterError):
            validate_rgb_color(0, -1, 0)
        with pytest.raises(InvalidParameterError):
            validate_rgb_color(0, 0, -1)

    def test_validate_rgb_color_too_large(self) -> None:
        """Test validation fails with values > 255."""
        with pytest.raises(InvalidParameterError, match="between 0 and 255"):
            validate_rgb_color(256, 0, 0)
        with pytest.raises(InvalidParameterError):
            validate_rgb_color(0, 256, 0)
        with pytest.raises(InvalidParameterError):
            validate_rgb_color(0, 0, 256)


class TestValidatePositiveNumber:
    """Tests for validate_positive_number function."""

    def test_validate_positive_number_valid_int(self) -> None:
        """Test validation with valid positive integers."""
        assert validate_positive_number(1, "test") == 1
        assert validate_positive_number(100, "test") == 100

    def test_validate_positive_number_valid_float(self) -> None:
        """Test validation with valid positive floats."""
        assert validate_positive_number(0.1, "test") == 0.1
        assert validate_positive_number(99.99, "test") == 99.99

    def test_validate_positive_number_zero_allowed(self) -> None:
        """Test validation with zero when allowed."""
        assert validate_positive_number(0, "test", allow_zero=True) == 0

    def test_validate_positive_number_zero_not_allowed(self) -> None:
        """Test validation fails with zero when not allowed."""
        with pytest.raises(InvalidParameterError, match="must be positive"):
            validate_positive_number(0, "test", allow_zero=False)

    def test_validate_positive_number_negative(self) -> None:
        """Test validation fails with negative number."""
        with pytest.raises(InvalidParameterError, match="must be positive"):
            validate_positive_number(-1, "test")
        with pytest.raises(InvalidParameterError):
            validate_positive_number(-0.1, "test")


class TestValidatePercentage:
    """Tests for validate_percentage function."""

    def test_validate_percentage_valid(self) -> None:
        """Test validation with valid percentages."""
        assert validate_percentage(0) == 0
        assert validate_percentage(50) == 50
        assert validate_percentage(100) == 100

    def test_validate_percentage_negative(self) -> None:
        """Test validation fails with negative percentage."""
        with pytest.raises(InvalidParameterError, match="between 0 and 100"):
            validate_percentage(-1)

    def test_validate_percentage_too_large(self) -> None:
        """Test validation fails with percentage > 100."""
        with pytest.raises(InvalidParameterError, match="between 0 and 100"):
            validate_percentage(101)


class TestValidateStringNotEmpty:
    """Tests for validate_string_not_empty function."""

    def test_validate_string_not_empty_valid(self) -> None:
        """Test validation with non-empty string."""
        assert validate_string_not_empty("test", "param") == "test"
        assert validate_string_not_empty("  spaces  ", "param") == "  spaces  "

    def test_validate_string_not_empty_empty(self) -> None:
        """Test validation fails with empty string."""
        with pytest.raises(InvalidParameterError, match="cannot be empty"):
            validate_string_not_empty("", "param")

    def test_validate_string_not_empty_whitespace_only(self) -> None:
        """Test validation fails with whitespace-only string."""
        with pytest.raises(InvalidParameterError, match="cannot be empty"):
            validate_string_not_empty("   ", "param")


class TestValidateDimensions:
    """Tests for validate_dimensions function."""

    def test_validate_dimensions_valid(self) -> None:
        """Test validation with valid dimensions."""
        width, height = validate_dimensions(100, 200)
        assert width == 100
        assert height == 200

    def test_validate_dimensions_optional_height(self) -> None:
        """Test validation with optional height."""
        width, height = validate_dimensions(100, None)
        assert width == 100
        assert height is None

    def test_validate_dimensions_invalid_width(self) -> None:
        """Test validation fails with invalid width."""
        with pytest.raises(InvalidParameterError, match="Width must be positive"):
            validate_dimensions(0, 100)
        with pytest.raises(InvalidParameterError):
            validate_dimensions(-1, 100)

    def test_validate_dimensions_invalid_height(self) -> None:
        """Test validation fails with invalid height."""
        with pytest.raises(InvalidParameterError, match="Height must be positive"):
            validate_dimensions(100, 0)
        with pytest.raises(InvalidParameterError):
            validate_dimensions(100, -1)


class TestValidateChoice:
    """Tests for validate_choice function."""

    def test_validate_choice_valid(self) -> None:
        """Test validation with valid choice."""
        choices = ["option1", "option2", "option3"]
        assert validate_choice("option1", choices, "param") == "option1"
        assert validate_choice("option2", choices, "param") == "option2"

    def test_validate_choice_invalid(self) -> None:
        """Test validation fails with invalid choice."""
        choices = ["option1", "option2", "option3"]
        with pytest.raises(InvalidParameterError, match="must be one of"):
            validate_choice("invalid", choices, "param")

    def test_validate_choice_case_sensitive(self) -> None:
        """Test validation is case sensitive by default."""
        choices = ["Option1", "Option2"]
        with pytest.raises(InvalidParameterError):
            validate_choice("option1", choices, "param")
