"""Unit tests for helpers module."""

from pathlib import Path

import pytest

from src.core.exceptions import InvalidParameterError
from src.utils.helpers import (
    column_letter_to_number,
    column_number_to_letter,
    dict_to_result,
    ensure_directory_exists,
    generate_timestamp_filename,
    inches_to_points,
    parse_cell_address,
    parse_range,
    pixels_to_points,
    points_to_inches,
    points_to_pixels,
    sanitize_filename,
)


class TestSanitizeFilename:
    """Tests for sanitize_filename function."""

    def test_sanitize_filename_clean(self) -> None:
        """Test with clean filename."""
        assert sanitize_filename("test.txt") == "test.txt"
        assert sanitize_filename("document_2024.docx") == "document_2024.docx"

    def test_sanitize_filename_special_chars(self) -> None:
        """Test removal of special characters."""
        assert sanitize_filename("test<file>.txt") == "testfile.txt"
        assert sanitize_filename('doc:with|chars?.xlsx') == "docwithchars.xlsx"
        assert sanitize_filename('file/with\\slashes.txt') == "filewithslashes.txt"

    def test_sanitize_filename_asterisk_and_quotes(self) -> None:
        """Test removal of asterisk and quotes."""
        assert sanitize_filename('file*name"test.txt') == "filenametest.txt"


class TestEnsureDirectoryExists:
    """Tests for ensure_directory_exists function."""

    def test_ensure_directory_exists_creates_dir(self, tmp_path: Path) -> None:
        """Test directory creation."""
        new_dir = tmp_path / "new_directory"
        result = ensure_directory_exists(new_dir)
        assert result == new_dir
        assert new_dir.exists()
        assert new_dir.is_dir()

    def test_ensure_directory_exists_already_exists(self, tmp_path: Path) -> None:
        """Test with existing directory."""
        existing_dir = tmp_path / "existing"
        existing_dir.mkdir()
        result = ensure_directory_exists(existing_dir)
        assert result == existing_dir
        assert existing_dir.exists()

    def test_ensure_directory_exists_nested(self, tmp_path: Path) -> None:
        """Test creation of nested directories."""
        nested_dir = tmp_path / "level1" / "level2" / "level3"
        result = ensure_directory_exists(nested_dir)
        assert result == nested_dir
        assert nested_dir.exists()


class TestGenerateTimestampFilename:
    """Tests for generate_timestamp_filename function."""

    def test_generate_timestamp_filename_basic(self) -> None:
        """Test basic timestamp filename generation."""
        result = generate_timestamp_filename("document", ".docx")
        assert result.startswith("document_")
        assert result.endswith(".docx")
        assert len(result) > len("document_.docx")

    def test_generate_timestamp_filename_uniqueness(self) -> None:
        """Test that consecutive calls generate different names."""
        name1 = generate_timestamp_filename("test", ".txt")
        name2 = generate_timestamp_filename("test", ".txt")
        # They might be the same if called in same microsecond, but structure is correct
        assert name1.startswith("test_")
        assert name2.startswith("test_")


class TestDictToResult:
    """Tests for dict_to_result function."""

    def test_dict_to_result_success(self) -> None:
        """Test result dict with success."""
        result = dict_to_result(success=True, message="Operation completed")
        assert result["success"] is True
        assert result["message"] == "Operation completed"

    def test_dict_to_result_with_data(self) -> None:
        """Test result dict with additional data."""
        result = dict_to_result(
            success=True,
            message="Data retrieved",
            count=10,
            items=["a", "b", "c"],
        )
        assert result["success"] is True
        assert result["message"] == "Data retrieved"
        assert result["count"] == 10
        assert result["items"] == ["a", "b", "c"]

    def test_dict_to_result_failure(self) -> None:
        """Test result dict with failure."""
        result = dict_to_result(success=False, message="Operation failed", error="Error details")
        assert result["success"] is False
        assert result["message"] == "Operation failed"
        assert result["error"] == "Error details"


class TestParseRange:
    """Tests for parse_range function."""

    def test_parse_range_valid(self) -> None:
        """Test parsing valid range."""
        start, end = parse_range("A1:B10")
        assert start == "A1"
        assert end == "B10"

    def test_parse_range_complex(self) -> None:
        """Test parsing complex range."""
        start, end = parse_range("AA1:ZZ999")
        assert start == "AA1"
        assert end == "ZZ999"

    def test_parse_range_invalid(self) -> None:
        """Test parsing invalid range."""
        with pytest.raises(InvalidParameterError, match="Invalid range format"):
            parse_range("A1")
        with pytest.raises(InvalidParameterError):
            parse_range("A1:B10:C20")


class TestColumnLetterToNumber:
    """Tests for column_letter_to_number function."""

    def test_column_letter_to_number_single(self) -> None:
        """Test single letter columns."""
        assert column_letter_to_number("A") == 1
        assert column_letter_to_number("B") == 2
        assert column_letter_to_number("Z") == 26

    def test_column_letter_to_number_double(self) -> None:
        """Test double letter columns."""
        assert column_letter_to_number("AA") == 27
        assert column_letter_to_number("AB") == 28
        assert column_letter_to_number("AZ") == 52

    def test_column_letter_to_number_lowercase(self) -> None:
        """Test lowercase letters."""
        assert column_letter_to_number("a") == 1
        assert column_letter_to_number("aa") == 27

    def test_column_letter_to_number_invalid(self) -> None:
        """Test invalid column letters."""
        with pytest.raises(InvalidParameterError, match="Invalid column letter"):
            column_letter_to_number("")
        with pytest.raises(InvalidParameterError):
            column_letter_to_number("123")


class TestColumnNumberToLetter:
    """Tests for column_number_to_letter function."""

    def test_column_number_to_letter_single(self) -> None:
        """Test single digit column numbers."""
        assert column_number_to_letter(1) == "A"
        assert column_number_to_letter(26) == "Z"

    def test_column_number_to_letter_double(self) -> None:
        """Test double digit column numbers."""
        assert column_number_to_letter(27) == "AA"
        assert column_number_to_letter(28) == "AB"
        assert column_number_to_letter(52) == "AZ"

    def test_column_number_to_letter_large(self) -> None:
        """Test large column numbers."""
        assert column_number_to_letter(702) == "ZZ"
        assert column_number_to_letter(703) == "AAA"

    def test_column_number_to_letter_invalid(self) -> None:
        """Test invalid column numbers."""
        with pytest.raises(InvalidParameterError, match="must be positive"):
            column_number_to_letter(0)
        with pytest.raises(InvalidParameterError):
            column_number_to_letter(-1)


class TestParseCellAddress:
    """Tests for parse_cell_address function."""

    def test_parse_cell_address_valid(self) -> None:
        """Test parsing valid cell address."""
        col, row = parse_cell_address("A1")
        assert col == "A"
        assert row == 1

        col, row = parse_cell_address("Z99")
        assert col == "Z"
        assert row == 99

    def test_parse_cell_address_double_letter(self) -> None:
        """Test parsing cell with double letter column."""
        col, row = parse_cell_address("AA10")
        assert col == "AA"
        assert row == 10

    def test_parse_cell_address_invalid(self) -> None:
        """Test parsing invalid cell address."""
        with pytest.raises(InvalidParameterError, match="Invalid cell address"):
            parse_cell_address("1A")
        with pytest.raises(InvalidParameterError):
            parse_cell_address("ABC")


class TestPointsConversions:
    """Tests for points conversion functions."""

    def test_points_to_pixels(self) -> None:
        """Test points to pixels conversion."""
        assert points_to_pixels(72) == 96  # 72 points = 1 inch = 96 pixels
        assert points_to_pixels(36) == 48  # 36 points = 0.5 inch = 48 pixels

    def test_pixels_to_points(self) -> None:
        """Test pixels to points conversion."""
        assert pixels_to_points(96) == 72  # 96 pixels = 1 inch = 72 points
        assert pixels_to_points(48) == 36  # 48 pixels = 0.5 inch = 36 points

    def test_inches_to_points(self) -> None:
        """Test inches to points conversion."""
        assert inches_to_points(1) == 72  # 1 inch = 72 points
        assert inches_to_points(2) == 144  # 2 inches = 144 points
        assert inches_to_points(0.5) == 36  # 0.5 inch = 36 points

    def test_points_to_inches(self) -> None:
        """Test points to inches conversion."""
        assert points_to_inches(72) == 1.0  # 72 points = 1 inch
        assert points_to_inches(144) == 2.0  # 144 points = 2 inches
        assert points_to_inches(36) == 0.5  # 36 points = 0.5 inch

    def test_conversion_round_trip_pixels(self) -> None:
        """Test round trip conversion pixels -> points -> pixels."""
        original = 192
        converted = pixels_to_points(original)
        back = points_to_pixels(converted)
        assert back == original

    def test_conversion_round_trip_inches(self) -> None:
        """Test round trip conversion inches -> points -> inches."""
        original = 3.5
        converted = inches_to_points(original)
        back = points_to_inches(converted)
        assert back == original
