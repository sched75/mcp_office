"""Helper utilities for Office automation.

This module provides miscellaneous helper functions that don't fit
into other categories.
"""

from datetime import datetime
from pathlib import Path
from typing import Any


def sanitize_filename(filename: str) -> str:
    """Sanitize a filename by removing invalid characters.

    Args:
        filename: Filename to sanitize

    Returns:
        Sanitized filename
    """
    # Characters not allowed in Windows filenames
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, "_")

    # Remove leading/trailing spaces and dots
    filename = filename.strip(". ")

    return filename or "unnamed"


def ensure_directory_exists(file_path: str | Path) -> Path:
    """Ensure that the directory for a file path exists.

    Args:
        file_path: File path

    Returns:
        Path object for the file
    """
    path = Path(file_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def generate_timestamp_filename(base_name: str, extension: str) -> str:
    """Generate a filename with timestamp.

    Args:
        base_name: Base filename without extension
        extension: File extension (with or without dot)

    Returns:
        Filename with timestamp
    """
    if not extension.startswith("."):
        extension = f".{extension}"

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name = sanitize_filename(base_name)

    return f"{base_name}_{timestamp}{extension}"


def dict_to_result(success: bool = True, message: str = "", **kwargs: Any) -> dict[str, Any]:
    """Create a standardized result dictionary.

    Args:
        success: Whether the operation was successful
        message: Status message
        **kwargs: Additional data to include

    Returns:
        Result dictionary
    """
    result = {
        "success": success,
        "message": message,
        "timestamp": datetime.now().isoformat(),
    }
    result.update(kwargs)
    return result


def parse_range(range_address: str) -> tuple[str, str]:
    """Parse a range address into start and end cells.

    Args:
        range_address: Range address (e.g., "A1:B10")

    Returns:
        Tuple of (start_cell, end_cell)
    """
    parts = range_address.split(":")
    if len(parts) != 2:
        msg = f"Invalid range format: {range_address}"
        raise ValueError(msg)

    return parts[0].strip(), parts[1].strip()


def column_letter_to_number(column: str) -> int:
    """Convert Excel column letter to number.

    Args:
        column: Column letter (e.g., "A", "Z", "AA")

    Returns:
        Column number (1-based)

    Examples:
        A -> 1, B -> 2, Z -> 26, AA -> 27
    """
    number = 0
    for char in column.upper():
        number = number * 26 + (ord(char) - ord("A") + 1)
    return number


def column_number_to_letter(number: int) -> str:
    """Convert Excel column number to letter.

    Args:
        number: Column number (1-based)

    Returns:
        Column letter

    Examples:
        1 -> A, 2 -> B, 26 -> Z, 27 -> AA
    """
    letter = ""
    while number > 0:
        number, remainder = divmod(number - 1, 26)
        letter = chr(65 + remainder) + letter
    return letter


def parse_cell_address(address: str) -> tuple[str, int]:
    """Parse a cell address into column and row.

    Args:
        address: Cell address (e.g., "A1", "B10")

    Returns:
        Tuple of (column_letter, row_number)
    """
    import re

    match = re.match(r"([A-Z]+)([0-9]+)", address.upper())
    if not match:
        msg = f"Invalid cell address: {address}"
        raise ValueError(msg)

    return match.group(1), int(match.group(2))


def points_to_pixels(points: float, dpi: int = 96) -> int:
    """Convert points to pixels.

    Args:
        points: Size in points
        dpi: Dots per inch (default 96 for screen)

    Returns:
        Size in pixels
    """
    return int(points * dpi / 72)


def pixels_to_points(pixels: int, dpi: int = 96) -> float:
    """Convert pixels to points.

    Args:
        pixels: Size in pixels
        dpi: Dots per inch (default 96 for screen)

    Returns:
        Size in points
    """
    return pixels * 72 / dpi


def inches_to_points(inches: float) -> float:
    """Convert inches to points.

    Args:
        inches: Size in inches

    Returns:
        Size in points
    """
    return inches * 72


def points_to_inches(points: float) -> float:
    """Convert points to inches.

    Args:
        points: Size in points

    Returns:
        Size in inches
    """
    return points / 72
