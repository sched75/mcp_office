"""Validation utilities for Office automation parameters.

This module provides validation functions following the Fail-Fast principle
and providing clear error messages.
"""

import re
from pathlib import Path
from typing import Any

from ..core.exceptions import InvalidParameterError


def validate_file_path(
    file_path: str | Path,
    must_exist: bool = False,
    extensions: list[str] | None = None,
) -> Path:
    """Validate a file path.

    Args:
        file_path: Path to validate
        must_exist: Whether the file must already exist
        extensions: Allowed file extensions (e.g., ['.docx', '.doc'])

    Returns:
        Validated Path object

    Raises:
        InvalidParameterError: If validation fails
    """
    if not file_path:
        raise InvalidParameterError("file_path", file_path, "File path cannot be empty")

    path = Path(file_path)

    if must_exist and not path.exists():
        raise InvalidParameterError("file_path", file_path, "File does not exist")

    if extensions and path.suffix.lower() not in [ext.lower() for ext in extensions]:
        raise InvalidParameterError(
            "file_path",
            file_path,
            f"File extension must be one of: {', '.join(extensions)}",
        )

    return path


def validate_cell_address(address: str) -> str:
    """Validate an Excel cell address.

    Args:
        address: Cell address (e.g., "A1", "B10")

    Returns:
        Validated cell address

    Raises:
        InvalidParameterError: If address is invalid
    """
    if not address:
        raise InvalidParameterError("address", address, "Cell address cannot be empty")

    # Pattern: Column letters followed by row number
    pattern = r"^[A-Z]{1,3}[0-9]+$"
    if not re.match(pattern, address.upper()):
        raise InvalidParameterError(
            "address",
            address,
            "Invalid cell address format. Expected format: A1, B10, etc.",
        )

    return address.upper()


def validate_range_address(range_address: str) -> str:
    """Validate an Excel range address.

    Args:
        range_address: Range address (e.g., "A1:B10")

    Returns:
        Validated range address

    Raises:
        InvalidParameterError: If address is invalid
    """
    if not range_address:
        raise InvalidParameterError(
            "range_address", range_address, "Range address cannot be empty"
        )

    # Pattern: Cell:Cell
    pattern = r"^[A-Z]{1,3}[0-9]+:[A-Z]{1,3}[0-9]+$"
    if not re.match(pattern, range_address.upper()):
        raise InvalidParameterError(
            "range_address",
            range_address,
            "Invalid range format. Expected format: A1:B10",
        )

    return range_address.upper()


def validate_rgb_color(r: int, g: int, b: int) -> tuple[int, int, int]:
    """Validate RGB color values.

    Args:
        r: Red component
        g: Green component
        b: Blue component

    Returns:
        Validated RGB tuple

    Raises:
        InvalidParameterError: If any value is out of range
    """
    for name, value in [("r", r), ("g", g), ("b", b)]:
        if not isinstance(value, int) or not 0 <= value <= 255:
            raise InvalidParameterError(
                name, value, "RGB values must be integers between 0 and 255"
            )

    return (r, g, b)


def validate_positive_number(
    name: str, value: int | float, allow_zero: bool = False
) -> int | float:
    """Validate that a number is positive.

    Args:
        name: Parameter name for error messages
        value: Value to validate
        allow_zero: Whether zero is allowed

    Returns:
        Validated value

    Raises:
        InvalidParameterError: If value is invalid
    """
    if not isinstance(value, (int, float)):
        raise InvalidParameterError(name, value, "Value must be a number")

    min_value = 0 if allow_zero else 0.001
    if value < min_value:
        msg = "Value must be positive" + (" or zero" if allow_zero else "")
        raise InvalidParameterError(name, value, msg)

    return value


def validate_percentage(name: str, value: int | float) -> int | float:
    """Validate a percentage value.

    Args:
        name: Parameter name for error messages
        value: Value to validate (0-100)

    Returns:
        Validated value

    Raises:
        InvalidParameterError: If value is out of range
    """
    if not isinstance(value, (int, float)):
        raise InvalidParameterError(name, value, "Value must be a number")

    if not 0 <= value <= 100:
        raise InvalidParameterError(name, value, "Percentage must be between 0 and 100")

    return value


def validate_string_not_empty(name: str, value: str) -> str:
    """Validate that a string is not empty.

    Args:
        name: Parameter name for error messages
        value: String to validate

    Returns:
        Validated string

    Raises:
        InvalidParameterError: If string is empty
    """
    if not isinstance(value, str):
        raise InvalidParameterError(name, value, "Value must be a string")

    if not value.strip():
        raise InvalidParameterError(name, value, "String cannot be empty")

    return value


def validate_dimensions(
    width: int | float | None = None,
    height: int | float | None = None,
    rows: int | None = None,
    cols: int | None = None,
) -> dict[str, int | float]:
    """Validate dimensions for tables, images, etc.

    Args:
        width: Width value
        height: Height value
        rows: Number of rows
        cols: Number of columns

    Returns:
        Dictionary of validated dimensions

    Raises:
        InvalidParameterError: If any dimension is invalid
    """
    result: dict[str, int | float] = {}

    if width is not None:
        result["width"] = validate_positive_number("width", width)

    if height is not None:
        result["height"] = validate_positive_number("height", height)

    if rows is not None:
        if not isinstance(rows, int) or rows < 1:
            raise InvalidParameterError("rows", rows, "Rows must be a positive integer")
        result["rows"] = rows

    if cols is not None:
        if not isinstance(cols, int) or cols < 1:
            raise InvalidParameterError("cols", cols, "Columns must be a positive integer")
        result["cols"] = cols

    return result


def validate_choice(name: str, value: Any, choices: list[Any]) -> Any:
    """Validate that a value is in a list of choices.

    Args:
        name: Parameter name for error messages
        value: Value to validate
        choices: List of allowed values

    Returns:
        Validated value

    Raises:
        InvalidParameterError: If value is not in choices
    """
    if value not in choices:
        raise InvalidParameterError(
            name, value, f"Value must be one of: {', '.join(map(str, choices))}"
        )

    return value
