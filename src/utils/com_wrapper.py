"""COM wrapper utilities for safe Office COM operations.

This module provides wrapper functions and decorators to handle COM operations
safely with proper error handling and resource management.
"""

import functools
from collections.abc import Callable
from typing import Any, TypeVar

from ..core.exceptions import COMOperationError

T = TypeVar("T")


def com_safe(operation_name: str | None = None) -> Callable[[Callable[..., T]], Callable[..., T]]:
    """Decorator to wrap COM operations with error handling.

    Args:
        operation_name: Name of the operation for error messages

    Returns:
        Decorated function with COM error handling
    """

    def decorator(func: Callable[..., T]) -> Callable[..., T]:
        @functools.wraps(func)
        def wrapper(*args: Any, **kwargs: Any) -> T:
            op_name = operation_name or func.__name__
            try:
                return func(*args, **kwargs)
            except Exception as e:
                raise COMOperationError(op_name, e) from e

        return wrapper

    return decorator


class COMConstants:
    """Wrapper for accessing Office COM constants.

    This class provides a clean interface to Office constants,
    following the Dependency Inversion Principle.
    """

    # Word constants
    WD_ALIGN_PARAGRAPH_LEFT = 0
    WD_ALIGN_PARAGRAPH_CENTER = 1
    WD_ALIGN_PARAGRAPH_RIGHT = 2
    WD_ALIGN_PARAGRAPH_JUSTIFY = 3

    WD_LINE_SPACING_SINGLE = 0
    WD_LINE_SPACING_1_5 = 1
    WD_LINE_SPACING_DOUBLE = 2

    WD_UNDERLINE_NONE = 0
    WD_UNDERLINE_SINGLE = 1
    WD_UNDERLINE_DOUBLE = 3

    WD_COLOR_BLACK = 0
    WD_COLOR_BLUE = 16711680
    WD_COLOR_RED = 255

    WD_SAVE_FORMAT_PDF = 17
    WD_SAVE_FORMAT_DOCX = 16
    WD_SAVE_FORMAT_DOC = 0

    # Excel constants
    XL_HALIGN_LEFT = -4131
    XL_HALIGN_CENTER = -4108
    XL_HALIGN_RIGHT = -4152

    XL_VALIGN_TOP = -4160
    XL_VALIGN_CENTER = -4108
    XL_VALIGN_BOTTOM = -4107

    XL_LINE_STYLE_NONE = -4142
    XL_LINE_STYLE_CONTINUOUS = 1

    XL_CHART_COLUMN = -4100
    XL_CHART_BAR = -4099
    XL_CHART_LINE = -4101
    XL_CHART_PIE = -4102

    XL_FILE_FORMAT_PDF = 57
    XL_FILE_FORMAT_XLSX = 51
    XL_FILE_FORMAT_CSV = 6

    # PowerPoint constants
    PP_SLIDE_LAYOUT_TITLE = 1
    PP_SLIDE_LAYOUT_TEXT = 2
    PP_SLIDE_LAYOUT_BLANK = 12

    PP_SAVE_AS_PDF = 32
    PP_SAVE_AS_PPTX = 24

    PP_EFFECT_FADE = 1
    PP_EFFECT_FLY = 2

    @classmethod
    def get_word_alignment(cls, alignment: str) -> int:
        """Get Word alignment constant from string.

        Args:
            alignment: Alignment name (left, center, right, justify)

        Returns:
            Word alignment constant
        """
        alignment_map = {
            "left": cls.WD_ALIGN_PARAGRAPH_LEFT,
            "center": cls.WD_ALIGN_PARAGRAPH_CENTER,
            "right": cls.WD_ALIGN_PARAGRAPH_RIGHT,
            "justify": cls.WD_ALIGN_PARAGRAPH_JUSTIFY,
        }
        return alignment_map.get(alignment.lower(), cls.WD_ALIGN_PARAGRAPH_LEFT)

    @classmethod
    def get_excel_halignment(cls, alignment: str) -> int:
        """Get Excel horizontal alignment constant from string.

        Args:
            alignment: Alignment name (left, center, right)

        Returns:
            Excel alignment constant
        """
        alignment_map = {
            "left": cls.XL_HALIGN_LEFT,
            "center": cls.XL_HALIGN_CENTER,
            "right": cls.XL_HALIGN_RIGHT,
        }
        return alignment_map.get(alignment.lower(), cls.XL_HALIGN_LEFT)

    @classmethod
    def get_excel_valignment(cls, alignment: str) -> int:
        """Get Excel vertical alignment constant from string.

        Args:
            alignment: Alignment name (top, center, bottom)

        Returns:
            Excel alignment constant
        """
        alignment_map = {
            "top": cls.XL_VALIGN_TOP,
            "center": cls.XL_VALIGN_CENTER,
            "bottom": cls.XL_VALIGN_BOTTOM,
        }
        return alignment_map.get(alignment.lower(), cls.XL_VALIGN_CENTER)

    @classmethod
    def get_chart_type(cls, chart_type: str) -> int:
        """Get Excel chart type constant from string.

        Args:
            chart_type: Chart type name

        Returns:
            Excel chart type constant
        """
        chart_map = {
            "column": cls.XL_CHART_COLUMN,
            "bar": cls.XL_CHART_BAR,
            "line": cls.XL_CHART_LINE,
            "pie": cls.XL_CHART_PIE,
        }
        return chart_map.get(chart_type.lower(), cls.XL_CHART_COLUMN)


def rgb_to_office_color(r: int, g: int, b: int) -> int:
    """Convert RGB values to Office color integer.

    Office uses BGR format internally (Blue-Green-Red).

    Args:
        r: Red component (0-255)
        g: Green component (0-255)
        b: Blue component (0-255)

    Returns:
        Office color integer
    """
    return b * 65536 + g * 256 + r


def office_color_to_rgb(color: int) -> tuple[int, int, int]:
    """Convert Office color integer to RGB tuple.

    Args:
        color: Office color integer

    Returns:
        Tuple of (r, g, b) values
    """
    r = color % 256
    g = (color // 256) % 256
    b = (color // 65536) % 256
    return (r, g, b)
