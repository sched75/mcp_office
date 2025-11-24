"""Type definitions and enumerations for Office automation.

This module defines common types, enums, and constants used across
the Office automation services following PEP 8 and type hinting best practices.
"""

from enum import Enum, auto
from typing import Literal, TypeAlias

# Type aliases for better readability
CellAddress: TypeAlias = str  # e.g., "A1", "B5"
CellRange: TypeAlias = str  # e.g., "A1:B10"
FilePath: TypeAlias = str
ColorRGB: TypeAlias = tuple[int, int, int]  # RGB values 0-255
WdColor: TypeAlias = int  # Word color constant


class ApplicationType(Enum):
    """Office application types."""

    WORD = "Word.Application"
    EXCEL = "Excel.Application"
    POWERPOINT = "PowerPoint.Application"


class DocumentFormat(Enum):
    """Document format types."""

    # Word formats
    DOCX = auto()
    DOC = auto()
    PDF = auto()
    RTF = auto()
    DOTX = auto()  # Word template
    HTML = auto()
    TXT = auto()

    # Excel formats
    XLSX = auto()
    XLS = auto()
    CSV = auto()
    XLTX = auto()  # Excel template

    # PowerPoint formats
    PPTX = auto()
    PPT = auto()
    POTX = auto()  # PowerPoint template


class TextAlignment(Enum):
    """Text alignment options."""

    LEFT = auto()
    CENTER = auto()
    RIGHT = auto()
    JUSTIFY = auto()


class VerticalAlignment(Enum):
    """Vertical alignment options."""

    TOP = auto()
    MIDDLE = auto()
    BOTTOM = auto()


class FontStyle(Enum):
    """Font styling options."""

    BOLD = auto()
    ITALIC = auto()
    UNDERLINE = auto()
    STRIKETHROUGH = auto()


class BorderStyle(Enum):
    """Border style options."""

    NONE = auto()
    SINGLE = auto()
    DOUBLE = auto()
    DASHED = auto()
    DOTTED = auto()


class ImagePosition(Enum):
    """Image positioning options."""

    INLINE = auto()
    FLOAT = auto()
    ANCHOR = auto()


class ChartType(Enum):
    """Chart type options."""

    COLUMN = auto()
    BAR = auto()
    LINE = auto()
    PIE = auto()
    SCATTER = auto()
    AREA = auto()
    DOUGHNUT = auto()


class ProtectionType(Enum):
    """Document protection types."""

    READ_ONLY = auto()
    COMMENTS_ONLY = auto()
    TRACKED_CHANGES = auto()
    FORMS = auto()


class SlideLayout(Enum):
    """PowerPoint slide layout types."""

    TITLE_SLIDE = auto()
    TITLE_AND_CONTENT = auto()
    SECTION_HEADER = auto()
    TWO_CONTENT = auto()
    COMPARISON = auto()
    TITLE_ONLY = auto()
    BLANK = auto()
    CONTENT_WITH_CAPTION = auto()
    PICTURE_WITH_CAPTION = auto()


class AnimationType(Enum):
    """Animation types."""

    ENTRANCE = auto()
    EXIT = auto()
    EMPHASIS = auto()
    MOTION_PATH = auto()


class TransitionType(Enum):
    """Slide transition types."""

    NONE = auto()
    FADE = auto()
    PUSH = auto()
    WIPE = auto()
    SPLIT = auto()
    REVEAL = auto()
    RANDOM = auto()


# Literal types for specific parameters
WdUnderlineStyle: TypeAlias = Literal[
    "none", "single", "double", "dotted", "dashed", "wave"
]
NumberFormat: TypeAlias = Literal[
    "general", "number", "currency", "accounting", "date", "time", "percentage",
    "fraction", "scientific", "text"
]
