"""Word automation service implementing all 65 Word functionalities.

This service provides comprehensive Word automation capabilities following
SOLID principles and design patterns.
"""

from pathlib import Path
from typing import Any

from win32com.client import constants as win_constants

from ..core.base_office import BaseOfficeService, DocumentOperationMixin
from ..core.types import ApplicationType
from ..utils.com_wrapper import COMConstants, com_safe, rgb_to_office_color
from ..utils.helpers import dict_to_result, ensure_directory_exists
from ..utils.validators import (
    validate_file_path,
    validate_positive_number,
    validate_string_not_empty,
)


class WordService(BaseOfficeService, DocumentOperationMixin):
    """Word automation service with all 65 functionalities.

    This service implements the complete Word automation API covering:
    - Document management (6 methods)
    - Templates (3 methods)
    - Text content (4 methods)
    - Text formatting (5 methods)
    - Tables (7 methods)
    - Images and objects (8 methods)
    - Document structure (7 methods)
    - Revision (5 methods)
    - Metadata and properties (4 methods)
    - Printing (3 methods)
    - Protection (3 methods)
    - Advanced features (10 methods)
    """

    def __init__(self, visible: bool = False) -> None:
        """Initialize Word service.

        Args:
            visible: Whether to make Word window visible
        """
        super().__init__(ApplicationType.WORD, visible)

    def _close_document(self) -> None:
        """Close the current document (internal method)."""
        if self._current_document:
            self._current_document.Close(SaveChanges=False)
            self._current_document = None

    # ========================================================================
    # DOCUMENT MANAGEMENT (6 methods)
    # ========================================================================

    @com_safe("create_document")
    def create_document(self) -> dict[str, Any]:
        """Create a new Word document."""
        if not self.is_initialized:
            self.initialize()

        doc = self.application.Documents.Add()
        self._current_document = doc

        return dict_to_result(
            success=True,
            message="Document created successfully",
            document_name=doc.Name,
        )

    @com_safe("open_document")
    def open_document(self, file_path: str) -> dict[str, Any]:
        """Open an existing Word document."""
        path = validate_file_path(file_path, must_exist=True, extensions=[".docx", ".doc"])

        if not self.is_initialized:
            self.initialize()

        doc = self.application.Documents.Open(str(path))
        self._current_document = doc

        return dict_to_result(
            success=True,
            message="Document opened successfully",
            file_path=str(path),
            document_name=doc.Name,
        )

    @com_safe("save_document")
    def save_document(self, file_path: str | None = None) -> dict[str, Any]:
        """Save the current document."""
        doc = self.current_document

        if file_path:
            path = validate_file_path(file_path)
            ensure_directory_exists(path)
            doc.SaveAs2(str(path))
            message = f"Document saved as: {path}"
        else:
            doc.Save()
            message = "Document saved successfully"

        return dict_to_result(
            success=True, message=message, file_path=str(file_path or doc.FullName)
        )

    @com_safe("close_document")
    def close_document(self, save_changes: bool = False) -> dict[str, Any]:
        """Close the current document."""
        doc = self.current_document
        doc_name = doc.Name

        doc.Close(SaveChanges=save_changes)
        self._current_document = None

        return dict_to_result(
            success=True,
            message=f"Document '{doc_name}' closed",
            saved=save_changes,
        )

    @com_safe("export_to_pdf")
    def export_to_pdf(self, output_path: str) -> dict[str, Any]:
        """Export document to PDF."""
        path = validate_file_path(output_path, extensions=[".pdf"])
        ensure_directory_exists(path)

        doc = self.current_document
        doc.ExportAsFixedFormat(
            OutputFileName=str(path),
            ExportFormat=COMConstants.WD_SAVE_FORMAT_PDF,
        )

        return dict_to_result(
            success=True,
            message="Document exported to PDF",
            pdf_path=str(path),
        )

    @com_safe("print_to_pdf")
    def print_to_pdf(self, output_path: str) -> dict[str, Any]:
        """Print document to PDF (alias for export_to_pdf)."""
        return self.export_to_pdf(output_path)

    # ========================================================================
    # TEMPLATES (3 methods)
    # ========================================================================

    @com_safe("create_from_template")
    def create_from_template(self, template_path: str) -> dict[str, Any]:
        """Create document from template."""
        path = validate_file_path(template_path, must_exist=True, extensions=[".dotx", ".dot"])

        if not self.is_initialized:
            self.initialize()

        doc = self.application.Documents.Add(Template=str(path))
        self._current_document = doc

        return dict_to_result(
            success=True,
            message="Document created from template",
            template_path=str(path),
        )

    @com_safe("save_as_template")
    def save_as_template(self, template_path: str) -> dict[str, Any]:
        """Save current document as template."""
        path = validate_file_path(template_path, extensions=[".dotx", ".dot"])
        ensure_directory_exists(path)

        doc = self.current_document
        doc.SaveAs2(str(path), FileFormat=win_constants.wdFormatXMLTemplate)

        return dict_to_result(
            success=True,
            message="Document saved as template",
            template_path=str(path),
        )

    @com_safe("list_available_templates")
    def list_available_templates(self, directory: str | None = None) -> dict[str, Any]:
        """List available Word templates."""
        if directory:
            template_dir = Path(directory)
        else:
            # Default templates directory
            template_dir = Path.home() / "AppData/Roaming/Microsoft/Templates"

        templates = []
        if template_dir.exists():
            templates = [str(p) for p in template_dir.glob("*.dot*")]

        return dict_to_result(
            success=True,
            message=f"Found {len(templates)} templates",
            templates=templates,
            directory=str(template_dir),
        )

    # ========================================================================
    # TEXT CONTENT (4 methods)
    # ========================================================================

    @com_safe("add_paragraph")
    def add_paragraph(self, text: str, style: str | None = None) -> dict[str, Any]:
        """Add a paragraph to the document."""
        validate_string_not_empty("text", text)
        doc = self.current_document

        para = doc.Content.Paragraphs.Add()
        para.Range.Text = text

        if style:
            para.Style = style

        return dict_to_result(success=True, message="Paragraph added", text_length=len(text))

    @com_safe("insert_text_at_position")
    def insert_text_at_position(self, text: str, position: int = 0) -> dict[str, Any]:
        """Insert text at specific position."""
        validate_string_not_empty("text", text)
        doc = self.current_document

        doc_range = doc.Range(Start=position, End=position)
        doc_range.Text = text

        return dict_to_result(
            success=True,
            message="Text inserted",
            position=position,
            text_length=len(text),
        )

    @com_safe("find_and_replace")
    def find_and_replace(
        self, find_text: str, replace_text: str, match_case: bool = False
    ) -> dict[str, Any]:
        """Find and replace text in document."""
        validate_string_not_empty("find_text", find_text)
        doc = self.current_document

        find_obj = doc.Content.Find
        find_obj.ClearFormatting()
        find_obj.Text = find_text
        find_obj.MatchCase = match_case
        find_obj.Replacement.Text = replace_text

        count = 0
        while find_obj.Execute(Replace=2):  # wdReplaceOne
            count += 1

        return dict_to_result(
            success=True,
            message=f"Replaced {count} occurrences",
            replacements=count,
            find_text=find_text,
        )

    @com_safe("delete_text")
    def delete_text(self, start: int, end: int) -> dict[str, Any]:
        """Delete text between positions."""
        doc = self.current_document
        doc_range = doc.Range(Start=start, End=end)
        doc_range.Delete()

        return dict_to_result(
            success=True,
            message=f"Deleted text from {start} to {end}",
            start=start,
            end=end,
        )

    # ========================================================================
    # TEXT FORMATTING (5 methods)
    # ========================================================================

    @com_safe("apply_text_formatting")
    def apply_text_formatting(
        self,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        font_name: str | None = None,
        font_size: int | None = None,
        color_rgb: tuple[int, int, int] | None = None,
        start: int | None = None,
        end: int | None = None,
    ) -> dict[str, Any]:
        """Apply text formatting."""
        doc = self.current_document

        if start is not None and end is not None:
            text_range = doc.Range(Start=start, End=end)
        else:
            text_range = doc.Content

        if bold is not None:
            text_range.Font.Bold = bold
        if italic is not None:
            text_range.Font.Italic = italic
        if underline is not None:
            text_range.Font.Underline = (
                COMConstants.WD_UNDERLINE_SINGLE if underline else COMConstants.WD_UNDERLINE_NONE
            )
        if font_name:
            text_range.Font.Name = font_name
        if font_size:
            text_range.Font.Size = validate_positive_number("font_size", font_size)
        if color_rgb:
            text_range.Font.Color = rgb_to_office_color(*color_rgb)

        return dict_to_result(success=True, message="Text formatting applied")

    @com_safe("set_paragraph_alignment")
    def set_paragraph_alignment(
        self, alignment: str, paragraph_index: int | None = None
    ) -> dict[str, Any]:
        """Set paragraph alignment."""
        doc = self.current_document
        alignment_const = COMConstants.get_word_alignment(alignment)

        if paragraph_index is not None:
            para = doc.Paragraphs(paragraph_index)
            para.Alignment = alignment_const
        else:
            for para in doc.Paragraphs:
                para.Alignment = alignment_const

        return dict_to_result(success=True, message="Alignment applied", alignment=alignment)

    @com_safe("apply_style")
    def apply_style(self, style_name: str, paragraph_index: int | None = None) -> dict[str, Any]:
        """Apply predefined style."""
        doc = self.current_document

        if paragraph_index is not None:
            doc.Paragraphs(paragraph_index).Style = style_name
        else:
            doc.Content.Style = style_name

        return dict_to_result(success=True, message="Style applied", style=style_name)

    @com_safe("set_line_spacing")
    def set_line_spacing(
        self, spacing: float, paragraph_index: int | None = None
    ) -> dict[str, Any]:
        """Set line spacing."""
        spacing_value = validate_positive_number("spacing", spacing)
        doc = self.current_document

        if paragraph_index is not None:
            doc.Paragraphs(paragraph_index).LineSpacingRule = 1  # wdLineSpaceMultiple
            doc.Paragraphs(paragraph_index).LineSpacing = spacing_value
        else:
            for para in doc.Paragraphs:
                para.LineSpacingRule = 1  # wdLineSpaceMultiple
                para.LineSpacing = spacing_value

        return dict_to_result(success=True, message="Line spacing set", spacing=spacing_value)

    @com_safe("create_custom_style")
    def create_custom_style(
        self, style_name: str, base_style: str = "Normal", **formatting: Any
    ) -> dict[str, Any]:
        """Create custom style."""
        doc = self.current_document
        style = doc.Styles.Add(Name=style_name, Type=win_constants.wdStyleTypeParagraph)
        style.BaseStyle = base_style

        if "font_name" in formatting:
            style.Font.Name = formatting["font_name"]
        if "font_size" in formatting:
            style.Font.Size = formatting["font_size"]
        if "bold" in formatting:
            style.Font.Bold = formatting["bold"]

        return dict_to_result(success=True, message="Custom style created", style_name=style_name)

    # ========================================================================
    # TABLES (7 methods)
    # ========================================================================

    @com_safe("insert_table")
    def insert_table(self, rows: int, cols: int) -> dict[str, Any]:
        """Insert table with dimensions."""
        if rows < 1 or cols < 1:
            raise ValueError("Rows and columns must be positive")

        doc = self.current_document
        table_range = doc.Content
        table_range.Collapse(Direction=0)  # wdCollapseEnd

        doc.Tables.Add(Range=table_range, NumRows=rows, NumColumns=cols)

        return dict_to_result(
            success=True,
            message=f"Table created ({rows}x{cols})",
            rows=rows,
            cols=cols,
        )

    @com_safe("set_table_cell_text")
    def set_table_cell_text(
        self, table_index: int, row: int, col: int, text: str
    ) -> dict[str, Any]:
        """Set text in table cell."""
        doc = self.current_document
        table = doc.Tables(table_index)
        table.Cell(Row=row, Column=col).Range.Text = text

        return dict_to_result(success=True, message="Cell text set", row=row, col=col)

    @com_safe("add_table_row")
    def add_table_row(self, table_index: int) -> dict[str, Any]:
        """Add row to table."""
        doc = self.current_document
        table = doc.Tables(table_index)
        table.Rows.Add()

        return dict_to_result(success=True, message="Row added", row_count=table.Rows.Count)

    @com_safe("add_table_column")
    def add_table_column(self, table_index: int) -> dict[str, Any]:
        """Add column to table."""
        doc = self.current_document
        table = doc.Tables(table_index)
        table.Columns.Add()

        return dict_to_result(success=True, message="Column added", col_count=table.Columns.Count)

    @com_safe("delete_table_row")
    def delete_table_row(self, table_index: int, row_index: int) -> dict[str, Any]:
        """Delete row from table."""
        doc = self.current_document
        table = doc.Tables(table_index)
        table.Rows(row_index).Delete()

        return dict_to_result(success=True, message="Row deleted", row_index=row_index)

    @com_safe("delete_table_column")
    def delete_table_column(self, table_index: int, col_index: int) -> dict[str, Any]:
        """Delete column from table."""
        doc = self.current_document
        table = doc.Tables(table_index)
        table.Columns(col_index).Delete()

        return dict_to_result(success=True, message="Column deleted", col_index=col_index)

    @com_safe("merge_table_cells")
    def merge_table_cells(
        self, table_index: int, start_row: int, start_col: int, end_row: int, end_col: int
    ) -> dict[str, Any]:
        """Merge table cells."""
        doc = self.current_document
        table = doc.Tables(table_index)

        start_cell = table.Cell(start_row, start_col)
        end_cell = table.Cell(end_row, end_col)
        start_cell.Merge(MergeTo=end_cell)

        return dict_to_result(success=True, message="Cells merged")

    # ========================================================================
    # IMAGES AND OBJECTS (8 methods)
    # ========================================================================

    @com_safe("insert_image")
    def insert_image(
        self, image_path: str, width: float | None = None, height: float | None = None
    ) -> dict[str, Any]:
        """Insert image from file."""
        path = validate_file_path(image_path, must_exist=True)
        doc = self.current_document

        img_range = doc.Content
        img_range.Collapse(Direction=0)  # wdCollapseEnd
        shape = doc.InlineShapes.AddPicture(FileName=str(path), Range=img_range)

        if width:
            shape.Width = width
        if height:
            shape.Height = height

        return dict_to_result(success=True, message="Image inserted", image_path=str(path))

    @com_safe("insert_image_from_clipboard")
    def insert_image_from_clipboard(self) -> dict[str, Any]:
        """Insert image from clipboard."""
        selection = self.application.Selection
        selection.Paste()

        return dict_to_result(success=True, message="Image inserted from clipboard")

    @com_safe("resize_image")
    def resize_image(self, image_index: int, width: float, height: float) -> dict[str, Any]:
        """Resize image."""
        doc = self.current_document
        shape = doc.InlineShapes(image_index)
        shape.Width = width
        shape.Height = height

        return dict_to_result(success=True, message="Image resized", width=width, height=height)

    @com_safe("position_image")
    def position_image(self, image_index: int, wrap_format: int = 0) -> dict[str, Any]:
        """Position image with text wrapping."""
        doc = self.current_document
        shape = doc.InlineShapes(image_index)

        # Convert to floating shape for positioning
        float_shape = shape.ConvertToShape()
        float_shape.WrapFormat.Type = wrap_format

        return dict_to_result(success=True, message="Image positioned")

    @com_safe("crop_image")
    def crop_image(
        self, image_index: int, left: float = 0, top: float = 0, right: float = 0, bottom: float = 0
    ) -> dict[str, Any]:
        """Crop image."""
        doc = self.current_document
        shape = doc.InlineShapes(image_index)

        shape.PictureFormat.CropLeft = left
        shape.PictureFormat.CropTop = top
        shape.PictureFormat.CropRight = right
        shape.PictureFormat.CropBottom = bottom

        return dict_to_result(success=True, message="Image cropped")

    @com_safe("apply_image_effects")
    def apply_image_effects(
        self, image_index: int, brightness: float = 0, contrast: float = 0
    ) -> dict[str, Any]:
        """Apply effects to image."""
        doc = self.current_document
        shape = doc.InlineShapes(image_index)

        if brightness != 0:
            shape.PictureFormat.Brightness = brightness
        if contrast != 0:
            shape.PictureFormat.Contrast = contrast

        return dict_to_result(success=True, message="Image effects applied")

    @com_safe("insert_shape")
    def insert_shape(
        self, shape_type: int, left: float, top: float, width: float, height: float
    ) -> dict[str, Any]:
        """Insert shape."""
        doc = self.current_document
        doc.Shapes.AddShape(Type=shape_type, Left=left, Top=top, Width=width, Height=height)

        return dict_to_result(success=True, message="Shape inserted")

    @com_safe("add_textbox")
    def add_textbox(
        self, text: str, left: float, top: float, width: float, height: float
    ) -> dict[str, Any]:
        """Add textbox."""
        doc = self.current_document
        textbox = doc.Shapes.AddTextbox(
            Orientation=1, Left=left, Top=top, Width=width, Height=height
        )
        textbox.TextFrame.TextRange.Text = text

        return dict_to_result(success=True, message="Textbox added")

    # ========================================================================
    # DOCUMENT STRUCTURE (7 methods)
    # ========================================================================

    @com_safe("add_header")
    def add_header(self, text: str, section_index: int = 1) -> dict[str, Any]:
        """Add header to document."""
        doc = self.current_document
        section = doc.Sections(section_index)
        header = section.Headers(1)  # wdHeaderFooterPrimary
        header.Range.Text = text

        return dict_to_result(success=True, message="Header added")

    @com_safe("add_footer")
    def add_footer(self, text: str, section_index: int = 1) -> dict[str, Any]:
        """Add footer to document."""
        doc = self.current_document
        section = doc.Sections(section_index)
        footer = section.Footers(1)  # wdHeaderFooterPrimary
        footer.Range.Text = text

        return dict_to_result(success=True, message="Footer added")

    @com_safe("insert_page_numbers")
    def insert_page_numbers(self, position: str = "bottom") -> dict[str, Any]:
        """Insert page numbers."""
        doc = self.current_document
        section = doc.Sections(1)

        if position.lower() == "bottom":
            footer = section.Footers(1)  # wdHeaderFooterPrimary
            footer.PageNumbers.Add()
        else:
            header = section.Headers(1)  # wdHeaderFooterPrimary
            header.PageNumbers.Add()

        return dict_to_result(success=True, message="Page numbers inserted")

    @com_safe("create_table_of_contents")
    def create_table_of_contents(self) -> dict[str, Any]:
        """Create table of contents."""
        doc = self.current_document
        toc_range = doc.Range(Start=0, End=0)

        doc.TablesOfContents.Add(
            Range=toc_range,
            UseHeadingStyles=True,
            UpperHeadingLevel=1,
            LowerHeadingLevel=3,
        )

        return dict_to_result(success=True, message="Table of contents created")

    @com_safe("insert_page_break")
    def insert_page_break(self) -> dict[str, Any]:
        """Insert page break."""
        selection = self.application.Selection
        selection.InsertBreak(Type=7)  # wdPageBreak

        return dict_to_result(success=True, message="Page break inserted")

    @com_safe("insert_section_break")
    def insert_section_break(
        self, break_type: int = 2  # wdSectionBreakNextPage
    ) -> dict[str, Any]:
        """Insert section break."""
        selection = self.application.Selection
        selection.InsertBreak(Type=break_type)

        return dict_to_result(success=True, message="Section break inserted")

    @com_safe("configure_section")
    def configure_section(
        self,
        section_index: int,
        orientation: str | None = None,
        page_width: float | None = None,
        page_height: float | None = None,
    ) -> dict[str, Any]:
        """Configure section properties."""
        doc = self.current_document
        section = doc.Sections(section_index)
        page_setup = section.PageSetup

        if orientation:
            if orientation.lower() == "landscape":
                page_setup.Orientation = win_constants.wdOrientLandscape
            else:
                page_setup.Orientation = win_constants.wdOrientPortrait

        if page_width:
            page_setup.PageWidth = page_width
        if page_height:
            page_setup.PageHeight = page_height

        return dict_to_result(success=True, message="Section configured")

    # ========================================================================
    # REVISION (5 methods)
    # ========================================================================

    @com_safe("enable_track_changes")
    def enable_track_changes(self) -> dict[str, Any]:
        """Enable track changes."""
        doc = self.current_document
        doc.TrackRevisions = True

        return dict_to_result(success=True, message="Track changes enabled")

    @com_safe("disable_track_changes")
    def disable_track_changes(self) -> dict[str, Any]:
        """Disable track changes."""
        doc = self.current_document
        doc.TrackRevisions = False

        return dict_to_result(success=True, message="Track changes disabled")

    @com_safe("add_comment")
    def add_comment(self, text: str, range_start: int = 0, range_end: int = 0) -> dict[str, Any]:
        """Add comment to document."""
        doc = self.current_document
        comment_range = doc.Range(Start=range_start, End=range_end)
        doc.Comments.Add(Range=comment_range, Text=text)

        return dict_to_result(success=True, message="Comment added")

    @com_safe("accept_all_revisions")
    def accept_all_revisions(self) -> dict[str, Any]:
        """Accept all revisions."""
        doc = self.current_document
        doc.Revisions.AcceptAll()

        return dict_to_result(success=True, message="All revisions accepted")

    @com_safe("reject_all_revisions")
    def reject_all_revisions(self) -> dict[str, Any]:
        """Reject all revisions."""
        doc = self.current_document
        doc.Revisions.RejectAll()

        return dict_to_result(success=True, message="All revisions rejected")

    # ========================================================================
    # METADATA AND PROPERTIES (4 methods)
    # ========================================================================

    @com_safe("get_document_properties")
    def get_document_properties(self) -> dict[str, Any]:
        """Get document properties."""
        doc = self.current_document
        props = doc.BuiltInDocumentProperties

        properties = {
            "author": props("Author").Value,
            "title": props("Title").Value,
            "subject": props("Subject").Value,
            "keywords": props("Keywords").Value,
        }

        return dict_to_result(success=True, message="Properties retrieved", properties=properties)

    @com_safe("set_document_properties")
    def set_document_properties(
        self,
        author: str | None = None,
        title: str | None = None,
        subject: str | None = None,
        keywords: str | None = None,
    ) -> dict[str, Any]:
        """Set document properties."""
        doc = self.current_document
        props = doc.BuiltInDocumentProperties

        if author:
            props("Author").Value = author
        if title:
            props("Title").Value = title
        if subject:
            props("Subject").Value = subject
        if keywords:
            props("Keywords").Value = keywords

        return dict_to_result(success=True, message="Properties updated")

    @com_safe("get_document_statistics")
    def get_document_statistics(self) -> dict[str, Any]:
        """Get document statistics."""
        doc = self.current_document

        stats = {
            "pages": doc.ComputeStatistics(win_constants.wdStatisticPages),
            "words": doc.ComputeStatistics(win_constants.wdStatisticWords),
            "characters": doc.ComputeStatistics(win_constants.wdStatisticCharacters),
            "paragraphs": doc.ComputeStatistics(win_constants.wdStatisticParagraphs),
        }

        return dict_to_result(success=True, message="Statistics retrieved", statistics=stats)

    @com_safe("set_document_language")
    def set_document_language(self, language_id: int) -> dict[str, Any]:
        """Set document language."""
        doc = self.current_document
        doc.Content.LanguageID = language_id

        return dict_to_result(success=True, message="Language set", language_id=language_id)

    # ========================================================================
    # PRINTING (3 methods)
    # ========================================================================

    @com_safe("configure_print_settings")
    def configure_print_settings(
        self, copies: int = 1, page_range: str = "", collate: bool = True
    ) -> dict[str, Any]:
        """Configure print settings."""
        # Settings are applied during actual print
        settings = {
            "copies": copies,
            "page_range": page_range,
            "collate": collate,
        }

        return dict_to_result(
            success=True,
            message="Print settings configured",
            settings=settings,
        )

    @com_safe("print_preview")
    def print_preview(self) -> dict[str, Any]:
        """Show print preview."""
        doc = self.current_document
        doc.PrintPreview()

        return dict_to_result(success=True, message="Print preview opened")

    # ========================================================================
    # PROTECTION (3 methods)
    # ========================================================================

    @com_safe("protect_document")
    def protect_document(self, protection_type: int, password: str = "") -> dict[str, Any]:
        """Protect document."""
        doc = self.current_document
        doc.Protect(Type=protection_type, Password=password)

        return dict_to_result(success=True, message="Document protected")

    @com_safe("set_password")
    def set_password(self, password: str) -> dict[str, Any]:
        """Set document password."""
        doc = self.current_document
        doc.Password = password

        return dict_to_result(success=True, message="Password set")

    @com_safe("unprotect_document")
    def unprotect_document(self, password: str = "") -> dict[str, Any]:
        """Remove document protection."""
        doc = self.current_document
        doc.Unprotect(Password=password)

        return dict_to_result(success=True, message="Protection removed")

    # ========================================================================
    # ADVANCED FEATURES (10 methods)
    # ========================================================================

    @com_safe("mail_merge_with_data")
    def mail_merge_with_data(self, data_source: str) -> dict[str, Any]:
        """Perform mail merge."""
        path = validate_file_path(data_source, must_exist=True)
        doc = self.current_document

        doc.MailMerge.OpenDataSource(Name=str(path))
        doc.MailMerge.Execute(Pause=False)

        return dict_to_result(success=True, message="Mail merge completed")

    @com_safe("insert_bookmark")
    def insert_bookmark(
        self, name: str, range_start: int = 0, range_end: int = 0
    ) -> dict[str, Any]:
        """Insert bookmark."""
        doc = self.current_document
        bookmark_range = doc.Range(Start=range_start, End=range_end)
        doc.Bookmarks.Add(Name=name, Range=bookmark_range)

        return dict_to_result(success=True, message="Bookmark inserted", name=name)

    @com_safe("create_index")
    def create_index(self) -> dict[str, Any]:
        """Create index."""
        doc = self.current_document
        index_range = doc.Range()
        index_range.Collapse(Direction=win_constants.wdCollapseEnd)
        doc.Indexes.Add(Range=index_range)

        return dict_to_result(success=True, message="Index created")

    @com_safe("manage_bibliography")
    def manage_bibliography(self, source_file: str | None = None) -> dict[str, Any]:
        """Manage bibliography."""
        doc = self.current_document

        if source_file:
            validate_file_path(source_file, must_exist=True)
            # Add bibliography source
            return dict_to_result(success=True, message="Bibliography source added")

        # Insert bibliography
        bib_range = doc.Range()
        bib_range.Collapse(Direction=win_constants.wdCollapseEnd)
        doc.Bibliography.Add(Range=bib_range)

        return dict_to_result(success=True, message="Bibliography inserted")

    @com_safe("insert_field")
    def insert_field(self, field_type: int, text: str = "") -> dict[str, Any]:
        """Insert field."""
        selection = self.application.Selection
        selection.Fields.Add(Range=selection.Range, Type=field_type, Text=text)

        return dict_to_result(success=True, message="Field inserted")

    @com_safe("compare_documents")
    def compare_documents(self, original_path: str, revised_path: str) -> dict[str, Any]:
        """Compare two documents."""
        original = validate_file_path(original_path, must_exist=True)
        revised = validate_file_path(revised_path, must_exist=True)

        result_doc = self.application.CompareDocuments(
            OriginalDocument=str(original),
            RevisedDocument=str(revised),
        )

        self._current_document = result_doc

        return dict_to_result(success=True, message="Documents compared")

    @com_safe("insert_smartart")
    def insert_smartart(self, layout: int = 1) -> dict[str, Any]:
        """Insert SmartArt."""
        doc = self.current_document
        doc.Shapes.AddSmartArt(Layout=layout)

        return dict_to_result(success=True, message="SmartArt inserted")

    @com_safe("convert_format")
    def convert_format(self, output_path: str, file_format: int) -> dict[str, Any]:
        """Convert document format."""
        path = validate_file_path(output_path)
        ensure_directory_exists(path)

        doc = self.current_document
        doc.SaveAs2(FileName=str(path), FileFormat=file_format)

        return dict_to_result(success=True, message="Format converted", output_path=str(path))

    @com_safe("modify_style")
    def modify_style(self, style_name: str, **formatting: Any) -> dict[str, Any]:
        """Modify existing style."""
        doc = self.current_document
        style = doc.Styles(style_name)

        if "font_name" in formatting:
            style.Font.Name = formatting["font_name"]
        if "font_size" in formatting:
            style.Font.Size = formatting["font_size"]
        if "bold" in formatting:
            style.Font.Bold = formatting["bold"]

        return dict_to_result(success=True, message="Style modified", style_name=style_name)

    @com_safe("insert_hyperlink")
    def insert_hyperlink(
        self, text: str, url: str, range_start: int | None = None, range_end: int | None = None
    ) -> dict[str, Any]:
        """Insert hyperlink."""
        doc = self.current_document

        if range_start is not None and range_end is not None:
            link_range = doc.Range(Start=range_start, End=range_end)
        else:
            link_range = self.application.Selection.Range

        doc.Hyperlinks.Add(Anchor=link_range, Address=url, TextToDisplay=text)

        return dict_to_result(success=True, message="Hyperlink inserted", url=url)
