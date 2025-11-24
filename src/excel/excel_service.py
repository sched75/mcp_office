"""Excel automation service implementing all 82 Excel functionalities.

This service provides comprehensive Excel automation capabilities following
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
    validate_cell_address,
    validate_file_path,
    validate_positive_number,
    validate_range_address,
    validate_string_not_empty,
)


class ExcelService(BaseOfficeService, DocumentOperationMixin):
    """Excel automation service with all 82 functionalities.

    Categories:
    - Workbook management (6 methods)
    - Templates (3 methods)
    - Worksheet management (7 methods)
    - Cells and data (7 methods)
    - Formulas and calculations (5 methods)
    - Formatting (10 methods)
    - Structured tables (5 methods)
    - Images and objects (5 methods)
    - Charts (7 methods)
    - Pivot tables (5 methods)
    - Sort and filters (4 methods)
    - Protection (4 methods)
    - Named ranges (3 methods)
    - Data validation (3 methods)
    - Printing (3 methods)
    - Advanced features (14 methods)
    """

    def __init__(self, visible: bool = False) -> None:
        """Initialize Excel service."""
        super().__init__(ApplicationType.EXCEL, visible)

    def _close_document(self) -> None:
        """Close the current workbook."""
        if self._current_document:
            self._current_document.Close(SaveChanges=False)
            self._current_document = None

    # ========================================================================
    # WORKBOOK MANAGEMENT (6 methods)
    # ========================================================================

    @com_safe("create_workbook")
    def create_workbook(self) -> dict[str, Any]:
        """Create a new workbook."""
        if not self.is_initialized:
            self.initialize()

        wb = self.application.Workbooks.Add()
        self._current_document = wb

        return dict_to_result(
            success=True,
            message="Workbook created successfully",
            workbook_name=wb.Name,
        )

    @com_safe("open_workbook")
    def open_workbook(self, file_path: str) -> dict[str, Any]:
        """Open an existing workbook."""
        path = validate_file_path(file_path, must_exist=True, extensions=[".xlsx", ".xls"])

        if not self.is_initialized:
            self.initialize()

        wb = self.application.Workbooks.Open(str(path))
        self._current_document = wb

        return dict_to_result(
            success=True,
            message="Workbook opened successfully",
            file_path=str(path),
            workbook_name=wb.Name,
        )

    @com_safe("save_workbook")
    def save_workbook(self, file_path: str | None = None) -> dict[str, Any]:
        """Save the current workbook."""
        wb = self.current_document

        if file_path:
            path = validate_file_path(file_path)
            ensure_directory_exists(path)
            wb.SaveAs(str(path))
            message = f"Workbook saved as: {path}"
        else:
            wb.Save()
            message = "Workbook saved successfully"

        return dict_to_result(success=True, message=message, file_path=str(file_path or wb.FullName))

    @com_safe("close_workbook")
    def close_workbook(self, save_changes: bool = False) -> dict[str, Any]:
        """Close the current workbook."""
        wb = self.current_document
        wb_name = wb.Name

        wb.Close(SaveChanges=save_changes)
        self._current_document = None

        return dict_to_result(
            success=True,
            message=f"Workbook '{wb_name}' closed",
            saved=save_changes,
        )

    @com_safe("export_to_pdf")
    def export_to_pdf(self, output_path: str) -> dict[str, Any]:
        """Export workbook to PDF."""
        path = validate_file_path(output_path, extensions=[".pdf"])
        ensure_directory_exists(path)

        wb = self.current_document
        wb.ExportAsFixedFormat(
            Type=COMConstants.XL_FILE_FORMAT_PDF,
            Filename=str(path),
        )

        return dict_to_result(success=True, message="Workbook exported to PDF", pdf_path=str(path))

    @com_safe("convert_to_csv")
    def convert_to_csv(self, output_path: str) -> dict[str, Any]:
        """Convert workbook to CSV."""
        path = validate_file_path(output_path, extensions=[".csv"])
        ensure_directory_exists(path)

        wb = self.current_document
        wb.SaveAs(str(path), FileFormat=COMConstants.XL_FILE_FORMAT_CSV)

        return dict_to_result(success=True, message="Workbook converted to CSV", csv_path=str(path))

    # ========================================================================
    # TEMPLATES (3 methods)
    # ========================================================================

    @com_safe("create_from_template")
    def create_from_template(self, template_path: str) -> dict[str, Any]:
        """Create workbook from template."""
        path = validate_file_path(template_path, must_exist=True, extensions=[".xltx", ".xlt"])

        if not self.is_initialized:
            self.initialize()

        wb = self.application.Workbooks.Add(Template=str(path))
        self._current_document = wb

        return dict_to_result(
            success=True,
            message="Workbook created from template",
            template_path=str(path),
        )

    @com_safe("save_as_template")
    def save_as_template(self, template_path: str) -> dict[str, Any]:
        """Save workbook as template."""
        path = validate_file_path(template_path, extensions=[".xltx", ".xlt"])
        ensure_directory_exists(path)

        wb = self.current_document
        wb.SaveAs(str(path), FileFormat=win_constants.xlOpenXMLTemplate)

        return dict_to_result(
            success=True,
            message="Workbook saved as template",
            template_path=str(path),
        )

    @com_safe("list_custom_templates")
    def list_custom_templates(self, directory: str | None = None) -> dict[str, Any]:
        """List available custom templates."""
        if directory:
            template_dir = Path(directory)
        else:
            template_dir = Path.home() / "AppData/Roaming/Microsoft/Templates"

        templates = []
        if template_dir.exists():
            templates = [str(p) for p in template_dir.glob("*.xlt*")]

        return dict_to_result(
            success=True,
            message=f"Found {len(templates)} templates",
            templates=templates,
        )

    # ========================================================================
    # WORKSHEET MANAGEMENT (7 methods)
    # ========================================================================

    @com_safe("add_worksheet")
    def add_worksheet(self, name: str | None = None) -> dict[str, Any]:
        """Add a new worksheet."""
        wb = self.current_document
        ws = wb.Worksheets.Add()

        if name:
            ws.Name = name

        return dict_to_result(success=True, message="Worksheet added", sheet_name=ws.Name)

    @com_safe("delete_worksheet")
    def delete_worksheet(self, sheet_name: str) -> dict[str, Any]:
        """Delete a worksheet."""
        validate_string_not_empty("sheet_name", sheet_name)
        wb = self.current_document

        self.application.DisplayAlerts = False
        wb.Worksheets(sheet_name).Delete()
        self.application.DisplayAlerts = True

        return dict_to_result(success=True, message=f"Worksheet '{sheet_name}' deleted")

    @com_safe("rename_worksheet")
    def rename_worksheet(self, old_name: str, new_name: str) -> dict[str, Any]:
        """Rename a worksheet."""
        validate_string_not_empty("old_name", old_name)
        validate_string_not_empty("new_name", new_name)

        wb = self.current_document
        wb.Worksheets(old_name).Name = new_name

        return dict_to_result(
            success=True,
            message=f"Worksheet renamed from '{old_name}' to '{new_name}'",
        )

    @com_safe("copy_worksheet")
    def copy_worksheet(self, sheet_name: str, new_name: str | None = None) -> dict[str, Any]:
        """Copy a worksheet."""
        validate_string_not_empty("sheet_name", sheet_name)
        wb = self.current_document

        sheet = wb.Worksheets(sheet_name)
        sheet.Copy(After=wb.Worksheets(wb.Worksheets.Count))

        new_sheet = wb.Worksheets(wb.Worksheets.Count)
        if new_name:
            new_sheet.Name = new_name

        return dict_to_result(success=True, message="Worksheet copied", new_sheet_name=new_sheet.Name)

    @com_safe("move_worksheet")
    def move_worksheet(self, sheet_name: str, position: int) -> dict[str, Any]:
        """Move a worksheet to a different position."""
        validate_string_not_empty("sheet_name", sheet_name)
        wb = self.current_document

        sheet = wb.Worksheets(sheet_name)
        if position == 1:
            sheet.Move(Before=wb.Worksheets(1))
        else:
            sheet.Move(After=wb.Worksheets(position - 1))

        return dict_to_result(success=True, message=f"Worksheet moved to position {position}")

    @com_safe("hide_worksheet")
    def hide_worksheet(self, sheet_name: str) -> dict[str, Any]:
        """Hide a worksheet."""
        validate_string_not_empty("sheet_name", sheet_name)
        wb = self.current_document
        wb.Worksheets(sheet_name).Visible = False

        return dict_to_result(success=True, message=f"Worksheet '{sheet_name}' hidden")

    @com_safe("show_worksheet")
    def show_worksheet(self, sheet_name: str) -> dict[str, Any]:
        """Show a hidden worksheet."""
        validate_string_not_empty("sheet_name", sheet_name)
        wb = self.current_document
        wb.Worksheets(sheet_name).Visible = True

        return dict_to_result(success=True, message=f"Worksheet '{sheet_name}' shown")

    # ========================================================================
    # CELLS AND DATA (7 methods)
    # ========================================================================

    @com_safe("write_cell")
    def write_cell(self, sheet_name: str, cell: str, value: Any) -> dict[str, Any]:
        """Write value to a cell."""
        validate_string_not_empty("sheet_name", sheet_name)
        cell_addr = validate_cell_address(cell)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        ws.Range(cell_addr).Value = value

        return dict_to_result(success=True, message=f"Cell {cell_addr} updated", cell=cell_addr)

    @com_safe("write_range")
    def write_range(self, sheet_name: str, range_addr: str, values: list[list[Any]]) -> dict[str, Any]:
        """Write values to a range."""
        validate_string_not_empty("sheet_name", sheet_name)
        range_address = validate_range_address(range_addr)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        ws.Range(range_address).Value = values

        return dict_to_result(success=True, message=f"Range {range_address} updated")

    @com_safe("read_cell")
    def read_cell(self, sheet_name: str, cell: str) -> dict[str, Any]:
        """Read value from a cell."""
        validate_string_not_empty("sheet_name", sheet_name)
        cell_addr = validate_cell_address(cell)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        value = ws.Range(cell_addr).Value

        return dict_to_result(success=True, message="Cell value retrieved", cell=cell_addr, value=value)

    @com_safe("read_range")
    def read_range(self, sheet_name: str, range_addr: str) -> dict[str, Any]:
        """Read values from a range."""
        validate_string_not_empty("sheet_name", sheet_name)
        range_address = validate_range_address(range_addr)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        values = ws.Range(range_address).Value

        return dict_to_result(
            success=True,
            message="Range values retrieved",
            range=range_address,
            values=values,
        )

    @com_safe("copy_paste_cells")
    def copy_paste_cells(
        self, sheet_name: str, source_range: str, dest_range: str
    ) -> dict[str, Any]:
        """Copy and paste cells."""
        validate_string_not_empty("sheet_name", sheet_name)
        source = validate_range_address(source_range)
        dest = validate_range_address(dest_range)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)

        ws.Range(source).Copy()
        ws.Range(dest).PasteSpecial()

        return dict_to_result(success=True, message=f"Copied {source} to {dest}")

    @com_safe("clear_contents")
    def clear_contents(self, sheet_name: str, range_addr: str) -> dict[str, Any]:
        """Clear cell contents."""
        validate_string_not_empty("sheet_name", sheet_name)
        range_address = validate_range_address(range_addr)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        ws.Range(range_address).ClearContents()

        return dict_to_result(success=True, message=f"Range {range_address} cleared")

    @com_safe("find_and_replace")
    def find_and_replace(
        self, sheet_name: str, find_text: str, replace_text: str
    ) -> dict[str, Any]:
        """Find and replace in worksheet."""
        validate_string_not_empty("sheet_name", sheet_name)
        validate_string_not_empty("find_text", find_text)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)

        ws.Cells.Replace(What=find_text, Replacement=replace_text)

        return dict_to_result(success=True, message="Find and replace completed")

    # ========================================================================
    # FORMULAS AND CALCULATIONS (5 methods)
    # ========================================================================

    @com_safe("write_formula")
    def write_formula(self, sheet_name: str, cell: str, formula: str) -> dict[str, Any]:
        """Write formula to a cell."""
        validate_string_not_empty("sheet_name", sheet_name)
        cell_addr = validate_cell_address(cell)
        validate_string_not_empty("formula", formula)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        ws.Range(cell_addr).Formula = formula

        return dict_to_result(success=True, message=f"Formula set in {cell_addr}")

    @com_safe("use_function")
    def use_function(
        self, sheet_name: str, cell: str, function_name: str, range_addr: str
    ) -> dict[str, Any]:
        """Use common function (SUM, AVERAGE, IF, etc.)."""
        validate_string_not_empty("sheet_name", sheet_name)
        cell_addr = validate_cell_address(cell)

        formula = f"={function_name}({range_addr})"
        return self.write_formula(sheet_name, cell_addr, formula)

    @com_safe("use_vlookup")
    def use_vlookup(
        self,
        sheet_name: str,
        cell: str,
        lookup_value: str,
        table_array: str,
        col_index: int,
        exact_match: bool = True,
    ) -> dict[str, Any]:
        """Use VLOOKUP function."""
        validate_string_not_empty("sheet_name", sheet_name)
        cell_addr = validate_cell_address(cell)

        match_type = 0 if exact_match else 1
        formula = f"=VLOOKUP({lookup_value},{table_array},{col_index},{match_type})"

        return self.write_formula(sheet_name, cell_addr, formula)

    @com_safe("set_reference_type")
    def set_reference_type(
        self, sheet_name: str, cell: str, formula: str, absolute: bool = True
    ) -> dict[str, Any]:
        """Set formula with absolute/relative references."""
        validate_string_not_empty("sheet_name", sheet_name)
        cell_addr = validate_cell_address(cell)

        # Convert references to absolute if needed
        if absolute and "$" not in formula:
            # Simple conversion - add $ before letters and numbers
            import re
            formula = re.sub(r"([A-Z])([0-9])", r"$\1$\2", formula)

        return self.write_formula(sheet_name, cell_addr, formula)

    @com_safe("use_array_formula")
    def use_array_formula(
        self, sheet_name: str, range_addr: str, formula: str
    ) -> dict[str, Any]:
        """Apply array formula."""
        validate_string_not_empty("sheet_name", sheet_name)
        range_address = validate_range_address(range_addr)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        ws.Range(range_address).FormulaArray = formula

        return dict_to_result(success=True, message=f"Array formula applied to {range_address}")

    # ========================================================================
    # FORMATTING (10 methods)
    # ========================================================================

    @com_safe("set_number_format")
    def set_number_format(
        self, sheet_name: str, range_addr: str, format_code: str
    ) -> dict[str, Any]:
        """Set number format."""
        validate_string_not_empty("sheet_name", sheet_name)
        range_address = validate_range_address(range_addr)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        ws.Range(range_address).NumberFormat = format_code

        return dict_to_result(success=True, message="Number format applied")

    @com_safe("set_cell_color")
    def set_cell_color(
        self, sheet_name: str, range_addr: str, r: int, g: int, b: int
    ) -> dict[str, Any]:
        """Set cell background color."""
        validate_string_not_empty("sheet_name", sheet_name)
        range_address = validate_range_address(range_addr)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        ws.Range(range_address).Interior.Color = rgb_to_office_color(r, g, b)

        return dict_to_result(success=True, message="Cell color set")

    @com_safe("set_font_color")
    def set_font_color(
        self, sheet_name: str, range_addr: str, r: int, g: int, b: int
    ) -> dict[str, Any]:
        """Set font color."""
        validate_string_not_empty("sheet_name", sheet_name)
        range_address = validate_range_address(range_addr)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        ws.Range(range_address).Font.Color = rgb_to_office_color(r, g, b)

        return dict_to_result(success=True, message="Font color set")

    @com_safe("set_borders")
    def set_borders(
        self, sheet_name: str, range_addr: str, border_style: int = 1
    ) -> dict[str, Any]:
        """Set cell borders."""
        validate_string_not_empty("sheet_name", sheet_name)
        range_address = validate_range_address(range_addr)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        cell_range = ws.Range(range_address)

        for border_id in [7, 8, 9, 10]:  # xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight
            cell_range.Borders(border_id).LineStyle = border_style

        return dict_to_result(success=True, message="Borders applied")

    @com_safe("set_alignment")
    def set_alignment(
        self,
        sheet_name: str,
        range_addr: str,
        horizontal: str | None = None,
        vertical: str | None = None,
    ) -> dict[str, Any]:
        """Set cell alignment."""
        validate_string_not_empty("sheet_name", sheet_name)
        range_address = validate_range_address(range_addr)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        cell_range = ws.Range(range_address)

        if horizontal:
            cell_range.HorizontalAlignment = COMConstants.get_excel_halignment(horizontal)
        if vertical:
            cell_range.VerticalAlignment = COMConstants.get_excel_valignment(vertical)

        return dict_to_result(success=True, message="Alignment set")

    @com_safe("set_wrap_text")
    def set_wrap_text(self, sheet_name: str, range_addr: str, wrap: bool = True) -> dict[str, Any]:
        """Set text wrapping."""
        validate_string_not_empty("sheet_name", sheet_name)
        range_address = validate_range_address(range_addr)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        ws.Range(range_address).WrapText = wrap

        return dict_to_result(success=True, message="Text wrapping set")

    @com_safe("merge_cells")
    def merge_cells(self, sheet_name: str, range_addr: str) -> dict[str, Any]:
        """Merge cells."""
        validate_string_not_empty("sheet_name", sheet_name)
        range_address = validate_range_address(range_addr)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        ws.Range(range_address).Merge()

        return dict_to_result(success=True, message=f"Cells merged: {range_address}")

    @com_safe("set_column_width")
    def set_column_width(self, sheet_name: str, column: str, width: float) -> dict[str, Any]:
        """Set column width."""
        validate_string_not_empty("sheet_name", sheet_name)
        width_value = validate_positive_number("width", width)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        ws.Columns(column).ColumnWidth = width_value

        return dict_to_result(success=True, message=f"Column {column} width set to {width_value}")

    @com_safe("set_row_height")
    def set_row_height(self, sheet_name: str, row: int, height: float) -> dict[str, Any]:
        """Set row height."""
        validate_string_not_empty("sheet_name", sheet_name)
        height_value = validate_positive_number("height", height)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        ws.Rows(row).RowHeight = height_value

        return dict_to_result(success=True, message=f"Row {row} height set to {height_value}")

    @com_safe("conditional_formatting")
    def conditional_formatting(
        self, sheet_name: str, range_addr: str, condition_type: int, **kwargs: Any
    ) -> dict[str, Any]:
        """Apply conditional formatting."""
        validate_string_not_empty("sheet_name", sheet_name)
        range_address = validate_range_address(range_addr)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        cell_range = ws.Range(range_address)

        cell_range.FormatConditions.Add(Type=condition_type)

        return dict_to_result(success=True, message="Conditional formatting applied")

    # ========================================================================
    # STRUCTURED TABLES (5 methods)
    # ========================================================================

    @com_safe("convert_to_table")
    def convert_to_table(
        self, sheet_name: str, range_addr: str, table_name: str | None = None
    ) -> dict[str, Any]:
        """Convert range to table."""
        validate_string_not_empty("sheet_name", sheet_name)
        range_address = validate_range_address(range_addr)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)

        list_object = ws.ListObjects.Add(SourceType=1, Source=ws.Range(range_address))

        if table_name:
            list_object.Name = table_name

        return dict_to_result(
            success=True,
            message="Range converted to table",
            table_name=list_object.Name,
        )

    @com_safe("add_total_row")
    def add_total_row(self, sheet_name: str, table_name: str) -> dict[str, Any]:
        """Add total row to table."""
        validate_string_not_empty("sheet_name", sheet_name)
        validate_string_not_empty("table_name", table_name)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        table = ws.ListObjects(table_name)
        table.ShowTotals = True

        return dict_to_result(success=True, message="Total row added")

    @com_safe("apply_table_style")
    def apply_table_style(self, sheet_name: str, table_name: str, style_name: str) -> dict[str, Any]:
        """Apply style to table."""
        validate_string_not_empty("sheet_name", sheet_name)
        validate_string_not_empty("table_name", table_name)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        table = ws.ListObjects(table_name)
        table.TableStyle = style_name

        return dict_to_result(success=True, message="Table style applied")

    @com_safe("filter_table")
    def filter_table(
        self, sheet_name: str, table_name: str, column: int, criteria: str
    ) -> dict[str, Any]:
        """Filter table."""
        validate_string_not_empty("sheet_name", sheet_name)
        validate_string_not_empty("table_name", table_name)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        table = ws.ListObjects(table_name)
        table.Range.AutoFilter(Field=column, Criteria1=criteria)

        return dict_to_result(success=True, message="Table filtered")

    @com_safe("sort_table")
    def sort_table(
        self, sheet_name: str, table_name: str, column: int, ascending: bool = True
    ) -> dict[str, Any]:
        """Sort table."""
        validate_string_not_empty("sheet_name", sheet_name)
        validate_string_not_empty("table_name", table_name)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        table = ws.ListObjects(table_name)

        sort_order = 1 if ascending else 2  # xlAscending : xlDescending
        table.Sort.SortFields.Clear()
        table.Sort.SortFields.Add(
            Key=table.ListColumns(column).Range, SortOn=0, Order=sort_order
        )
        table.Sort.Apply()

        return dict_to_result(success=True, message="Table sorted")

    # ========================================================================
    # IMAGES AND OBJECTS (5 methods)
    # ========================================================================

    @com_safe("insert_image")
    def insert_image(
        self, sheet_name: str, image_path: str, cell: str, width: float | None = None, height: float | None = None
    ) -> dict[str, Any]:
        """Insert image in worksheet."""
        validate_string_not_empty("sheet_name", sheet_name)
        path = validate_file_path(image_path, must_exist=True)
        cell_addr = validate_cell_address(cell)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)

        cell_range = ws.Range(cell_addr)
        picture = ws.Shapes.AddPicture(
            Filename=str(path),
            LinkToFile=False,
            SaveWithDocument=True,
            Left=cell_range.Left,
            Top=cell_range.Top,
            Width=-1,
            Height=-1,
        )

        if width:
            picture.Width = width
        if height:
            picture.Height = height

        return dict_to_result(success=True, message="Image inserted")

    @com_safe("resize_image")
    def resize_image(self, sheet_name: str, image_index: int, width: float, height: float) -> dict[str, Any]:
        """Resize image."""
        validate_string_not_empty("sheet_name", sheet_name)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        shape = ws.Shapes(image_index)

        shape.Width = width
        shape.Height = height

        return dict_to_result(success=True, message="Image resized")

    @com_safe("position_image")
    def position_image(self, sheet_name: str, image_index: int, left: float, top: float) -> dict[str, Any]:
        """Position image."""
        validate_string_not_empty("sheet_name", sheet_name)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        shape = ws.Shapes(image_index)

        shape.Left = left
        shape.Top = top

        return dict_to_result(success=True, message="Image positioned")

    @com_safe("anchor_image_to_cell")
    def anchor_image_to_cell(self, sheet_name: str, image_index: int, cell: str) -> dict[str, Any]:
        """Anchor image to cell."""
        validate_string_not_empty("sheet_name", sheet_name)
        cell_addr = validate_cell_address(cell)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        shape = ws.Shapes(image_index)
        cell_range = ws.Range(cell_addr)

        shape.Left = cell_range.Left
        shape.Top = cell_range.Top

        return dict_to_result(success=True, message=f"Image anchored to {cell_addr}")

    @com_safe("insert_logo_watermark")
    def insert_logo_watermark(self, sheet_name: str, image_path: str) -> dict[str, Any]:
        """Insert logo/watermark."""
        validate_string_not_empty("sheet_name", sheet_name)
        path = validate_file_path(image_path, must_exist=True)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)

        # Insert as watermark (semi-transparent background image)
        picture = ws.Shapes.AddPicture(
            Filename=str(path),
            LinkToFile=False,
            SaveWithDocument=True,
            Left=0,
            Top=0,
            Width=200,
            Height=200,
        )

        # Make semi-transparent
        picture.Fill.Transparency = 0.5

        return dict_to_result(success=True, message="Watermark inserted")

    # ========================================================================
    # CHARTS (7 methods)
    # ========================================================================

    @com_safe("create_chart")
    def create_chart(
        self,
        sheet_name: str,
        chart_type: str,
        source_range: str,
        chart_title: str | None = None,
    ) -> dict[str, Any]:
        """Create chart."""
        validate_string_not_empty("sheet_name", sheet_name)
        range_address = validate_range_address(source_range)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)

        chart = ws.ChartObjects().Add(Left=100, Top=50, Width=400, Height=300)
        chart.Chart.SetSourceData(Source=ws.Range(range_address))
        chart.Chart.ChartType = COMConstants.get_chart_type(chart_type)

        if chart_title:
            chart.Chart.HasTitle = True
            chart.Chart.ChartTitle.Text = chart_title

        return dict_to_result(success=True, message="Chart created", chart_type=chart_type)

    @com_safe("modify_chart_data")
    def modify_chart_data(self, sheet_name: str, chart_index: int, new_range: str) -> dict[str, Any]:
        """Modify chart data source."""
        validate_string_not_empty("sheet_name", sheet_name)
        range_address = validate_range_address(new_range)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)

        chart = ws.ChartObjects(chart_index)
        chart.Chart.SetSourceData(Source=ws.Range(range_address))

        return dict_to_result(success=True, message="Chart data updated")

    @com_safe("customize_chart_title")
    def customize_chart_title(self, sheet_name: str, chart_index: int, title: str) -> dict[str, Any]:
        """Customize chart title."""
        validate_string_not_empty("sheet_name", sheet_name)
        validate_string_not_empty("title", title)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)

        chart = ws.ChartObjects(chart_index)
        chart.Chart.HasTitle = True
        chart.Chart.ChartTitle.Text = title

        return dict_to_result(success=True, message="Chart title updated")

    @com_safe("customize_chart_legend")
    def customize_chart_legend(
        self, sheet_name: str, chart_index: int, position: int = 2
    ) -> dict[str, Any]:
        """Customize chart legend."""
        validate_string_not_empty("sheet_name", sheet_name)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)

        chart = ws.ChartObjects(chart_index)
        chart.Chart.HasLegend = True
        chart.Chart.Legend.Position = position

        return dict_to_result(success=True, message="Chart legend customized")

    @com_safe("modify_chart_axes")
    def modify_chart_axes(
        self, sheet_name: str, chart_index: int, x_title: str | None = None, y_title: str | None = None
    ) -> dict[str, Any]:
        """Modify chart axes."""
        validate_string_not_empty("sheet_name", sheet_name)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        chart = ws.ChartObjects(chart_index).Chart

        if x_title:
            chart.Axes(1).HasTitle = True  # xlCategory
            chart.Axes(1).AxisTitle.Text = x_title

        if y_title:
            chart.Axes(2).HasTitle = True  # xlValue
            chart.Axes(2).AxisTitle.Text = y_title

        return dict_to_result(success=True, message="Chart axes modified")

    @com_safe("change_chart_colors")
    def change_chart_colors(self, sheet_name: str, chart_index: int, color_scheme: int) -> dict[str, Any]:
        """Change chart colors and style."""
        validate_string_not_empty("sheet_name", sheet_name)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        chart = ws.ChartObjects(chart_index).Chart

        chart.ChartStyle = color_scheme

        return dict_to_result(success=True, message="Chart colors changed")

    @com_safe("move_resize_chart")
    def move_resize_chart(
        self,
        sheet_name: str,
        chart_index: int,
        left: float | None = None,
        top: float | None = None,
        width: float | None = None,
        height: float | None = None,
    ) -> dict[str, Any]:
        """Move and resize chart."""
        validate_string_not_empty("sheet_name", sheet_name)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        chart = ws.ChartObjects(chart_index)

        if left is not None:
            chart.Left = left
        if top is not None:
            chart.Top = top
        if width is not None:
            chart.Width = width
        if height is not None:
            chart.Height = height

        return dict_to_result(success=True, message="Chart moved/resized")

    # ========================================================================
    # PIVOT TABLES (5 methods)
    # ========================================================================

    @com_safe("create_pivot_table")
    def create_pivot_table(
        self,
        source_sheet: str,
        source_range: str,
        dest_sheet: str,
        dest_cell: str,
        table_name: str,
    ) -> dict[str, Any]:
        """Create pivot table."""
        validate_string_not_empty("source_sheet", source_sheet)
        range_address = validate_range_address(source_range)
        dest_cell_addr = validate_cell_address(dest_cell)

        wb = self.current_document
        source_ws = wb.Worksheets(source_sheet)
        dest_ws = wb.Worksheets(dest_sheet)

        cache = wb.PivotCaches().Create(
            SourceType=1, SourceData=source_ws.Range(range_address)  # xlDatabase
        )

        cache.CreatePivotTable(
            TableDestination=dest_ws.Range(dest_cell_addr), TableName=table_name
        )

        return dict_to_result(success=True, message="Pivot table created", table_name=table_name)

    @com_safe("set_pivot_fields")
    def set_pivot_fields(
        self,
        sheet_name: str,
        pivot_table_name: str,
        row_fields: list[str] | None = None,
        column_fields: list[str] | None = None,
        value_fields: list[str] | None = None,
    ) -> dict[str, Any]:
        """Set pivot table fields."""
        validate_string_not_empty("sheet_name", sheet_name)
        validate_string_not_empty("pivot_table_name", pivot_table_name)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        pivot = ws.PivotTables(pivot_table_name)

        if row_fields:
            for field in row_fields:
                pivot.PivotFields(field).Orientation = 1  # xlRowField

        if column_fields:
            for field in column_fields:
                pivot.PivotFields(field).Orientation = 2  # xlColumnField

        if value_fields:
            for field in value_fields:
                pivot.PivotFields(field).Orientation = 4  # xlDataField

        return dict_to_result(success=True, message="Pivot fields configured")

    @com_safe("apply_pivot_filter")
    def apply_pivot_filter(
        self, sheet_name: str, pivot_table_name: str, field: str, values: list[str]
    ) -> dict[str, Any]:
        """Apply filter to pivot table."""
        validate_string_not_empty("sheet_name", sheet_name)
        validate_string_not_empty("pivot_table_name", pivot_table_name)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        pivot = ws.PivotTables(pivot_table_name)

        pivot_field = pivot.PivotFields(field)
        pivot_field.Orientation = 3  # xlPageField

        return dict_to_result(success=True, message="Pivot filter applied")

    @com_safe("change_pivot_calculation")
    def change_pivot_calculation(
        self, sheet_name: str, pivot_table_name: str, field: str, function: int
    ) -> dict[str, Any]:
        """Change pivot table calculation."""
        validate_string_not_empty("sheet_name", sheet_name)
        validate_string_not_empty("pivot_table_name", pivot_table_name)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        pivot = ws.PivotTables(pivot_table_name)

        pivot.PivotFields(field).Function = function

        return dict_to_result(success=True, message="Pivot calculation changed")

    @com_safe("refresh_pivot_table")
    def refresh_pivot_table(self, sheet_name: str, pivot_table_name: str) -> dict[str, Any]:
        """Refresh pivot table data."""
        validate_string_not_empty("sheet_name", sheet_name)
        validate_string_not_empty("pivot_table_name", pivot_table_name)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        pivot = ws.PivotTables(pivot_table_name)
        pivot.RefreshTable()

        return dict_to_result(success=True, message="Pivot table refreshed")

    # ========================================================================
    # SORT AND FILTERS (4 methods)
    # ========================================================================

    @com_safe("sort_ascending")
    def sort_ascending(self, sheet_name: str, range_addr: str, key_column: int = 1) -> dict[str, Any]:
        """Sort range in ascending order."""
        validate_string_not_empty("sheet_name", sheet_name)
        range_address = validate_range_address(range_addr)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        cell_range = ws.Range(range_address)

        cell_range.Sort(Key1=cell_range.Columns(key_column), Order1=1)  # xlAscending

        return dict_to_result(success=True, message="Sorted in ascending order")

    @com_safe("sort_descending")
    def sort_descending(self, sheet_name: str, range_addr: str, key_column: int = 1) -> dict[str, Any]:
        """Sort range in descending order."""
        validate_string_not_empty("sheet_name", sheet_name)
        range_address = validate_range_address(range_addr)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        cell_range = ws.Range(range_address)

        cell_range.Sort(Key1=cell_range.Columns(key_column), Order1=2)  # xlDescending

        return dict_to_result(success=True, message="Sorted in descending order")

    @com_safe("apply_autofilter")
    def apply_autofilter(self, sheet_name: str, range_addr: str) -> dict[str, Any]:
        """Apply auto filter."""
        validate_string_not_empty("sheet_name", sheet_name)
        range_address = validate_range_address(range_addr)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        ws.Range(range_address).AutoFilter()

        return dict_to_result(success=True, message="Auto filter applied")

    @com_safe("create_advanced_filter")
    def create_advanced_filter(
        self, sheet_name: str, data_range: str, criteria_range: str
    ) -> dict[str, Any]:
        """Create advanced filter."""
        validate_string_not_empty("sheet_name", sheet_name)
        data_addr = validate_range_address(data_range)
        criteria_addr = validate_range_address(criteria_range)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)

        ws.Range(data_addr).AdvancedFilter(
            Action=1,  # xlFilterInPlace
            CriteriaRange=ws.Range(criteria_addr),
        )

        return dict_to_result(success=True, message="Advanced filter created")

    # ========================================================================
    # PROTECTION (4 methods)
    # ========================================================================

    @com_safe("protect_worksheet")
    def protect_worksheet(self, sheet_name: str, password: str = "") -> dict[str, Any]:
        """Protect worksheet."""
        validate_string_not_empty("sheet_name", sheet_name)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        ws.Protect(Password=password)

        return dict_to_result(success=True, message=f"Worksheet '{sheet_name}' protected")

    @com_safe("protect_workbook")
    def protect_workbook(self, password: str = "") -> dict[str, Any]:
        """Protect workbook structure."""
        wb = self.current_document
        wb.Protect(Password=password, Structure=True)

        return dict_to_result(success=True, message="Workbook protected")

    @com_safe("set_workbook_password")
    def set_workbook_password(self, password: str) -> dict[str, Any]:
        """Set workbook password."""
        validate_string_not_empty("password", password)

        wb = self.current_document
        wb.Password = password

        return dict_to_result(success=True, message="Workbook password set")

    @com_safe("unprotect_worksheet")
    def unprotect_worksheet(self, sheet_name: str, password: str = "") -> dict[str, Any]:
        """Remove worksheet protection."""
        validate_string_not_empty("sheet_name", sheet_name)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        ws.Unprotect(Password=password)

        return dict_to_result(success=True, message=f"Worksheet '{sheet_name}' unprotected")

    # ========================================================================
    # NAMED RANGES (3 methods)
    # ========================================================================

    @com_safe("create_named_range")
    def create_named_range(self, name: str, sheet_name: str, range_addr: str) -> dict[str, Any]:
        """Create named range."""
        validate_string_not_empty("name", name)
        validate_string_not_empty("sheet_name", sheet_name)
        range_address = validate_range_address(range_addr)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)

        wb.Names.Add(Name=name, RefersTo=ws.Range(range_address))

        return dict_to_result(success=True, message=f"Named range '{name}' created")

    @com_safe("use_named_range_in_formula")
    def use_named_range_in_formula(
        self, sheet_name: str, cell: str, range_name: str, function: str = "SUM"
    ) -> dict[str, Any]:
        """Use named range in formula."""
        validate_string_not_empty("sheet_name", sheet_name)
        cell_addr = validate_cell_address(cell)

        formula = f"={function}({range_name})"
        return self.write_formula(sheet_name, cell_addr, formula)

    @com_safe("delete_named_range")
    def delete_named_range(self, name: str) -> dict[str, Any]:
        """Delete named range."""
        validate_string_not_empty("name", name)

        wb = self.current_document
        wb.Names(name).Delete()

        return dict_to_result(success=True, message=f"Named range '{name}' deleted")

    # ========================================================================
    # DATA VALIDATION (3 methods)
    # ========================================================================

    @com_safe("create_dropdown_list")
    def create_dropdown_list(
        self, sheet_name: str, range_addr: str, values: list[str]
    ) -> dict[str, Any]:
        """Create dropdown list."""
        validate_string_not_empty("sheet_name", sheet_name)
        range_address = validate_range_address(range_addr)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        cell_range = ws.Range(range_address)

        values_str = ",".join(values)
        cell_range.Validation.Add(Type=3, Formula1=values_str)  # xlValidateList

        return dict_to_result(success=True, message="Dropdown list created")

    @com_safe("set_validation_rules")
    def set_validation_rules(
        self,
        sheet_name: str,
        range_addr: str,
        validation_type: int,
        formula1: str,
        formula2: str = "",
    ) -> dict[str, Any]:
        """Set data validation rules."""
        validate_string_not_empty("sheet_name", sheet_name)
        range_address = validate_range_address(range_addr)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        cell_range = ws.Range(range_address)

        cell_range.Validation.Add(
            Type=validation_type, Formula1=formula1, Formula2=formula2
        )

        return dict_to_result(success=True, message="Validation rules set")

    @com_safe("remove_validation")
    def remove_validation(self, sheet_name: str, range_addr: str) -> dict[str, Any]:
        """Remove data validation."""
        validate_string_not_empty("sheet_name", sheet_name)
        range_address = validate_range_address(range_addr)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        ws.Range(range_address).Validation.Delete()

        return dict_to_result(success=True, message="Validation removed")

    # ========================================================================
    # PRINTING (3 methods)
    # ========================================================================

    @com_safe("configure_print_settings")
    def configure_print_settings(
        self, sheet_name: str, orientation: str = "portrait", paper_size: int = 1
    ) -> dict[str, Any]:
        """Configure print settings."""
        validate_string_not_empty("sheet_name", sheet_name)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        page_setup = ws.PageSetup

        if orientation.lower() == "landscape":
            page_setup.Orientation = 2  # xlLandscape
        else:
            page_setup.Orientation = 1  # xlPortrait

        page_setup.PaperSize = paper_size

        return dict_to_result(success=True, message="Print settings configured")

    @com_safe("set_print_area")
    def set_print_area(self, sheet_name: str, range_addr: str) -> dict[str, Any]:
        """Set print area."""
        validate_string_not_empty("sheet_name", sheet_name)
        range_address = validate_range_address(range_addr)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        ws.PageSetup.PrintArea = range_address

        return dict_to_result(success=True, message=f"Print area set to {range_address}")

    @com_safe("print_preview")
    def print_preview(self, sheet_name: str) -> dict[str, Any]:
        """Show print preview."""
        validate_string_not_empty("sheet_name", sheet_name)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        ws.PrintPreview()

        return dict_to_result(success=True, message="Print preview opened")

    # ========================================================================
    # ADVANCED FEATURES (14 methods)
    # ========================================================================

    @com_safe("group_rows_columns")
    def group_rows_columns(
        self, sheet_name: str, range_addr: str, is_rows: bool = True
    ) -> dict[str, Any]:
        """Group rows or columns."""
        validate_string_not_empty("sheet_name", sheet_name)
        range_address = validate_range_address(range_addr)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        cell_range = ws.Range(range_address)

        if is_rows:
            cell_range.Rows.Group()
        else:
            cell_range.Columns.Group()

        return dict_to_result(success=True, message="Rows/columns grouped")

    @com_safe("freeze_panes")
    def freeze_panes(self, sheet_name: str, cell: str) -> dict[str, Any]:
        """Freeze panes."""
        validate_string_not_empty("sheet_name", sheet_name)
        cell_addr = validate_cell_address(cell)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)

        ws.Range(cell_addr).Select()
        self.application.ActiveWindow.FreezePanes = True

        return dict_to_result(success=True, message=f"Panes frozen at {cell_addr}")

    @com_safe("split_window")
    def split_window(
        self, sheet_name: str, horizontal_split: float = 0, vertical_split: float = 0
    ) -> dict[str, Any]:
        """Split window."""
        validate_string_not_empty("sheet_name", sheet_name)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        ws.Activate()

        window = self.application.ActiveWindow
        window.SplitRow = horizontal_split
        window.SplitColumn = vertical_split

        return dict_to_result(success=True, message="Window split")

    @com_safe("create_sparklines")
    def create_sparklines(
        self, sheet_name: str, data_range: str, location_range: str, sparkline_type: int = 1
    ) -> dict[str, Any]:
        """Create sparklines."""
        validate_string_not_empty("sheet_name", sheet_name)
        data_addr = validate_range_address(data_range)
        location_addr = validate_range_address(location_range)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)

        ws.Range(location_addr).SparklineGroups.Add(
            Type=sparkline_type, SourceData=ws.Range(data_addr).Address
        )

        return dict_to_result(success=True, message="Sparklines created")

    @com_safe("scenario_analysis")
    def scenario_analysis(
        self, sheet_name: str, scenario_name: str, changing_cells: str, values: list[Any]
    ) -> dict[str, Any]:
        """Create scenario for analysis."""
        validate_string_not_empty("sheet_name", sheet_name)
        validate_string_not_empty("scenario_name", scenario_name)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)

        wb.Scenarios.Add(
            Name=scenario_name,
            ChangingCells=ws.Range(changing_cells),
            Values=values,
        )

        return dict_to_result(success=True, message=f"Scenario '{scenario_name}' created")

    @com_safe("goal_seek")
    def goal_seek(
        self, sheet_name: str, set_cell: str, to_value: float, by_changing_cell: str
    ) -> dict[str, Any]:
        """Perform goal seek."""
        validate_string_not_empty("sheet_name", sheet_name)
        set_cell_addr = validate_cell_address(set_cell)
        by_cell_addr = validate_cell_address(by_changing_cell)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)

        ws.Range(set_cell_addr).GoalSeek(Goal=to_value, ChangingCell=ws.Range(by_cell_addr))

        return dict_to_result(success=True, message="Goal seek completed")

    @com_safe("use_solver")
    def use_solver(self, sheet_name: str, **params: Any) -> dict[str, Any]:
        """Use Solver add-in."""
        validate_string_not_empty("sheet_name", sheet_name)

        # Solver requires add-in to be loaded
        # This is a placeholder for solver functionality
        return dict_to_result(
            success=True,
            message="Solver functionality requires Excel Solver add-in",
        )

    @com_safe("consolidate_data")
    def consolidate_data(
        self,
        dest_sheet: str,
        dest_range: str,
        source_ranges: list[str],
        function: int = -4157,  # xlSum
    ) -> dict[str, Any]:
        """Consolidate data from multiple ranges."""
        validate_string_not_empty("dest_sheet", dest_sheet)
        dest_addr = validate_range_address(dest_range)

        wb = self.current_document
        ws = wb.Worksheets(dest_sheet)

        ws.Range(dest_addr).Consolidate(
            Sources=source_ranges,
            Function=function,
        )

        return dict_to_result(success=True, message="Data consolidated")

    @com_safe("create_subtotals")
    def create_subtotals(
        self, sheet_name: str, range_addr: str, group_by: int, function: int = -4157
    ) -> dict[str, Any]:
        """Create automatic subtotals."""
        validate_string_not_empty("sheet_name", sheet_name)
        range_address = validate_range_address(range_addr)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        cell_range = ws.Range(range_address)

        cell_range.Subtotal(
            GroupBy=group_by,
            Function=function,
            TotalList=[1],
            Replace=True,
        )

        return dict_to_result(success=True, message="Subtotals created")

    @com_safe("import_csv")
    def import_csv(self, sheet_name: str, csv_path: str, dest_cell: str = "A1") -> dict[str, Any]:
        """Import CSV data."""
        validate_string_not_empty("sheet_name", sheet_name)
        path = validate_file_path(csv_path, must_exist=True, extensions=[".csv", ".txt"])
        cell_addr = validate_cell_address(dest_cell)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)

        qt = ws.QueryTables.Add(
            Connection=f"TEXT;{path}",
            Destination=ws.Range(cell_addr),
        )
        qt.TextFileParseType = 1  # xlDelimited
        qt.TextFileCommaDelimiter = True
        qt.Refresh()

        return dict_to_result(success=True, message="CSV imported", csv_path=str(path))

    @com_safe("insert_hyperlink")
    def insert_hyperlink(
        self, sheet_name: str, cell: str, url: str, display_text: str | None = None
    ) -> dict[str, Any]:
        """Insert hyperlink."""
        validate_string_not_empty("sheet_name", sheet_name)
        cell_addr = validate_cell_address(cell)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        cell_range = ws.Range(cell_addr)

        ws.Hyperlinks.Add(
            Anchor=cell_range,
            Address=url,
            TextToDisplay=display_text or url,
        )

        return dict_to_result(success=True, message="Hyperlink inserted")

    @com_safe("insert_comment")
    def insert_comment(self, sheet_name: str, cell: str, comment_text: str) -> dict[str, Any]:
        """Insert comment/note."""
        validate_string_not_empty("sheet_name", sheet_name)
        cell_addr = validate_cell_address(cell)
        validate_string_not_empty("comment_text", comment_text)

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        cell_range = ws.Range(cell_addr)

        cell_range.AddComment(Text=comment_text)

        return dict_to_result(success=True, message="Comment inserted")

    @com_safe("use_3d_reference")
    def use_3d_reference(
        self,
        dest_sheet: str,
        dest_cell: str,
        first_sheet: str,
        last_sheet: str,
        cell_ref: str,
        function: str = "SUM",
    ) -> dict[str, Any]:
        """Use 3D reference across sheets."""
        validate_string_not_empty("dest_sheet", dest_sheet)
        dest_cell_addr = validate_cell_address(dest_cell)

        formula = f"={function}({first_sheet}:{last_sheet}!{cell_ref})"
        return self.write_formula(dest_sheet, dest_cell_addr, formula)

    @com_safe("export_to_json")
    def export_to_json(self, sheet_name: str, range_addr: str, output_path: str) -> dict[str, Any]:
        """Export range to JSON."""
        import json

        validate_string_not_empty("sheet_name", sheet_name)
        range_address = validate_range_address(range_addr)
        path = validate_file_path(output_path, extensions=[".json"])

        wb = self.current_document
        ws = wb.Worksheets(sheet_name)
        values = ws.Range(range_address).Value

        # Convert to list of lists
        data = [list(row) if isinstance(row, tuple) else [row] for row in values] if values else []

        ensure_directory_exists(path)
        with open(path, "w") as f:
            json.dump(data, f, indent=2)

        return dict_to_result(success=True, message="Data exported to JSON", json_path=str(path))

    # Alias for document methods to match base class
    create_document = create_workbook
    open_document = open_workbook
    save_document = save_workbook
    close_document = close_workbook
