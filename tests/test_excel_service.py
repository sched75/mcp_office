"""Unit tests for ExcelService with mocked COM objects."""

from pathlib import Path
from typing import Any
from unittest.mock import MagicMock, patch

import pytest

from src.core.exceptions import DocumentNotOpenError
from src.excel.excel_service import ExcelService


@pytest.fixture
def excel_service(mock_pythoncom: Any, mocker: Any) -> ExcelService:
    """Create an ExcelService instance with mocked COM."""
    with patch("win32com.client.Dispatch") as mock_dispatch:
        mock_app = MagicMock()
        mock_dispatch.return_value = mock_app
        service = ExcelService()
        service.initialize()
        return service


@pytest.fixture
def excel_service_with_wb(
    excel_service: ExcelService, mock_excel_workbook: MagicMock
) -> ExcelService:
    """Create an ExcelService with an open workbook."""
    excel_service._current_document = mock_excel_workbook
    return excel_service


class TestExcelServiceInitialization:
    """Tests for ExcelService initialization."""

    def test_service_creates_successfully(self, mock_pythoncom: Any) -> None:
        """Test ExcelService can be instantiated."""
        with patch("win32com.client.Dispatch"):
            service = ExcelService()
            assert service is not None
            assert not service.is_initialized

    def test_service_initializes(self, excel_service: ExcelService) -> None:
        """Test ExcelService initializes COM application."""
        assert excel_service.is_initialized
        assert excel_service.application is not None

    def test_service_cleanup(self, excel_service: ExcelService) -> None:
        """Test ExcelService cleanup."""
        excel_service.cleanup()
        # Should not raise exception


class TestWorkbookCreation:
    """Tests for workbook creation methods."""

    def test_create_workbook(self, excel_service: ExcelService) -> None:
        """Test creating a new Excel workbook."""
        mock_wb = MagicMock()
        excel_service.application.Workbooks.Add.return_value = mock_wb

        result = excel_service.create_workbook()

        assert result["success"] is True
        excel_service.application.Workbooks.Add.assert_called_once()

    def test_open_workbook(self, excel_service: ExcelService, sample_xlsx_path: Path) -> None:
        """Test opening an existing Excel workbook."""
        sample_xlsx_path.touch()
        mock_wb = MagicMock()
        mock_wb.FullName = str(sample_xlsx_path)
        excel_service.application.Workbooks.Open.return_value = mock_wb

        result = excel_service.open_workbook(str(sample_xlsx_path))

        assert result["success"] is True
        excel_service.application.Workbooks.Open.assert_called_once()

    def test_create_from_template(self, excel_service: ExcelService, tmp_path: Path) -> None:
        """Test creating workbook from template."""
        template_path = tmp_path / "template.xltx"
        template_path.touch()

        mock_wb = MagicMock()
        excel_service.application.Workbooks.Add.return_value = mock_wb

        result = excel_service.create_from_template(str(template_path))

        assert result["success"] is True


class TestWorkbookOperations:
    """Tests for workbook operations."""

    def test_save_workbook(self, excel_service_with_wb: ExcelService) -> None:
        """Test saving a workbook."""
        result = excel_service_with_wb.save_workbook()

        assert result["success"] is True
        excel_service_with_wb.current_document.Save.assert_called_once()

    def test_save_workbook_no_doc_open(self, excel_service: ExcelService) -> None:
        """Test saving when no workbook is open."""
        with pytest.raises(DocumentNotOpenError):
            excel_service.save_workbook()

    def test_close_workbook(self, excel_service_with_wb: ExcelService) -> None:
        """Test closing a workbook."""
        result = excel_service_with_wb.close_workbook()

        assert result["success"] is True
        excel_service_with_wb.current_document.Close.assert_called_once()

    def test_export_to_pdf(self, excel_service_with_wb: ExcelService, tmp_path: Path) -> None:
        """Test exporting workbook to PDF."""
        pdf_path = tmp_path / "output.pdf"

        result = excel_service_with_wb.export_to_pdf(str(pdf_path))

        assert result["success"] is True
        excel_service_with_wb.current_document.ExportAsFixedFormat.assert_called_once()


class TestWorksheetOperations:
    """Tests for worksheet operations."""

    def test_add_worksheet(self, excel_service_with_wb: ExcelService) -> None:
        """Test adding a worksheet."""
        mock_ws = MagicMock()
        mock_ws.Name = "NewSheet"
        excel_service_with_wb.current_document.Worksheets.Add.return_value = mock_ws

        result = excel_service_with_wb.add_worksheet("NewSheet")

        assert result["success"] is True
        excel_service_with_wb.current_document.Worksheets.Add.assert_called_once()

    def test_delete_worksheet(self, excel_service_with_wb: ExcelService) -> None:
        """Test deleting a worksheet."""
        mock_ws = MagicMock()
        excel_service_with_wb.current_document.Worksheets.return_value = mock_ws

        result = excel_service_with_wb.delete_worksheet("Sheet1")

        assert result["success"] is True
        mock_ws.Delete.assert_called_once()

    def test_rename_worksheet(self, excel_service_with_wb: ExcelService) -> None:
        """Test renaming a worksheet."""
        mock_ws = MagicMock()
        excel_service_with_wb.current_document.Worksheets.return_value = mock_ws

        result = excel_service_with_wb.rename_worksheet("Sheet1", "NewName")

        assert result["success"] is True
        assert mock_ws.Name == "NewName"

    def test_hide_worksheet(self, excel_service_with_wb: ExcelService) -> None:
        """Test hiding a worksheet."""
        mock_ws = MagicMock()
        excel_service_with_wb.current_document.Worksheets.return_value = mock_ws

        result = excel_service_with_wb.hide_worksheet("Sheet1")

        assert result["success"] is True


class TestCellOperations:
    """Tests for cell operations."""

    def test_write_cell(self, excel_service_with_wb: ExcelService) -> None:
        """Test writing to a cell."""
        mock_ws = MagicMock()
        mock_range = MagicMock()
        mock_ws.Range.return_value = mock_range
        excel_service_with_wb.current_document.Worksheets.return_value = mock_ws

        result = excel_service_with_wb.write_cell("Sheet1", "A1", "Test value")

        assert result["success"] is True
        assert mock_range.Value == "Test value"

    def test_read_cell(self, excel_service_with_wb: ExcelService) -> None:
        """Test reading from a cell."""
        mock_ws = MagicMock()
        mock_range = MagicMock()
        mock_range.Value = "Cell value"
        mock_ws.Range.return_value = mock_range
        excel_service_with_wb.current_document.Worksheets.return_value = mock_ws

        result = excel_service_with_wb.read_cell("Sheet1", "A1")

        assert result["success"] is True
        assert result["value"] == "Cell value"

    def test_write_range(self, excel_service_with_wb: ExcelService) -> None:
        """Test writing to a range."""
        mock_ws = MagicMock()
        mock_range = MagicMock()
        mock_ws.Range.return_value = mock_range
        excel_service_with_wb.current_document.Worksheets.return_value = mock_ws

        data = [[1, 2, 3], [4, 5, 6]]
        result = excel_service_with_wb.write_range("Sheet1", "A1:C2", data)

        assert result["success"] is True

    def test_read_range(self, excel_service_with_wb: ExcelService) -> None:
        """Test reading from a range."""
        mock_ws = MagicMock()
        mock_range = MagicMock()
        mock_range.Value = ((1, 2, 3), (4, 5, 6))
        mock_ws.Range.return_value = mock_range
        excel_service_with_wb.current_document.Worksheets.return_value = mock_ws

        result = excel_service_with_wb.read_range("Sheet1", "A1:C2")

        assert result["success"] is True
        assert "values" in result


class TestFormulaOperations:
    """Tests for formula operations."""

    def test_write_formula(self, excel_service_with_wb: ExcelService) -> None:
        """Test writing a formula."""
        mock_ws = MagicMock()
        mock_range = MagicMock()
        mock_ws.Range.return_value = mock_range
        excel_service_with_wb.current_document.Worksheets.return_value = mock_ws

        result = excel_service_with_wb.write_formula("Sheet1", "A1", "=SUM(B1:B10)")

        assert result["success"] is True
        assert mock_range.Formula == "=SUM(B1:B10)"

    def test_use_function(self, excel_service_with_wb: ExcelService) -> None:
        """Test using an Excel function."""
        mock_ws = MagicMock()
        mock_range = MagicMock()
        mock_ws.Range.return_value = mock_range
        excel_service_with_wb.current_document.Worksheets.return_value = mock_ws

        result = excel_service_with_wb.use_function("Sheet1", "A1", "SUM", "B1:B10")

        assert result["success"] is True

    def test_use_vlookup(self, excel_service_with_wb: ExcelService) -> None:
        """Test using VLOOKUP."""
        mock_ws = MagicMock()
        mock_range = MagicMock()
        mock_ws.Range.return_value = mock_range
        excel_service_with_wb.current_document.Worksheets.return_value = mock_ws

        result = excel_service_with_wb.use_vlookup("Sheet1", "D1", "A1", "B1:C10", 2)

        assert result["success"] is True


class TestFormattingOperations:
    """Tests for formatting operations."""

    def test_set_number_format(self, excel_service_with_wb: ExcelService) -> None:
        """Test setting number format."""
        mock_ws = MagicMock()
        mock_range = MagicMock()
        mock_ws.Range.return_value = mock_range
        excel_service_with_wb.current_document.Worksheets.return_value = mock_ws

        result = excel_service_with_wb.set_number_format("Sheet1", "A1:A10", "0.00")

        assert result["success"] is True
        assert mock_range.NumberFormat == "0.00"

    def test_set_cell_color(self, excel_service_with_wb: ExcelService) -> None:
        """Test setting cell background color."""
        mock_ws = MagicMock()
        mock_range = MagicMock()
        mock_ws.Range.return_value = mock_range
        excel_service_with_wb.current_document.Worksheets.return_value = mock_ws

        result = excel_service_with_wb.set_cell_color("Sheet1", "A1", 255, 0, 0)

        assert result["success"] is True

    def test_set_alignment(self, excel_service_with_wb: ExcelService) -> None:
        """Test setting cell alignment."""
        mock_ws = MagicMock()
        mock_range = MagicMock()
        mock_ws.Range.return_value = mock_range
        excel_service_with_wb.current_document.Worksheets.return_value = mock_ws

        result = excel_service_with_wb.set_alignment("Sheet1", "A1:B10", "center", "middle")

        assert result["success"] is True

    def test_merge_cells(self, excel_service_with_wb: ExcelService) -> None:
        """Test merging cells."""
        mock_ws = MagicMock()
        mock_range = MagicMock()
        mock_ws.Range.return_value = mock_range
        excel_service_with_wb.current_document.Worksheets.return_value = mock_ws

        result = excel_service_with_wb.merge_cells("Sheet1", "A1:B2")

        assert result["success"] is True
        mock_range.Merge.assert_called_once()


class TestChartOperations:
    """Tests for chart operations."""

    def test_create_chart(self, excel_service_with_wb: ExcelService) -> None:
        """Test creating a chart."""
        mock_ws = MagicMock()
        mock_chart_objects = MagicMock()
        mock_chart_obj = MagicMock()
        mock_chart = MagicMock()
        mock_chart_obj.Chart = mock_chart
        mock_chart_objects.Add.return_value = mock_chart_obj
        mock_ws.ChartObjects.return_value = mock_chart_objects
        mock_ws.Range.return_value = MagicMock()
        excel_service_with_wb.current_document.Worksheets.return_value = mock_ws

        result = excel_service_with_wb.create_chart("Sheet1", "column", "A1:B10", "Sales Chart")

        assert result["success"] is True
        mock_chart_objects.Add.assert_called_once()

    def test_modify_chart_data(self, excel_service_with_wb: ExcelService) -> None:
        """Test modifying chart data source."""
        mock_ws = MagicMock()
        mock_chart_obj = MagicMock()
        mock_chart = MagicMock()
        mock_chart_obj.Chart = mock_chart
        mock_ws.ChartObjects.return_value = mock_chart_obj
        mock_ws.Range.return_value = MagicMock()
        excel_service_with_wb.current_document.Worksheets.return_value = mock_ws

        result = excel_service_with_wb.modify_chart_data("Sheet1", 1, "A1:C10")

        assert result["success"] is True


class TestTableOperations:
    """Tests for table operations."""

    def test_convert_to_table(self, excel_service_with_wb: ExcelService) -> None:
        """Test converting range to table."""
        mock_ws = MagicMock()
        mock_range = MagicMock()
        mock_table = MagicMock()
        mock_ws.Range.return_value = mock_range
        mock_ws.ListObjects.Add.return_value = mock_table
        excel_service_with_wb.current_document.Worksheets.return_value = mock_ws

        result = excel_service_with_wb.convert_to_table("Sheet1", "A1:C10", "MyTable")

        assert result["success"] is True
        mock_ws.ListObjects.Add.assert_called_once()

    def test_sort_table(self, excel_service_with_wb: ExcelService) -> None:
        """Test sorting a table."""
        mock_ws = MagicMock()
        mock_table = MagicMock()
        mock_ws.ListObjects.return_value = mock_table
        excel_service_with_wb.current_document.Worksheets.return_value = mock_ws

        result = excel_service_with_wb.sort_table("Sheet1", "MyTable", "Column1")

        assert result["success"] is True


class TestPivotTableOperations:
    """Tests for pivot table operations."""

    def test_create_pivot_table(self, excel_service_with_wb: ExcelService) -> None:
        """Test creating a pivot table."""
        mock_ws_source = MagicMock()
        mock_ws_dest = MagicMock()
        mock_range = MagicMock()
        mock_ws_source.Range.return_value = mock_range

        def get_worksheet(name: str) -> MagicMock:
            return mock_ws_source if name == "Source" else mock_ws_dest

        excel_service_with_wb.current_document.Worksheets.side_effect = get_worksheet

        mock_pivot_cache = MagicMock()
        mock_pivot_table = MagicMock()
        excel_service_with_wb.current_document.PivotCaches.return_value.Create.return_value = (
            mock_pivot_cache
        )
        mock_pivot_cache.CreatePivotTable.return_value = mock_pivot_table

        result = excel_service_with_wb.create_pivot_table(
            "Source", "A1:C10", "Dest", "A1", "MyPivot"
        )

        assert result["success"] is True


class TestProtectionOperations:
    """Tests for protection operations."""

    def test_protect_worksheet(self, excel_service_with_wb: ExcelService) -> None:
        """Test protecting a worksheet."""
        mock_ws = MagicMock()
        excel_service_with_wb.current_document.Worksheets.return_value = mock_ws

        result = excel_service_with_wb.protect_worksheet("Sheet1", "password")

        assert result["success"] is True
        mock_ws.Protect.assert_called_once()

    def test_protect_workbook(self, excel_service_with_wb: ExcelService) -> None:
        """Test protecting a workbook."""
        result = excel_service_with_wb.protect_workbook("password")

        assert result["success"] is True
        excel_service_with_wb.current_document.Protect.assert_called_once()

    def test_set_workbook_password(self, excel_service_with_wb: ExcelService) -> None:
        """Test setting workbook password."""
        result = excel_service_with_wb.set_workbook_password("mypassword")

        assert result["success"] is True


class TestAdvancedFeatures:
    """Tests for advanced Excel features."""

    def test_create_named_range(self, excel_service_with_wb: ExcelService) -> None:
        """Test creating a named range."""
        mock_ws = MagicMock()
        mock_range = MagicMock()
        mock_ws.Range.return_value = mock_range
        excel_service_with_wb.current_document.Worksheets.return_value = mock_ws

        result = excel_service_with_wb.create_named_range("Sheet1", "A1:B10", "MyRange")

        assert result["success"] is True

    def test_goal_seek(self, excel_service_with_wb: ExcelService) -> None:
        """Test goal seek operation."""
        mock_ws = MagicMock()
        mock_set_cell = MagicMock()
        mock_to_value_cell = MagicMock()
        mock_ws.Range.side_effect = [mock_set_cell, mock_to_value_cell]
        mock_to_value_cell.GoalSeek.return_value = True
        excel_service_with_wb.current_document.Worksheets.return_value = mock_ws

        result = excel_service_with_wb.goal_seek("Sheet1", "B1", 100, "A1")

        assert result["success"] is True
        mock_to_value_cell.GoalSeek.assert_called_once()
