"""Pytest configuration and shared fixtures for Office Automation tests."""

from pathlib import Path
from typing import Any
from unittest.mock import MagicMock, Mock

import pytest


@pytest.fixture
def mock_com_app() -> MagicMock:
    """Create a mock COM application object."""
    app = MagicMock()
    app.Visible = False
    app.DisplayAlerts = False
    return app


@pytest.fixture
def mock_word_app(mock_com_app: MagicMock) -> MagicMock:
    """Create a mock Word application."""
    app = mock_com_app
    app.Documents = MagicMock()
    app.Documents.Count = 0
    return app


@pytest.fixture
def mock_excel_app(mock_com_app: MagicMock) -> MagicMock:
    """Create a mock Excel application."""
    app = mock_com_app
    app.Workbooks = MagicMock()
    app.Workbooks.Count = 0
    return app


@pytest.fixture
def mock_powerpoint_app(mock_com_app: MagicMock) -> MagicMock:
    """Create a mock PowerPoint application."""
    app = mock_com_app
    app.Presentations = MagicMock()
    app.Presentations.Count = 0
    return app


@pytest.fixture
def mock_word_document() -> MagicMock:
    """Create a mock Word document."""
    doc = MagicMock()
    doc.Name = "test_document.docx"
    doc.FullName = str(Path("/tmp/test_document.docx"))
    doc.Content = MagicMock()
    doc.Paragraphs = MagicMock()
    doc.Tables = MagicMock()
    doc.Sections = MagicMock()
    doc.Range = MagicMock()
    return doc


@pytest.fixture
def mock_excel_workbook() -> MagicMock:
    """Create a mock Excel workbook."""
    wb = MagicMock()
    wb.Name = "test_workbook.xlsx"
    wb.FullName = str(Path("/tmp/test_workbook.xlsx"))

    # Mock worksheets
    ws = MagicMock()
    ws.Name = "Sheet1"
    ws.Range = MagicMock(return_value=MagicMock())
    ws.Cells = MagicMock(return_value=MagicMock())
    ws.UsedRange = MagicMock()

    wb.Worksheets = MagicMock()
    wb.Worksheets.return_value = ws
    wb.Worksheets.Count = 1
    wb.ActiveSheet = ws

    return wb


@pytest.fixture
def mock_powerpoint_presentation() -> MagicMock:
    """Create a mock PowerPoint presentation."""
    pres = MagicMock()
    pres.Name = "test_presentation.pptx"
    pres.FullName = str(Path("/tmp/test_presentation.pptx"))

    # Mock slides
    slide = MagicMock()
    slide.SlideIndex = 1
    slide.Shapes = MagicMock()

    pres.Slides = MagicMock()
    pres.Slides.Count = 0
    pres.Slides.Add = MagicMock(return_value=slide)

    return pres


@pytest.fixture
def mock_pythoncom(mocker: Any) -> None:
    """Mock pythoncom module."""
    mock_module = Mock()
    mock_module.CoInitialize = Mock()
    mock_module.CoUninitialize = Mock()
    mocker.patch("pythoncom.CoInitialize", mock_module.CoInitialize)
    mocker.patch("pythoncom.CoUninitialize", mock_module.CoUninitialize)


@pytest.fixture
def mock_win32com(mocker: Any, mock_com_app: MagicMock) -> MagicMock:
    """Mock win32com.client module."""
    mock_client = Mock()
    mock_client.Dispatch = Mock(return_value=mock_com_app)
    mocker.patch("win32com.client.Dispatch", mock_client.Dispatch)
    return mock_client


@pytest.fixture
def temp_test_file(tmp_path: Path) -> Path:
    """Create a temporary test file path."""
    return tmp_path / "test_file.txt"


@pytest.fixture
def sample_docx_path(tmp_path: Path) -> Path:
    """Create a sample .docx file path."""
    return tmp_path / "sample.docx"


@pytest.fixture
def sample_xlsx_path(tmp_path: Path) -> Path:
    """Create a sample .xlsx file path."""
    return tmp_path / "sample.xlsx"


@pytest.fixture
def sample_pptx_path(tmp_path: Path) -> Path:
    """Create a sample .pptx file path."""
    return tmp_path / "sample.pptx"


@pytest.fixture
def sample_image_path(tmp_path: Path) -> Path:
    """Create a sample image file path."""
    img_path = tmp_path / "sample.png"
    # Create a minimal valid PNG file
    img_path.write_bytes(
        b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01"
        b"\x08\x06\x00\x00\x00\x1f\x15\xc4\x89\x00\x00\x00\nIDATx\x9cc\x00\x01"
        b"\x00\x00\x05\x00\x01\r\n-\xb4\x00\x00\x00\x00IEND\xaeB`\x82"
    )
    return img_path
