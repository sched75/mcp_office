"""Unit tests for MCP server."""

from typing import Any
from unittest.mock import MagicMock, Mock, patch

import pytest

# Import after patching to avoid COM initialization
pytest_plugins = ("pytest_asyncio",)


@pytest.fixture
def mock_services(mocker: Any) -> dict[str, MagicMock]:
    """Mock all Office services."""
    mock_word = MagicMock()
    mock_excel = MagicMock()
    mock_ppt = MagicMock()

    mocker.patch("src.server.get_word_service", return_value=mock_word)
    mocker.patch("src.server.get_excel_service", return_value=mock_excel)
    mocker.patch("src.server.get_powerpoint_service", return_value=mock_ppt)

    return {"word": mock_word, "excel": mock_excel, "powerpoint": mock_ppt}


@pytest.fixture
def mock_mcp_server(mocker: Any) -> MagicMock:
    """Mock MCP Server class."""
    mock_server = MagicMock()
    mocker.patch("src.server.Server", return_value=mock_server)
    return mock_server


class TestServerInitialization:
    """Tests for server initialization."""

    def test_server_creates_successfully(self, mock_mcp_server: MagicMock) -> None:
        """Test server can be instantiated."""
        from src.server import app

        assert app is not None

    def test_get_word_service(self, mock_pythoncom: Any) -> None:
        """Test getting Word service instance."""
        with patch("win32com.client.Dispatch"):
            with patch("src.server._word_service", None):
                from src.server import get_word_service

                service = get_word_service()
                assert service is not None

    def test_get_excel_service(self, mock_pythoncom: Any) -> None:
        """Test getting Excel service instance."""
        with patch("win32com.client.Dispatch"):
            with patch("src.server._excel_service", None):
                from src.server import get_excel_service

                service = get_excel_service()
                assert service is not None

    def test_get_powerpoint_service(self, mock_pythoncom: Any) -> None:
        """Test getting PowerPoint service instance."""
        with patch("win32com.client.Dispatch"):
            with patch("src.server._powerpoint_service", None):
                from src.server import get_powerpoint_service

                service = get_powerpoint_service()
                assert service is not None


class TestToolListing:
    """Tests for MCP tool listing."""

    @pytest.mark.asyncio
    async def test_list_tools_returns_tools(self, mock_services: dict[str, MagicMock]) -> None:
        """Test that list_tools returns available tools."""
        with patch("win32com.client.Dispatch"):
            from src.server import list_tools

            tools = await list_tools()

            assert len(tools) > 0
            assert all(hasattr(tool, "name") for tool in tools)
            assert all(hasattr(tool, "description") for tool in tools)

    @pytest.mark.asyncio
    async def test_list_tools_contains_word_tools(
        self, mock_services: dict[str, MagicMock]
    ) -> None:
        """Test that Word tools are included."""
        with patch("win32com.client.Dispatch"):
            from src.server import list_tools

            tools = await list_tools()
            tool_names = [tool.name for tool in tools]

            assert "word_create_document" in tool_names
            assert "word_add_paragraph" in tool_names

    @pytest.mark.asyncio
    async def test_list_tools_contains_excel_tools(
        self, mock_services: dict[str, MagicMock]
    ) -> None:
        """Test that Excel tools are included."""
        with patch("win32com.client.Dispatch"):
            from src.server import list_tools

            tools = await list_tools()
            tool_names = [tool.name for tool in tools]

            assert "excel_create_workbook" in tool_names
            assert "excel_write_cell" in tool_names

    @pytest.mark.asyncio
    async def test_list_tools_contains_powerpoint_tools(
        self, mock_services: dict[str, MagicMock]
    ) -> None:
        """Test that PowerPoint tools are included."""
        with patch("win32com.client.Dispatch"):
            from src.server import list_tools

            tools = await list_tools()
            tool_names = [tool.name for tool in tools]

            assert "powerpoint_create_presentation" in tool_names
            assert "powerpoint_add_slide" in tool_names


class TestToolExecution:
    """Tests for MCP tool execution."""

    @pytest.mark.asyncio
    async def test_call_word_create_document(self, mock_services: dict[str, MagicMock]) -> None:
        """Test calling word_create_document tool."""
        mock_word = mock_services["word"]
        mock_word.create_document.return_value = {"success": True, "message": "Document created"}

        with patch("win32com.client.Dispatch"):
            from src.server import call_tool

            result = await call_tool("word_create_document", {})

            assert len(result) > 0
            assert result[0].type == "text"
            mock_word.create_document.assert_called_once()

    @pytest.mark.asyncio
    async def test_call_word_add_paragraph(self, mock_services: dict[str, MagicMock]) -> None:
        """Test calling word_add_paragraph tool."""
        mock_word = mock_services["word"]
        mock_word.add_paragraph.return_value = {"success": True, "message": "Paragraph added"}

        with patch("win32com.client.Dispatch"):
            from src.server import call_tool

            result = await call_tool("word_add_paragraph", {"text": "Test paragraph"})

            assert len(result) > 0
            mock_word.add_paragraph.assert_called_once_with("Test paragraph")

    @pytest.mark.asyncio
    async def test_call_excel_write_cell(self, mock_services: dict[str, MagicMock]) -> None:
        """Test calling excel_write_cell tool."""
        mock_excel = mock_services["excel"]
        mock_excel.write_cell.return_value = {"success": True, "message": "Cell written"}

        with patch("win32com.client.Dispatch"):
            from src.server import call_tool

            result = await call_tool(
                "excel_write_cell",
                {"sheet_name": "Sheet1", "cell": "A1", "value": "Test"},
            )

            assert len(result) > 0
            mock_excel.write_cell.assert_called_once()

    @pytest.mark.asyncio
    async def test_call_powerpoint_add_slide(self, mock_services: dict[str, MagicMock]) -> None:
        """Test calling powerpoint_add_slide tool."""
        mock_ppt = mock_services["powerpoint"]
        mock_ppt.add_slide.return_value = {
            "success": True,
            "message": "Slide added",
            "slide_index": 1,
        }

        with patch("win32com.client.Dispatch"):
            from src.server import call_tool

            result = await call_tool("powerpoint_add_slide", {"layout": 2})

            assert len(result) > 0
            mock_ppt.add_slide.assert_called_once()

    @pytest.mark.asyncio
    async def test_call_invalid_tool(self, mock_services: dict[str, MagicMock]) -> None:
        """Test calling an invalid tool name."""
        with patch("win32com.client.Dispatch"):
            from src.server import call_tool

            result = await call_tool("invalid_tool_name", {})

            assert len(result) > 0
            assert "Unknown tool" in result[0].text or "not found" in result[0].text.lower()

    @pytest.mark.asyncio
    async def test_call_tool_with_error(self, mock_services: dict[str, MagicMock]) -> None:
        """Test calling a tool that raises an error."""
        mock_word = mock_services["word"]
        mock_word.create_document.side_effect = Exception("Test error")

        with patch("win32com.client.Dispatch"):
            from src.server import call_tool

            result = await call_tool("word_create_document", {})

            assert len(result) > 0
            # Should handle error gracefully


class TestToolParameters:
    """Tests for tool parameter handling."""

    @pytest.mark.asyncio
    async def test_word_insert_table_with_params(
        self, mock_services: dict[str, MagicMock]
    ) -> None:
        """Test Word tool with multiple parameters."""
        mock_word = mock_services["word"]
        mock_word.insert_table.return_value = {"success": True, "message": "Table inserted"}

        with patch("win32com.client.Dispatch"):
            from src.server import call_tool

            result = await call_tool("word_insert_table", {"rows": 3, "cols": 4})

            assert len(result) > 0
            mock_word.insert_table.assert_called_once_with(3, 4)

    @pytest.mark.asyncio
    async def test_excel_write_formula_with_params(
        self, mock_services: dict[str, MagicMock]
    ) -> None:
        """Test Excel tool with string parameters."""
        mock_excel = mock_services["excel"]
        mock_excel.write_formula.return_value = {"success": True, "message": "Formula written"}

        with patch("win32com.client.Dispatch"):
            from src.server import call_tool

            result = await call_tool(
                "excel_write_formula",
                {"sheet_name": "Sheet1", "cell": "A1", "formula": "=SUM(B1:B10)"},
            )

            assert len(result) > 0
            mock_excel.write_formula.assert_called_once()

    @pytest.mark.asyncio
    async def test_powerpoint_insert_image_with_params(
        self, mock_services: dict[str, MagicMock]
    ) -> None:
        """Test PowerPoint tool with file path parameter."""
        mock_ppt = mock_services["powerpoint"]
        mock_ppt.insert_image.return_value = {"success": True, "message": "Image inserted"}

        with patch("win32com.client.Dispatch"):
            from src.server import call_tool

            result = await call_tool(
                "powerpoint_insert_image",
                {
                    "slide_index": 1,
                    "image_path": "/path/to/image.png",
                    "left": 100,
                    "top": 100,
                },
            )

            assert len(result) > 0
            mock_ppt.insert_image.assert_called_once()


class TestServiceSingletons:
    """Tests for service singleton pattern."""

    def test_word_service_singleton(self, mock_pythoncom: Any) -> None:
        """Test Word service uses singleton pattern."""
        with patch("win32com.client.Dispatch"):
            from src.server import get_word_service

            service1 = get_word_service()
            service2 = get_word_service()

            assert service1 is service2

    def test_excel_service_singleton(self, mock_pythoncom: Any) -> None:
        """Test Excel service uses singleton pattern."""
        with patch("win32com.client.Dispatch"):
            from src.server import get_excel_service

            service1 = get_excel_service()
            service2 = get_excel_service()

            assert service1 is service2

    def test_powerpoint_service_singleton(self, mock_pythoncom: Any) -> None:
        """Test PowerPoint service uses singleton pattern."""
        with patch("win32com.client.Dispatch"):
            from src.server import get_powerpoint_service

            service1 = get_powerpoint_service()
            service2 = get_powerpoint_service()

            assert service1 is service2
