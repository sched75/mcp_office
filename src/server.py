"""MCP Server for Office Automation.

This server exposes Word, Excel, and PowerPoint automation capabilities
through the Model Context Protocol (MCP).
"""

import asyncio
import logging
from typing import Any

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import TextContent, Tool

from .excel.excel_service import ExcelService
from .powerpoint.powerpoint_service import PowerPointService
from .word.word_service import WordService

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize services
word_service: WordService | None = None
excel_service: ExcelService | None = None
powerpoint_service: PowerPointService | None = None


def get_word_service() -> WordService:
    """Get or create Word service instance."""
    global word_service
    if word_service is None:
        word_service = WordService()
        word_service.initialize()
    return word_service


def get_excel_service() -> ExcelService:
    """Get or create Excel service instance."""
    global excel_service
    if excel_service is None:
        excel_service = ExcelService()
        excel_service.initialize()
    return excel_service


def get_powerpoint_service() -> PowerPointService:
    """Get or create PowerPoint service instance."""
    global powerpoint_service
    if powerpoint_service is None:
        powerpoint_service = PowerPointService()
        powerpoint_service.initialize()
    return powerpoint_service


# Create MCP server
app = Server("office-automation")


# ============================================================================
# WORD TOOLS
# ============================================================================

@app.list_tools()
async def list_tools() -> list[Tool]:
    """List all available tools."""
    return [
        # Word Document Management
        Tool(
            name="word_create_document",
            description="Create a new Word document",
            inputSchema={"type": "object", "properties": {}},
        ),
        Tool(
            name="word_open_document",
            description="Open an existing Word document",
            inputSchema={
                "type": "object",
                "properties": {"file_path": {"type": "string"}},
                "required": ["file_path"],
            },
        ),
        Tool(
            name="word_save_document",
            description="Save the current Word document",
            inputSchema={
                "type": "object",
                "properties": {"file_path": {"type": "string"}},
            },
        ),
        Tool(
            name="word_add_paragraph",
            description="Add a paragraph to Word document",
            inputSchema={
                "type": "object",
                "properties": {
                    "text": {"type": "string"},
                    "style": {"type": "string"},
                },
                "required": ["text"],
            },
        ),
        Tool(
            name="word_export_to_pdf",
            description="Export Word document to PDF",
            inputSchema={
                "type": "object",
                "properties": {"output_path": {"type": "string"}},
                "required": ["output_path"],
            },
        ),
        Tool(
            name="word_insert_table",
            description="Insert a table in Word document",
            inputSchema={
                "type": "object",
                "properties": {
                    "rows": {"type": "integer"},
                    "cols": {"type": "integer"},
                },
                "required": ["rows", "cols"],
            },
        ),
        Tool(
            name="word_insert_image",
            description="Insert an image in Word document",
            inputSchema={
                "type": "object",
                "properties": {
                    "image_path": {"type": "string"},
                    "width": {"type": "number"},
                    "height": {"type": "number"},
                },
                "required": ["image_path"],
            },
        ),
        # Excel Workbook Management
        Tool(
            name="excel_create_workbook",
            description="Create a new Excel workbook",
            inputSchema={"type": "object", "properties": {}},
        ),
        Tool(
            name="excel_open_workbook",
            description="Open an existing Excel workbook",
            inputSchema={
                "type": "object",
                "properties": {"file_path": {"type": "string"}},
                "required": ["file_path"],
            },
        ),
        Tool(
            name="excel_save_workbook",
            description="Save the current Excel workbook",
            inputSchema={
                "type": "object",
                "properties": {"file_path": {"type": "string"}},
            },
        ),
        Tool(
            name="excel_add_worksheet",
            description="Add a new worksheet to Excel workbook",
            inputSchema={
                "type": "object",
                "properties": {"name": {"type": "string"}},
            },
        ),
        Tool(
            name="excel_write_cell",
            description="Write value to Excel cell",
            inputSchema={
                "type": "object",
                "properties": {
                    "sheet_name": {"type": "string"},
                    "cell": {"type": "string"},
                    "value": {},
                },
                "required": ["sheet_name", "cell", "value"],
            },
        ),
        Tool(
            name="excel_read_cell",
            description="Read value from Excel cell",
            inputSchema={
                "type": "object",
                "properties": {
                    "sheet_name": {"type": "string"},
                    "cell": {"type": "string"},
                },
                "required": ["sheet_name", "cell"],
            },
        ),
        Tool(
            name="excel_write_formula",
            description="Write formula to Excel cell",
            inputSchema={
                "type": "object",
                "properties": {
                    "sheet_name": {"type": "string"},
                    "cell": {"type": "string"},
                    "formula": {"type": "string"},
                },
                "required": ["sheet_name", "cell", "formula"],
            },
        ),
        Tool(
            name="excel_create_chart",
            description="Create a chart in Excel",
            inputSchema={
                "type": "object",
                "properties": {
                    "sheet_name": {"type": "string"},
                    "chart_type": {"type": "string"},
                    "source_range": {"type": "string"},
                    "chart_title": {"type": "string"},
                },
                "required": ["sheet_name", "chart_type", "source_range"],
            },
        ),
        Tool(
            name="excel_export_to_pdf",
            description="Export Excel workbook to PDF",
            inputSchema={
                "type": "object",
                "properties": {"output_path": {"type": "string"}},
                "required": ["output_path"],
            },
        ),
        # PowerPoint Presentation Management
        Tool(
            name="powerpoint_create_presentation",
            description="Create a new PowerPoint presentation",
            inputSchema={"type": "object", "properties": {}},
        ),
        Tool(
            name="powerpoint_open_presentation",
            description="Open an existing PowerPoint presentation",
            inputSchema={
                "type": "object",
                "properties": {"file_path": {"type": "string"}},
                "required": ["file_path"],
            },
        ),
        Tool(
            name="powerpoint_save_presentation",
            description="Save the current PowerPoint presentation",
            inputSchema={
                "type": "object",
                "properties": {"file_path": {"type": "string"}},
            },
        ),
        Tool(
            name="powerpoint_add_slide",
            description="Add a new slide to PowerPoint presentation",
            inputSchema={
                "type": "object",
                "properties": {"layout": {"type": "integer"}},
            },
        ),
        Tool(
            name="powerpoint_modify_title",
            description="Modify slide title in PowerPoint",
            inputSchema={
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "title_text": {"type": "string"},
                },
                "required": ["slide_index", "title_text"],
            },
        ),
        Tool(
            name="powerpoint_insert_image",
            description="Insert an image in PowerPoint slide",
            inputSchema={
                "type": "object",
                "properties": {
                    "slide_index": {"type": "integer"},
                    "image_path": {"type": "string"},
                    "left": {"type": "number"},
                    "top": {"type": "number"},
                    "width": {"type": "number"},
                    "height": {"type": "number"},
                },
                "required": ["slide_index", "image_path", "left", "top"],
            },
        ),
        Tool(
            name="powerpoint_export_to_pdf",
            description="Export PowerPoint presentation to PDF",
            inputSchema={
                "type": "object",
                "properties": {"output_path": {"type": "string"}},
                "required": ["output_path"],
            },
        ),
    ]


@app.call_tool()
async def call_tool(name: str, arguments: dict[str, Any]) -> list[TextContent]:
    """Handle tool calls."""
    try:
        # Word tools
        if name == "word_create_document":
            service = get_word_service()
            result = service.create_document()
            return [TextContent(type="text", text=str(result))]

        if name == "word_open_document":
            service = get_word_service()
            result = service.open_document(arguments["file_path"])
            return [TextContent(type="text", text=str(result))]

        if name == "word_save_document":
            service = get_word_service()
            result = service.save_document(arguments.get("file_path"))
            return [TextContent(type="text", text=str(result))]

        if name == "word_add_paragraph":
            service = get_word_service()
            result = service.add_paragraph(
                arguments["text"], arguments.get("style")
            )
            return [TextContent(type="text", text=str(result))]

        if name == "word_export_to_pdf":
            service = get_word_service()
            result = service.export_to_pdf(arguments["output_path"])
            return [TextContent(type="text", text=str(result))]

        if name == "word_insert_table":
            service = get_word_service()
            result = service.insert_table(arguments["rows"], arguments["cols"])
            return [TextContent(type="text", text=str(result))]

        if name == "word_insert_image":
            service = get_word_service()
            result = service.insert_image(
                arguments["image_path"],
                arguments.get("width"),
                arguments.get("height"),
            )
            return [TextContent(type="text", text=str(result))]

        # Excel tools
        if name == "excel_create_workbook":
            service = get_excel_service()
            result = service.create_workbook()
            return [TextContent(type="text", text=str(result))]

        if name == "excel_open_workbook":
            service = get_excel_service()
            result = service.open_workbook(arguments["file_path"])
            return [TextContent(type="text", text=str(result))]

        if name == "excel_save_workbook":
            service = get_excel_service()
            result = service.save_workbook(arguments.get("file_path"))
            return [TextContent(type="text", text=str(result))]

        if name == "excel_add_worksheet":
            service = get_excel_service()
            result = service.add_worksheet(arguments.get("name"))
            return [TextContent(type="text", text=str(result))]

        if name == "excel_write_cell":
            service = get_excel_service()
            result = service.write_cell(
                arguments["sheet_name"], arguments["cell"], arguments["value"]
            )
            return [TextContent(type="text", text=str(result))]

        if name == "excel_read_cell":
            service = get_excel_service()
            result = service.read_cell(arguments["sheet_name"], arguments["cell"])
            return [TextContent(type="text", text=str(result))]

        if name == "excel_write_formula":
            service = get_excel_service()
            result = service.write_formula(
                arguments["sheet_name"], arguments["cell"], arguments["formula"]
            )
            return [TextContent(type="text", text=str(result))]

        if name == "excel_create_chart":
            service = get_excel_service()
            result = service.create_chart(
                arguments["sheet_name"],
                arguments["chart_type"],
                arguments["source_range"],
                arguments.get("chart_title"),
            )
            return [TextContent(type="text", text=str(result))]

        if name == "excel_export_to_pdf":
            service = get_excel_service()
            result = service.export_to_pdf(arguments["output_path"])
            return [TextContent(type="text", text=str(result))]

        # PowerPoint tools
        if name == "powerpoint_create_presentation":
            service = get_powerpoint_service()
            result = service.create_presentation()
            return [TextContent(type="text", text=str(result))]

        if name == "powerpoint_open_presentation":
            service = get_powerpoint_service()
            result = service.open_presentation(arguments["file_path"])
            return [TextContent(type="text", text=str(result))]

        if name == "powerpoint_save_presentation":
            service = get_powerpoint_service()
            result = service.save_presentation(arguments.get("file_path"))
            return [TextContent(type="text", text=str(result))]

        if name == "powerpoint_add_slide":
            service = get_powerpoint_service()
            result = service.add_slide(arguments.get("layout", 2))
            return [TextContent(type="text", text=str(result))]

        if name == "powerpoint_modify_title":
            service = get_powerpoint_service()
            result = service.modify_title(
                arguments["slide_index"], arguments["title_text"]
            )
            return [TextContent(type="text", text=str(result))]

        if name == "powerpoint_insert_image":
            service = get_powerpoint_service()
            result = service.insert_image(
                arguments["slide_index"],
                arguments["image_path"],
                arguments["left"],
                arguments["top"],
                arguments.get("width"),
                arguments.get("height"),
            )
            return [TextContent(type="text", text=str(result))]

        if name == "powerpoint_export_to_pdf":
            service = get_powerpoint_service()
            result = service.export_to_pdf(arguments["output_path"])
            return [TextContent(type="text", text=str(result))]

        return [TextContent(type="text", text=f"Unknown tool: {name}")]

    except Exception as e:
        logger.error(f"Error calling tool {name}: {e}", exc_info=True)
        return [TextContent(type="text", text=f"Error: {e!s}")]


async def main() -> None:
    """Run the MCP server."""
    async with stdio_server() as (read_stream, write_stream):
        await app.run(
            read_stream,
            write_stream,
            app.create_initialization_options(),
        )


if __name__ == "__main__":
    asyncio.run(main())
