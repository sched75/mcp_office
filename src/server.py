"""
MCP Office Server - Complete Integration (295+ tools).

PRODUCTION-READY server with ALL Office services integrated.

Architecture:
- Word: 59 tools
- Excel: 82 tools
- PowerPoint: 63 tools
- Outlook: 67 tools
Total: 271 tools

Author: Pascal-Louis
Version: 3.0.0 - Complete Integration
"""

import asyncio
import logging
from typing import Any, Dict, Optional, Callable

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

from src.word.word_service import WordService
from src.excel.excel_service import ExcelService
from src.powerpoint.powerpoint_service import PowerPointService
from src.outlook.outlook_service import OutlookService
from src.core.exceptions import (
    COMInitializationError,
    DocumentNotFoundError,
    InvalidParameterError,
)

# Import des configurations d'outils
from tools_configs import (
    WORD_TOOLS_CONFIG,
    EXCEL_TOOLS_CONFIG,
    POWERPOINT_TOOLS_CONFIG,
    OUTLOOK_TOOLS_CONFIG,
)

# Configuration du logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger("mcp_office")

# Initialisation du serveur MCP
app = Server("mcp-office-server")

# Services Office
word_service: Optional[WordService] = None
excel_service: Optional[ExcelService] = None
powerpoint_service: Optional[PowerPointService] = None
outlook_service: Optional[OutlookService] = None


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def format_result(result: Dict[str, Any]) -> str:
    """Formate un résultat pour l'affichage."""
    if not result.get("success", False):
        error_msg = result.get("error", "Unknown error")
        return f"❌ Erreur: {error_msg}"

    lines = ["✅ Opération réussie"]
    for key, value in result.items():
        if key not in ["success", "error"] and value is not None:
            if isinstance(value, (list, dict)) and len(str(value)) > 200:
                lines.append(f"  • {key}: [données volumineuses - {len(value)} éléments]")
            else:
                lines.append(f"  • {key}: {value}")

    return "\n".join(lines)


def validate_parameters(params: Dict[str, Any], required: list[str]) -> None:
    """Valide la présence des paramètres requis."""
    missing = [p for p in required if p not in params or params[p] is None]
    if missing:
        raise InvalidParameterError(f"Paramètres manquants: {', '.join(missing)}")


def generate_tool(service_prefix: str, name: str, config: dict) -> Tool:
    """Génère un outil MCP à partir de sa configuration."""
    properties = {}
    
    # Ajouter tous les paramètres (requis et optionnels)
    for param in config.get("required", []) + config.get("optional", []):
        properties[param] = {"type": "string"}  # Type par défaut
    
    return Tool(
        name=f"{service_prefix}_{name}",
        description=config["desc"],
        inputSchema={
            "type": "object",
            "properties": properties,
            "required": config.get("required", []),
        },
    )


def build_handlers(service: Any, service_config: Dict, service_prefix: str) -> Dict[str, Callable]:
    """Construit le mapping des handlers de manière dynamique."""
    
    def create_handler(method_name: str, config: dict):
        """Crée un handler dynamique pour une méthode."""
        method = getattr(service, method_name)
        
        def handler(args: dict):
            # Extraire tous les arguments (requis + optionnels)
            kwargs = {}
            for param in config.get("required", []) + config.get("optional", []):
                if param in args:
                    kwargs[param] = args[param]
            return method(**kwargs)
        
        return handler
    
    # Générer tous les handlers automatiquement
    handlers = {}
    for method_name, config in service_config.items():
        tool_name = f"{service_prefix}_{method_name}"
        handlers[tool_name] = create_handler(method_name, config)
    
    return handlers


# =============================================================================
# MCP SERVER HANDLERS
# =============================================================================

@app.list_tools()
async def list_tools() -> list[Tool]:
    """Liste tous les outils disponibles (271 outils)."""
    tools = []
    
    # Générer les outils Word
    for method_name, config in WORD_TOOLS_CONFIG.items():
        tools.append(generate_tool("word", method_name, config))
    
    # Générer les outils Excel
    for method_name, config in EXCEL_TOOLS_CONFIG.items():
        tools.append(generate_tool("excel", method_name, config))
    
    # Générer les outils PowerPoint
    for method_name, config in POWERPOINT_TOOLS_CONFIG.items():
        tools.append(generate_tool("powerpoint", method_name, config))
    
    # Générer les outils Outlook
    for method_name, config in OUTLOOK_TOOLS_CONFIG.items():
        tools.append(generate_tool("outlook", method_name, config))
    
    logger.info(f"Loaded {len(tools)} tools total")
    logger.info(f"  - Word: {len(WORD_TOOLS_CONFIG)} tools")
    logger.info(f"  - Excel: {len(EXCEL_TOOLS_CONFIG)} tools")
    logger.info(f"  - PowerPoint: {len(POWERPOINT_TOOLS_CONFIG)} tools")
    logger.info(f"  - Outlook: {len(OUTLOOK_TOOLS_CONFIG)} tools")
    
    return tools


@app.call_tool()
async def call_tool(name: str, arguments: Any) -> list[TextContent]:
    """Exécute un outil MCP avec routing automatique."""
    logger.info(f"Calling tool: {name}")
    
    try:
        # Convertir arguments en dictionnaire
        if not isinstance(arguments, dict):
            arguments = {}
        
        result = None
        
        # === WORD TOOLS ===
        if name.startswith("word_"):
            if word_service is None:
                raise COMInitializationError("Word service not initialized")
            
            handlers = build_handlers(word_service, WORD_TOOLS_CONFIG, "word")
            
            if name in handlers:
                result = handlers[name](arguments)
            else:
                return [TextContent(
                    type="text",
                    text=f"❌ Outil Word non implémenté: {name}"
                )]
        
        # === EXCEL TOOLS ===
        elif name.startswith("excel_"):
            if excel_service is None:
                raise COMInitializationError("Excel service not initialized")
            
            handlers = build_handlers(excel_service, EXCEL_TOOLS_CONFIG, "excel")
            
            if name in handlers:
                result = handlers[name](arguments)
            else:
                return [TextContent(
                    type="text",
                    text=f"❌ Outil Excel non implémenté: {name}"
                )]
        
        # === POWERPOINT TOOLS ===
        elif name.startswith("powerpoint_"):
            if powerpoint_service is None:
                raise COMInitializationError("PowerPoint service not initialized")
            
            handlers = build_handlers(powerpoint_service, POWERPOINT_TOOLS_CONFIG, "powerpoint")
            
            if name in handlers:
                result = handlers[name](arguments)
            else:
                return [TextContent(
                    type="text",
                    text=f"❌ Outil PowerPoint non implémenté: {name}"
                )]
        
        # === OUTLOOK TOOLS ===
        elif name.startswith("outlook_"):
            if outlook_service is None:
                raise COMInitializationError("Outlook service not initialized")
            
            handlers = build_handlers(outlook_service, OUTLOOK_TOOLS_CONFIG, "outlook")
            
            if name in handlers:
                result = handlers[name](arguments)
            else:
                return [TextContent(
                    type="text",
                    text=f"❌ Outil Outlook non implémenté: {name}"
                )]
        
        else:
            return [TextContent(type="text", text=f"❌ Outil inconnu: {name}")]
        
        # Formater et retourner le résultat
        if result is None:
            return [TextContent(type="text", text=f"❌ Aucun résultat retourné")]
        
        formatted = format_result(result)
        return [TextContent(type="text", text=formatted)]
    
    except InvalidParameterError as e:
        logger.error(f"Invalid parameters for {name}: {e}")
        return [TextContent(type="text", text=f"❌ Paramètres invalides: {e}")]
    
    except DocumentNotFoundError as e:
        logger.error(f"Document not found: {e}")
        return [TextContent(type="text", text=f"❌ Document non trouvé: {e}")]
    
    except COMInitializationError as e:
        logger.error(f"COM initialization error: {e}")
        return [TextContent(type="text", text=f"❌ Erreur d'initialisation: {e}")]
    
    except Exception as e:
        logger.exception(f"Error calling tool {name}")
        return [TextContent(type="text", text=f"❌ Erreur inattendue: {str(e)}")]


# =============================================================================
# LIFECYCLE MANAGEMENT
# =============================================================================

async def initialize_services():
    """Initialise tous les services Office."""
    global word_service, excel_service, powerpoint_service, outlook_service
    
    logger.info("Initializing Office services...")
    
    try:
        # Initialiser Word
        word_service = WordService()
        word_service.initialize()
        logger.info(f"✅ Word service initialized ({len(WORD_TOOLS_CONFIG)} tools)")
        
        # Initialiser Excel
        excel_service = ExcelService()
        excel_service.initialize()
        logger.info(f"✅ Excel service initialized ({len(EXCEL_TOOLS_CONFIG)} tools)")
        
        # Initialiser PowerPoint
        powerpoint_service = PowerPointService()
        powerpoint_service.initialize()
        logger.info(f"✅ PowerPoint service initialized ({len(POWERPOINT_TOOLS_CONFIG)} tools)")
        
        # Initialiser Outlook
        outlook_service = OutlookService()
        outlook_service.initialize()
        logger.info(f"✅ Outlook service initialized ({len(OUTLOOK_TOOLS_CONFIG)} tools)")
        
        total_tools = (len(WORD_TOOLS_CONFIG) + len(EXCEL_TOOLS_CONFIG) + 
                      len(POWERPOINT_TOOLS_CONFIG) + len(OUTLOOK_TOOLS_CONFIG))
        
        logger.info("🚀 All Office services ready!")
        logger.info(f"📊 Total tools available: {total_tools}")
    
    except Exception as e:
        logger.error(f"Failed to initialize services: {e}")
        raise


async def cleanup_services():
    """Nettoie tous les services Office."""
    logger.info("Cleaning up Office services...")
    
    try:
        if word_service:
            word_service.cleanup()
        if excel_service:
            excel_service.cleanup()
        if powerpoint_service:
            powerpoint_service.cleanup()
        if outlook_service:
            outlook_service.cleanup()
        
        logger.info("✅ All services cleaned up")
    except Exception as e:
        logger.error(f"Error during cleanup: {e}")


async def main():
    """Point d'entrée principal du serveur MCP."""
    logger.info("Starting MCP Office Server...")
    logger.info("=" * 80)
    logger.info("MCP Office - Complete Office Automation Server")
    logger.info("=" * 80)
    
    try:
        # Initialiser les services
        await initialize_services()
        
        # Démarrer le serveur MCP
        async with stdio_server() as (read_stream, write_stream):
            await app.run(
                read_stream,
                write_stream,
                app.create_initialization_options(),
            )
    
    except KeyboardInterrupt:
        logger.info("Server interrupted by user")
    except Exception as e:
        logger.exception("Fatal error in server")
        raise
    finally:
        # Nettoyer les services
        await cleanup_services()
        logger.info("Server stopped")


if __name__ == "__main__":
    asyncio.run(main())
