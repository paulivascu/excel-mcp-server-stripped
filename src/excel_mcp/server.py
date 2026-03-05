import json
import logging
import os
from typing import List, Optional

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

from excel_mcp.data import write_data
from excel_mcp.exceptions import DataError, ValidationError, WorkbookError
from excel_mcp.workbook import get_workbook_info

# Get project root directory path for log file path.
# When using stdio transport, writing logs to stdout breaks MCP protocol,
# so log to a file under project root.
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
LOG_FILE = os.path.join(ROOT_DIR, "excel-mcp.log")

# In SSE/HTTP modes this is populated from environment.
EXCEL_FILES_PATH = None

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler(LOG_FILE)],
)
logger = logging.getLogger("excel-mcp")

mcp = FastMCP(
    "excel-mcp",
    host=os.environ.get("FASTMCP_HOST", "0.0.0.0"),
    port=int(os.environ.get("FASTMCP_PORT", "8017")),
    instructions="Excel MCP Server for workbook and data operations",
)


def get_excel_path(filename: str) -> str:
    """Get full path to Excel file."""
    if os.path.isabs(filename):
        return filename

    if EXCEL_FILES_PATH is None:
        raise ValueError(
            f"Invalid filename: {filename}, must be an absolute path when not in SSE/HTTP mode"
        )

    return os.path.join(EXCEL_FILES_PATH, filename)


@mcp.tool(
    annotations=ToolAnnotations(
        title="Read Data from Excel",
        readOnlyHint=True,
    ),
)
def read_data_from_excel(
    filepath: str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: Optional[str] = None,
    row_limit: Optional[int] = None,
) -> str:
    """
    Read data from Excel worksheet in compact row-map format.

    Returns minified JSON:
    {"range":"A1:H10","sheet_name":"Sheet1","cells":[{"A1":...,"B1":...},...]}
    """
    try:
        full_path = get_excel_path(filepath)
        from excel_mcp.data import read_excel_range_as_row_maps

        result = read_excel_range_as_row_maps(
            full_path,
            sheet_name,
            start_cell,
            end_cell,
            row_limit,
        )
        return json.dumps(result, separators=(",", ":"), ensure_ascii=False, default=str)
    except Exception as e:
        logger.error(f"Error reading data: {e}")
        raise


@mcp.tool(
    annotations=ToolAnnotations(
        title="Write Data to Excel",
        destructiveHint=True,
    ),
)
def write_data_to_excel(
    filepath: str,
    sheet_name: str,
    data: List[List],
    start_cell: str = "A1",
) -> str:
    """Write data to Excel worksheet."""
    try:
        full_path = get_excel_path(filepath)
        result = write_data(full_path, sheet_name, data, start_cell)
        return result["message"]
    except (ValidationError, DataError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error writing data: {e}")
        raise


@mcp.tool(
    annotations=ToolAnnotations(
        title="Create Workbook",
        destructiveHint=True,
    ),
)
def create_workbook(filepath: str) -> str:
    """Create new Excel workbook."""
    try:
        full_path = get_excel_path(filepath)
        from excel_mcp.workbook import create_workbook as create_workbook_impl

        create_workbook_impl(full_path)
        return f"Created workbook at {full_path}"
    except WorkbookError as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error creating workbook: {e}")
        raise


@mcp.tool(
    annotations=ToolAnnotations(
        title="Create Worksheet",
        destructiveHint=True,
    ),
)
def create_worksheet(filepath: str, sheet_name: str) -> str:
    """Create new worksheet in workbook."""
    try:
        full_path = get_excel_path(filepath)
        from excel_mcp.workbook import create_sheet as create_worksheet_impl

        result = create_worksheet_impl(full_path, sheet_name)
        return result["message"]
    except (ValidationError, WorkbookError) as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error creating worksheet: {e}")
        raise


@mcp.tool(
    annotations=ToolAnnotations(
        title="Get Workbook Metadata",
        readOnlyHint=True,
    ),
)
def get_workbook_metadata(
    filepath: str,
    include_ranges: bool = False,
) -> str:
    """Get metadata about workbook including sheets and ranges."""
    try:
        full_path = get_excel_path(filepath)
        result = get_workbook_info(full_path, include_ranges=include_ranges)
        return str(result)
    except WorkbookError as e:
        return f"Error: {str(e)}"
    except Exception as e:
        logger.error(f"Error getting workbook metadata: {e}")
        raise


def run_sse():
    """Run Excel MCP server in SSE mode."""
    global EXCEL_FILES_PATH
    EXCEL_FILES_PATH = os.environ.get("EXCEL_FILES_PATH", "./excel_files")
    os.makedirs(EXCEL_FILES_PATH, exist_ok=True)

    try:
        logger.info(
            f"Starting Excel MCP server with SSE transport (files directory: {EXCEL_FILES_PATH})"
        )
        mcp.run(transport="sse")
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception as e:
        logger.error(f"Server failed: {e}")
        raise
    finally:
        logger.info("Server shutdown complete")


def run_streamable_http():
    """Run Excel MCP server in streamable HTTP mode."""
    global EXCEL_FILES_PATH
    EXCEL_FILES_PATH = os.environ.get("EXCEL_FILES_PATH", "./excel_files")
    os.makedirs(EXCEL_FILES_PATH, exist_ok=True)

    try:
        logger.info(
            "Starting Excel MCP server with streamable HTTP transport "
            f"(files directory: {EXCEL_FILES_PATH})"
        )
        mcp.run(transport="streamable-http")
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception as e:
        logger.error(f"Server failed: {e}")
        raise
    finally:
        logger.info("Server shutdown complete")


def run_stdio():
    """Run Excel MCP server in stdio mode."""
    try:
        logger.info("Starting Excel MCP server with stdio transport")
        mcp.run(transport="stdio")
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception as e:
        logger.error(f"Server failed: {e}")
        raise
    finally:
        logger.info("Server shutdown complete")
