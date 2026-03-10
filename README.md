<p align="center">
  <img src="https://raw.githubusercontent.com/paulivascu/excel-mcp-server-stripped/main/assets/logo.png" alt="Excel MCP Server Logo" width="300"/>
</p>

[![PyPI version](https://img.shields.io/pypi/v/excel-mcp-server-stripped.svg)](https://pypi.org/project/excel-mcp-server-stripped/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Install MCP Server](https://cursor.com/deeplink/mcp-install-dark.svg)](https://cursor.com/install-mcp?name=excel-mcp-server-stripped&config=eyJjb21tYW5kIjoidXZ4IGV4Y2VsLW1jcC1zZXJ2ZXItc3RyaXBwZWQgc3RkaW8ifQ%3D%3D)

A stripped fork of the original Excel MCP Server, packaged for PyPI so it can be launched directly with `uvx`. Create, read, and modify Excel workbooks without needing Microsoft Excel installed.

## Features

- **Workbook Operations**: Create workbooks, create worksheets, and read workbook metadata
- **Data Operations**: Read and write worksheet cell data
- **Compact Read Output**: Minified JSON with `range`, `sheet_name`, and per-row cell maps to reduce token usage
- **Triple transport support**: stdio, SSE (deprecated), and streamable HTTP
- **Remote & Local**: Works both locally and as a remote service

## Usage

The server supports three transport methods:

### 1. Stdio Transport (for local use)

```bash
uvx excel-mcp-server-stripped stdio
```

```json
{
   "mcpServers": {
      "excel": {
         "command": "uvx",
         "args": ["excel-mcp-server-stripped", "stdio"]
      }
   }
}
```

If you want to keep the original executable name after installation, this package also exposes `excel-mcp-server`, so `uvx --from excel-mcp-server-stripped excel-mcp-server stdio` works as well.

### 2. SSE Transport (Server-Sent Events - Deprecated)

```bash
uvx excel-mcp-server-stripped sse
```

**SSE transport connection**:
```json
{
   "mcpServers": {
      "excel": {
         "url": "http://localhost:8000/sse",
      }
   }
}
```

### 3. Streamable HTTP Transport (Recommended for remote connections)

```bash
uvx excel-mcp-server-stripped streamable-http
```

**Streamable HTTP transport connection**:
```json
{
   "mcpServers": {
      "excel": {
         "url": "http://localhost:8000/mcp",
      }
   }
}
```

## Windows EXE

You can build a standalone Windows executable for machines that should not need Python or `uv` installed.

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-exe.ps1
```

The executable is written to `dist\excel-mcp-server-stripped.exe` and can be used directly:

```powershell
.\dist\excel-mcp-server-stripped.exe stdio
```

For Goose or any other MCP client that launches a local process, point the extension command at the exe:

```json
{
  "mcpServers": {
    "excel": {
      "command": "C:\\path\\to\\excel-mcp-server-stripped.exe",
      "args": ["stdio"]
    }
  }
}
```

## GitHub Release

You can create a GitHub release and attach the built executable with the GitHub CLI:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\release-github.ps1
```

To create a draft release instead of publishing immediately:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\release-github.ps1 -Draft
```

To attach the wheel and source distribution as well:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\release-github.ps1 -IncludePythonArtifacts
```

The release script requires a clean git working tree and an authenticated `gh` session. Publishing a release will also trigger the PyPI workflow.

## Environment Variables & File Path Handling

### SSE and Streamable HTTP Transports

When running the server with the **SSE or Streamable HTTP protocols**, you **must set the `EXCEL_FILES_PATH` environment variable on the server side**. This variable tells the server where to read and write Excel files.
- If not set, it defaults to `./excel_files`.

You can also set the `FASTMCP_PORT` environment variable to control the port the server listens on (default is `8017` if not set).
- Example (Windows PowerShell):
  ```powershell
  $env:EXCEL_FILES_PATH="E:\MyExcelFiles"
  $env:FASTMCP_PORT="8007"
  uvx excel-mcp-server-stripped streamable-http
  ```
- Example (Linux/macOS):
  ```bash
  EXCEL_FILES_PATH=/path/to/excel_files FASTMCP_PORT=8007 uvx excel-mcp-server-stripped streamable-http
  ```

### Stdio Transport

When using the **stdio protocol**, the file path is provided with each tool call, so you do **not** need to set `EXCEL_FILES_PATH` on the server. The server will use the path sent by the client for each operation.

## Available Tools

The server exposes workbook and data tools only. See [TOOLS.md](TOOLS.md) for complete documentation.

For `write_data_to_excel`, `data` must be a plain 2D array of scalar cell values such as `[['ID', 'Value1'], [1, 56]]`. Do not wrap cells in objects like `{"value":"ID"}`.

### Compact Read Response

`read_data_from_excel` returns a compact minified JSON payload that keeps range/sheet context and cell coordinates while reducing token usage:

```json
{"range":"A1:H10","sheet_name":"Model_Map","cells":[{"A1":"Path","B1":"BlockType"},{"A2":"DDD_test","B2":null}]}
```

`read_data_from_excel` has a hard cap of `50` rows per call. When more rows are requested, the result is auto-truncated and includes pagination guidance:

```json
{"range":"A1:B50","sheet_name":"Model_Map","cells":[...],"truncated":true,"max_rows_per_call":50,"message":"Maximum rows per call is 50. Result was truncated. Make multiple calls to read additional rows.","next_start_cell":"A51"}
```

## Repository

Source: [paulivascu/excel-mcp-server-stripped](https://github.com/paulivascu/excel-mcp-server-stripped)

## License

MIT License - see [LICENSE](LICENSE) for details.

