# Excel MCP Server Tools

This server intentionally exposes only workbook operations and data operations.

## Workbook Operations

### create_workbook

Creates a new Excel workbook.

```python
create_workbook(filepath: str) -> str
```

- `filepath`: Path where to create workbook
- Returns: Success message with created file path

### create_worksheet

Creates a new worksheet in an existing workbook.

```python
create_worksheet(filepath: str, sheet_name: str) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Name for the new worksheet
- Returns: Success message

### get_workbook_metadata

Get metadata about workbook including sheets and ranges.

```python
get_workbook_metadata(filepath: str, include_ranges: bool = False) -> str
```

- `filepath`: Path to Excel file
- `include_ranges`: Whether to include range information
- Returns: String representation of workbook metadata

## Data Operations

### write_data_to_excel

Write data to Excel worksheet.

```python
write_data_to_excel(
    filepath: str,
    sheet_name: str,
    data: List[List],
    start_cell: str = "A1"
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Target worksheet name
- `data`: List of rows to write
- `start_cell`: Starting cell (default: "A1")
- Returns: Success message

### read_data_from_excel

Read data from Excel worksheet.

```python
read_data_from_excel(
    filepath: str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: str = None,
    row_limit: int = None
) -> str
```

- `filepath`: Path to Excel file
- `sheet_name`: Source worksheet name
- `start_cell`: Starting cell (default: "A1")
- `end_cell`: Optional ending cell to bound width/height
- `row_limit`: Optional number of rows to return starting from `start_cell`
- Hard cap: A single call returns at most `50` rows.
- Returns: Minified JSON string:
  - `{"range":"A1:H10","sheet_name":"Model_Map","cells":[{"A1":"Path","B1":"BlockType"}, ...]}`
  - If request exceeds 50 rows, response is auto-truncated and includes:
    `{"truncated":true,"max_rows_per_call":50,"message":"Maximum rows per call is 50. Result was truncated. Make multiple calls to read additional rows.","next_start_cell":"A51"}`

Example:

```python
read_data_from_excel(
    filepath="D:\\00_work\\00_BTC\\btc-agent\\Traceability_KB.xlsx",
    sheet_name="Model_Map",
    start_cell="A1",
    row_limit=10
)
```

Pagination example (follow-up call after truncation):

```python
read_data_from_excel(
    filepath="D:\\00_work\\00_BTC\\btc-agent\\Traceability_KB.xlsx",
    sheet_name="Model_Map",
    start_cell="A51",
    row_limit=50
)
```
