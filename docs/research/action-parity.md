# Action Parity: Original Excel Node vs This Node

## Full Comparison

| Resource | Original Excel Node | This Node (exceljs) | Can Build? |
|----------|---------------------|---------------------|------------|
| **TABLE** | | | |
| | Append rows to table | — | ❌ No (limited exceljs support) |
| | Convert to range | — | ❌ No |
| | Create a table | — | ❌ No |
| | Delete a table | — | ❌ No |
| | Get columns | Get columns | ✅ Yes |
| | Get rows | Get rows | ✅ Yes |
| | Lookup a column | Lookup a column | ✅ Yes |
| **WORKBOOK** | | | |
| | Add a sheet to a workbook | Add sheet | ✅ Yes |
| | Delete workbook | Delete workbook | ✅ Yes (Graph API) |
| | Get workbooks | Get workbooks | ✅ Yes (Graph API) |
| **SHEET** | | | |
| | Append data to sheet | Append rows | ✅ Have |
| | Append or update a sheet | Upsert rows | ✅ Yes |
| | Clear sheet | Clear sheet | ✅ Yes |
| | Delete sheet | Delete sheet | ✅ Yes |
| | Get sheets | Get sheets | ✅ Have |
| | Get rows from sheet | Read rows | ✅ Have |
| | Update sheet | Update range | ✅ Have (cell only) |

## Final Action List

Dropping table write operations due to exceljs limitations:

```
TABLE (read-only)
  - Get columns
  - Get rows
  - Lookup a column

WORKBOOK
  - Add sheet
  - Delete workbook
  - Get workbooks

SHEET
  - Append rows        ✅ implemented
  - Upsert rows        (append or update by key)
  - Clear sheet
  - Delete sheet
  - Get sheets         ✅ implemented
  - Read rows          ✅ implemented
  - Update range       ⚠️ partial (single cell only)
```

## Why Table Write Operations Are Dropped

Excel Tables are a specific feature ("Format as Table") that creates structured, named ranges with:
- Auto-expanding boundaries
- Built-in filtering/sorting
- Structured references (`=Table1[Column]`)

The exceljs library has limited support for creating/modifying these table definitions. It can **read** existing tables but cannot reliably:
- Create new tables with proper formatting
- Delete tables while preserving underlying data
- Convert tables to ranges

The original Excel node uses Graph API `/workbook/tables/*` endpoints which require WAC tokens (the exact thing this node bypasses).

## Implementation Notes

- **Sheet operations**: Work directly with cells/rows using exceljs
- **Workbook operations**: Mix of exceljs (add sheet) and Graph API (delete file, list files)
- **Table operations**: Read-only via exceljs table parsing
