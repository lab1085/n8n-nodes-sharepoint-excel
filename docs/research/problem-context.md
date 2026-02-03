# SharePoint Excel Problem Context

## The Problem

n8n's Microsoft Excel 365 node has a fundamental limitation: it was designed for OneDrive and doesn't properly support SharePoint-hosted Excel files.

### Core Issues

1. **WAC Token Errors** - The `/workbook/` Graph API endpoints fail with 403 "Could not obtain a WAC access token"
2. **No "By URL" option** - Unlike Google Sheets, you can't just paste a SharePoint URL
3. **Empty dropdowns** - Sheet/Table selectors don't populate for SharePoint files
4. **ID mismatch** - SharePoint node returns list item IDs, not Graph driveItem.id format

### References

- GitHub Issue: https://github.com/n8n-io/n8n/issues/20040
- Reddit thread: https://www.reddit.com/r/n8n/comments/1obqv7o/excel_files_on_sharepoint/

## Existing Community Nodes (None Solve This)

| Package                        | Author      | What it does                           | Solves problem? |
| ------------------------------ | ----------- | -------------------------------------- | --------------- |
| n8n-nodes-microsoft-sharepoint | Savjee      | File operations only                   | No              |
| n8n-nodes-community-sharepoint | arisechurch | File/folder operations                 | No              |
| @bitovi/n8n-nodes-excel        | Bitovi      | Add/delete/list sheets on local binary | Partial         |

## The Solution: Download-Edit-Upload Pattern

Bypass WAC entirely by using the `/content` endpoint instead of `/workbook/`:

```
1. GET  /drives/{driveId}/items/{fileId}/content  → Download Excel as binary
2. Parse with exceljs library
3. Perform operation (read/write/append)
4. PUT  /drives/{driveId}/items/{fileId}/content  → Upload modified file
```

## Planned Node Operations

```
┌─────────────────────────────────────────────────┐
│         n8n-nodes-sharepoint-excel              │
├─────────────────────────────────────────────────┤
│ Operations:                                     │
│   • Append Row(s)                               │
│   • Update Cell / Range                         │
│   • Read Rows                                   │
│   • Delete Row                                  │
│   • Get Sheet Names                             │
│   • Create Sheet                                │
├─────────────────────────────────────────────────┤
│ Under the hood:                                 │
│   1. GET  /drives/{id}/items/{id}/content       │
│   2. Parse with exceljs                         │
│   3. Perform operation                          │
│   4. PUT  /drives/{id}/items/{id}/content       │
└─────────────────────────────────────────────────┘
```

## Technical Stack

| Component     | Choice                       | Why                                    |
| ------------- | ---------------------------- | -------------------------------------- |
| Language      | TypeScript                   | n8n requirement                        |
| Excel library | exceljs                      | Preserves formatting, formulas, styles |
| Auth          | microsoftSharePointOAuth2Api | Reuse existing n8n pattern             |
| API           | Graph `/content` endpoint    | Bypasses WAC entirely                  |
