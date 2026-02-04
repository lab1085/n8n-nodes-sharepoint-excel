# Bug & Inefficiency Report: n8n-nodes-sharepoint-excel

**Date:** 2026-02-04
**Analyzer:** Claude Code

## Critical Bugs

### 1. Table operations use WAC token endpoints

**Files:**

- `nodes/SharePointExcel/listSearch.ts:235`
- `nodes/SharePointExcel/actions/table/getRows.ts:40-58`
- `nodes/SharePointExcel/actions/table/lookup.ts:32-47`
- `nodes/SharePointExcel/actions/table/getColumns.ts:25`

These operations use `/workbook/tables/*` Graph API endpoints:

```typescript
const endpoint = `${context.basePath}/workbook/tables/${encodeURIComponent(tableName)}/columns`;
```

This contradicts the node's core purpose - these endpoints require WAC tokens, which is the exact issue this node was built to bypass.

**Impact:** Table operations may fail with the same 403 WAC token errors as the native n8n Excel node.

---

## Medium Bugs

### 2. ~~clearSheet inefficiently deletes rows one-by-one~~ ✅ FIXED

**File:** `nodes/SharePointExcel/actions/sheet/clearSheet.ts`

**Fix:** Now uses `worksheet.spliceRows(1, rowCount)` to delete all rows in a single call instead of iterating row-by-row.

---

### 3. ~~Silent error swallowing in listSearch methods~~ ✅ FIXED

**File:** `nodes/SharePointExcel/listSearch.ts`

**Fix:** All 5 search methods (`searchSites`, `getDrives`, `getFiles`, `getSheets`, `getTables`) now log errors using `this.logger.error()` instead of silently swallowing them. Errors appear in n8n server logs for debugging while still returning empty results gracefully (required for dropdown UI).

---

### 4. ~~resourceMapping.ts also swallows errors~~ ✅ FIXED

**File:** `nodes/SharePointExcel/resourceMapping.ts:127-131`

**Fix:** Now logs errors using `this.logger.error()` instead of silently swallowing them. Errors appear in n8n server logs for debugging while still returning empty fields gracefully (required for UI).

---

## Inefficiencies

### 5. getWorkbooks doesn't use siteId

**File:** `nodes/SharePointExcel/actions/workbook/getWorkbooks.ts:13`

```typescript
const endpoint = `/drives/${context.driveId}/root/children`;
```

The endpoint doesn't include `siteId` in the path, which may cause issues with SharePoint site-specific drives vs personal OneDrive. Should be `/sites/${siteId}/drives/${driveId}/root/children` for consistency.

---

### 6. Redundant getResourceValue function defined twice

**Files:**

- `nodes/SharePointExcel/listSearch.ts:20-22`
- `nodes/SharePointExcel/resourceMapping.ts:14-16`

```typescript
function getResourceValue(param: string | ResourceLocatorValue): string {
	return typeof param === 'object' ? param.value : param;
}
```

Same utility duplicated in two files. Should be extracted to a shared module.

---

### 7. Missing pairedItem tracking

Most operations don't set `pairedItem` on output items, which breaks n8n's data lineage tracking for debugging workflows.

**Affected files:** Most action handlers.

---

### 8. readRows doesn't handle filtered columns edge case

**File:** `nodes/SharePointExcel/actions/sheet/readRows.ts:94-117`

The `hasData` check only looks at filtered columns. If the user filters to columns that are empty but other columns have data, rows are skipped incorrectly.

---

### 9. Table lookup uses string comparison

**File:** `nodes/SharePointExcel/actions/table/lookup.ts:57`

```typescript
if (String(cellValue) === String(lookupValue))
```

Numeric values like `123` won't match string `"123"` intuitively, and dates will fail matching entirely.

---

## Type Safety Issues

### 10. Cell value type assertions are too broad

**Multiple files**

```typescript
const value = cell.value as string | number | boolean | null;
```

ExcelJS `CellValue` can also be `Date`, `CellErrorValue`, `CellRichTextValue`, `CellHyperlinkValue`, `CellFormulaValue`, etc. These are silently cast and may produce unexpected output.

---

### 11. GraphDriveItem.file typed as object

**File:** `nodes/SharePointExcel/types.ts:67`

```typescript
file?: object;
```

Should be more specific or use Graph API types.

---

## Minor Issues

### 12. Credential icon path may be incorrect

**File:** `credentials/MicrosoftGraphOAuth2Api.credentials.ts:12`

```typescript
icon = 'file:icons/Microsoft.svg' as const;
```

The `icons/` folder doesn't exist in the credentials directory.

---

### 13. GRAPH_BASE_URL duplicated

Defined in three files:

- `nodes/SharePointExcel/api.ts:11`
- `nodes/SharePointExcel/listSearch.ts:14`
- `nodes/SharePointExcel/resourceMapping.ts:8`

Should be a shared constant exported from `api.ts`.

---

### 14. REQUEST_TIMEOUT only 30 seconds

**File:** `nodes/SharePointExcel/api.ts:13`

```typescript
const REQUEST_TIMEOUT = 30000;
```

Large Excel files may take longer to download/upload over slow connections.

---

## Summary

| Severity     | Count            |
| ------------ | ---------------- |
| Critical     | 1                |
| Medium       | 3 (3 fixed)      |
| Inefficiency | 5                |
| Type Safety  | 2                |
| Minor        | 3                |
| **Total**    | **11 remaining** |

## Recommended Fix Priority

1. **Table operations using WAC endpoints (#1)** - Fundamental design conflict with node's purpose
2. **getWorkbooks missing siteId (#5)** - May cause incorrect behavior

~~clearSheet performance (#2)~~ ✅ Fixed
~~Silent error swallowing (#3, #4)~~ ✅ Fixed

## Notes

- The table operations (#1) represent a design decision that needs discussion - either document that table operations may fail on SharePoint, or reimplement using ExcelJS table parsing (read-only) instead of Graph API.

## Verified Non-Issues

The following were initially flagged but verified working correctly via unit tests:

- **appendRows column mapping** - ExcelJS `addRow` with sparse arrays uses 1-indexed column mapping, so data is written to correct columns. Verified with 8 passing tests in `appendRows.test.ts`.
- **upsertRows column mapping** - Same pattern as appendRows, works correctly.
