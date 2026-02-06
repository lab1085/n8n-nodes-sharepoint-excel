# WAC Tokens & Microsoft Graph Excel API Research

**Date:** 2026-02-06

## Executive Summary

The `/workbook/` endpoints (which use WAC) only support delegated permissions - they do NOT support application permissions. This is a Microsoft limitation, not something platforms like Make.com have solved.

---

## What Are WAC Tokens?

WAC (Web Application Companion) tokens are authentication credentials required by Microsoft's Excel Online service to perform edit operations on Excel workbooks stored in SharePoint Online or OneDrive.

**Key Architecture:**

- WAC operates through the WOPI protocol (Web Open Platform Interface)
- OAuth tokens are used for communication between SharePoint and Office Online Server
- The token is included in WOPI requests between SharePoint and Office Online Server
- WAC relies on Office Online Server infrastructure to handle document editing

---

## Permission Requirements

From [Microsoft's official documentation](https://learn.microsoft.com/en-us/graph/api/workbook-createsession?view=graph-rest-1.0):

| Permission Type                | Supported         |
| ------------------------------ | ----------------- |
| Delegated (work/school)        | `Files.ReadWrite` |
| Delegated (personal Microsoft) | Not supported     |
| **Application**                | **Not supported** |

**Critical implications:**

- Every Excel workbook operation requires a **signed-in user context**
- You cannot use client credentials flow (app-only) for automation
- WAC token generation happens internally and requires Office Online Server

---

## Graph API Endpoints: /workbook/ vs /content

### /workbook/ Endpoints (WAC-Dependent)

```
POST /me/drive/items/{id}/workbook/createSession
GET  /me/drive/items/{id}/workbook/worksheets
POST /me/drive/items/{id}/workbook/worksheets/{id}/range
```

**Characteristics:**

- Direct API calls to read/write cells, tables, worksheets
- Support real-time collaboration
- Require WAC token internally (non-transparent to caller)
- Return structured JSON with worksheet, table, and cell data
- Only supports `.xlsx` files (not `.xls`)

### /content Endpoint (No WAC)

```
GET  /drives/{drive-id}/items/{item-id}/content
PUT  /drives/{drive-id}/items/{item-id}/content
```

**Characteristics:**

- Returns a 302 redirect to a pre-authenticated download URL
- Downloads/uploads the complete binary Excel file
- No Office Online Server dependency
- Works with any Excel file format
- More reliable for SharePoint files

---

## Why WAC Fails on SharePoint

| Cause                                | Description                                  |
| ------------------------------------ | -------------------------------------------- |
| **File format**                      | `.xls` files not supported - must be `.xlsx` |
| **No Office 365 license**            | User needs E3+ with Excel Online access      |
| **Tenant policies**                  | Corporate security may block WAC             |
| **Office Online Server unreachable** | SharePoint can't contact the server          |
| **Personal accounts**                | Not supported at all                         |
| **Orphaned sessions**                | Previous sessions didn't close properly      |

### Error Codes

The node should detect these error patterns:

- HTTP status `403` with message "Could not obtain a WAC access token"
- HTTP status `423` (Locked)
- Error code `resourceLocked`
- Error code `AccessDenied` with WAC message

---

## How Other Platforms Handle This

### Make.com

From [Make.com's Excel documentation](https://apps.make.com/microsoft-excel):

1. **Uses OAuth2 with user sign-in** (delegated permissions)
2. **Provides a "Download workbook" action** - falls back to `/content` endpoint
3. **Offers location selection**: My Drive, Site's Drive, or Group Drive
4. **Uses Microsoft Graph API** under the hood

Make.com doesn't "solve" WAC - they offer both approaches:

- Direct API calls (which can fail with WAC errors)
- File download/upload (which bypasses WAC entirely)

### n8n Native Node

From [n8n community discussions](https://community.n8n.io/t/support-for-sharepoint-in-microsoft-excel-node/57538):

1. **Only searches OneDrive**: Uses `/drive/root/search(q='.xlsx')`
2. **No SharePoint support**: Cannot browse SharePoint sites/drives
3. **Relies on WAC**: All operations use `/workbook/` endpoints
4. **Known issue**: [GitHub #20040](https://github.com/n8n-io/n8n/issues/20040) - closed as "stale" without fix

### Power Automate

- Uses both `/workbook/` and `/content` endpoints internally
- Microsoft can update implementations as WAC issues arise
- Has fallback mechanisms for WAC failures
- User permissions flow through Azure AD seamlessly

---

## Session Management Best Practices

From [Microsoft Graph Best Practices](https://learn.microsoft.com/en-us/graph/workbook-best-practice):

### Creating a Session

```http
POST /me/drive/items/{id}/workbook/createSession
Content-type: application/json

{
  "persistChanges": true
}
```

### Using Session ID

Include in all subsequent requests:

```
workbook-session-id: {session-id}
```

### Session Lifecycle

- Persistent sessions: ~5 minutes inactivity timeout
- Non-persistent sessions: ~7 minutes inactivity timeout
- Keep-alive strategy: Call `refreshSession` endpoint periodically

### Throttling Guidelines

- **Don't parallelize requests** to the same workbook, especially writes
- **Use sequential operations**: Send next request only after receiving response
- **Respect Retry-After header** in throttling responses
- Concurrent writes cause: throttling errors, timeouts, merge conflicts

---

## Comparison: Approaches

| Aspect                | /workbook/ Endpoints    | /content Endpoint         |
| --------------------- | ----------------------- | ------------------------- |
| **Authentication**    | OAuth + WAC token       | OAuth only                |
| **File format**       | .xlsx only              | Any format                |
| **Speed**             | Slow (WAC negotiation)  | Fast (direct download)    |
| **Reliability**       | ~60-70% on SharePoint   | 99%+                      |
| **Real-time collab**  | Yes                     | No                        |
| **Requires sessions** | Recommended             | No                        |
| **Office Server**     | Yes                     | No                        |
| **Code complexity**   | Simple API calls        | Download → parse → upload |
| **Best use case**     | OneDrive personal files | SharePoint automation     |

---

## Recommended Architecture: Hybrid Approach

A professional node should implement both approaches with automatic fallback:

```
┌─────────────────────────────────────────────────────────────┐
│                    SharePoint Excel Node                     │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  1. User provides siteId, driveId, fileId                   │
│                                                             │
│  2. For READ operations:                                    │
│     ├─ Try: GET /workbook/worksheets (WAC)                  │
│     │   └─ If success → use /workbook/ endpoints            │
│     └─ If 403 WAC error → download file, parse with exceljs │
│                                                             │
│  3. For WRITE operations:                                   │
│     ├─ Check if file has Excel Tables                       │
│     │   └─ If yes → warn user about corruption risk         │
│     ├─ Try: createSession + /workbook/ operations           │
│     │   └─ If success → commit changes                      │
│     └─ If 403 WAC error → download/modify/upload            │
│                                                             │
│  4. Session management:                                     │
│     ├─ Create persistent session for batch operations       │
│     ├─ Include workbook-session-id header                   │
│     └─ Close session when done                              │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

### Benefits of Hybrid Approach

1. **Try WAC first** - faster, supports real-time features
2. **Fall back to download/upload** - reliable when WAC fails
3. **Detect Tables before writes** - warn users about corruption risk
4. **Proper session management** - use `createSession` with `workbook-session-id`
5. **Better error messages** - explain _why_ WAC failed and what the fallback is

---

## Implementation Considerations

### For WAC-Based Operations

```typescript
// Create session
const session = await graphApi.post(
	`/sites/${siteId}/drives/${driveId}/items/${fileId}/workbook/createSession`,
	{ persistChanges: true },
);

// Use session in subsequent calls
const headers = {
	'workbook-session-id': session.id,
};

// Close session when done
await graphApi.post(
	`/sites/${siteId}/drives/${driveId}/items/${fileId}/workbook/closeSession`,
	{},
	{ headers },
);
```

### For Fallback Operations

```typescript
// Download file
const fileBuffer = await graphApi.get(
	`/sites/${siteId}/drives/${driveId}/items/${fileId}/content`,
	{ responseType: 'arraybuffer' },
);

// Parse with exceljs
const workbook = new ExcelJS.Workbook();
await workbook.xlsx.load(fileBuffer);

// Modify...

// Upload back
const modifiedBuffer = await workbook.xlsx.writeBuffer();
await graphApi.put(`/sites/${siteId}/drives/${driveId}/items/${fileId}/content`, modifiedBuffer);
```

### Detecting Excel Tables

```typescript
// Check for tables before write operations
const workbook = new ExcelJS.Workbook();
await workbook.xlsx.load(fileBuffer);

for (const worksheet of workbook.worksheets) {
	if (worksheet.tables && Object.keys(worksheet.tables).length > 0) {
		// Warn user about table corruption risk
	}
}
```

---

## Sources

- [Microsoft Graph Excel Best Practices](https://learn.microsoft.com/en-us/graph/workbook-best-practice)
- [workbook: createSession API](https://learn.microsoft.com/en-us/graph/api/workbook-createsession?view=graph-rest-1.0)
- [Manage sessions and persistence](https://learn.microsoft.com/en-us/graph/excel-manage-sessions)
- [Working with Excel in Microsoft Graph](https://learn.microsoft.com/en-us/graph/api/resources/excel?view=graph-rest-1.0)
- [n8n GitHub Issue #20040](https://github.com/n8n-io/n8n/issues/20040)
- [n8n Community: SharePoint Support Request](https://community.n8n.io/t/support-for-sharepoint-in-microsoft-excel-node/57538)
- [Make.com Excel Integration](https://www.make.com/en/integrations/microsoft-excel)
- [WAC Token Error Solutions (Microsoft Q&A)](https://learn.microsoft.com/en-us/answers/questions/1654516/how-can-i-avoid-getting-a-wac-access-token-error-w)
- [Graph API Permissions Overview](https://learn.microsoft.com/en-us/graph/permissions-overview)
