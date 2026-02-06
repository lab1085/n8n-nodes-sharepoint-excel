# Security and Permissions Guide

This document covers Microsoft Graph API permissions, security considerations, and best practices for the n8n SharePoint Excel node.

## OAuth Scopes

### Required Scopes

```
openid offline_access Sites.Read.All Files.ReadWrite.All
```

| Scope                 | Purpose                                         |
| --------------------- | ----------------------------------------------- |
| `openid`              | Required for OAuth2 authentication              |
| `offline_access`      | Enables token refresh without re-authentication |
| `Sites.Read.All`      | Browse SharePoint sites and list drives         |
| `Files.ReadWrite.All` | Download and upload Excel files                 |

### Scope Breakdown

| Scope                 | What it does                                        | Why we need it                                                                                                              |
| --------------------- | --------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------- |
| `openid`              | Returns user identity info (ID, email) in the token | Required by OAuth2 spec for authentication to work                                                                          |
| `offline_access`      | Allows token refresh without user re-login          | Without this, users would need to re-authenticate every hour when the access token expires                                  |
| `Sites.Read.All`      | Read SharePoint sites and their drives              | Populates the **Sites** and **Drives** dropdown selectors                                                                   |
| `Files.ReadWrite.All` | Read and write files in OneDrive/SharePoint         | **Read**: Download Excel files, list files in drives. **Write**: Upload modified Excel files after append/update operations |

### How Scopes Map to Node Features

```
User clicks Site dropdown
    → API: GET /sites?search=*
    → Requires: Sites.Read.All

User clicks Drive dropdown
    → API: GET /sites/{siteId}/drives
    → Requires: Sites.Read.All

User clicks File dropdown
    → API: GET /drives/{driveId}/root/children
    → Requires: Files.ReadWrite.All (Read portion)

Node reads Excel data
    → API: GET .../items/{fileId}/content
    → Requires: Files.ReadWrite.All (Read portion)

Node writes Excel data
    → API: PUT .../items/{fileId}/content
    → Requires: Files.ReadWrite.All (Write portion)
```

### Why not Files.Read.All + Files.Write.All separately?

Microsoft doesn't offer a separate `Files.Write.All`. The options are:

- `Files.Read.All` — read only
- `Files.ReadWrite.All` — read + write

Since the node needs to write (append rows, update cells), we need `ReadWrite`.

### Why These Scopes?

The node uses a download-modify-upload pattern:

1. **Site/Drive browsing** (`Sites.Read.All`): Populates dropdown selectors for sites and drives
2. **File operations** (`Files.ReadWrite.All`): Downloads Excel files for reading, uploads modified files for write operations

## Delegated Permissions

This node uses **delegated permissions** (OAuth2 authorization code flow), not application permissions.

### What This Means

```
Effective Access = OAuth Scopes ∩ User's SharePoint Permissions
```

The OAuth scopes define the _maximum_ permissions the app can request, but actual access is limited to what the authenticated user can access in SharePoint.

### Example

If a user only has access to:

- Marketing site
- Project-X site

Then even with `Sites.Read.All`, the node can only see those 2 sites. It cannot access HR, Finance, or other sites the user doesn't have SharePoint permissions for.

### Comparison

| Permission Type            | `Sites.Read.All` Means            |
| -------------------------- | --------------------------------- |
| **Delegated** (this node)  | Read sites the user has access to |
| **Application** (not used) | Read ALL sites in the tenant      |

## Enterprise Considerations

### Default Configuration is Enterprise-Safe

1. **Scoped to user access**: Delegated permissions respect SharePoint access controls
2. **Minimum required scopes**: Only 2 resource scopes (Sites.Read.All, Files.ReadWrite.All)
3. **No admin-level permissions**: Does not request Sites.ReadWrite.All or admin scopes

### Optional: Sites.Selected for Maximum Restriction

For organizations requiring pre-approved site access, administrators can modify the scope:

```
openid offline_access Sites.Selected Files.ReadWrite.All
```

**Trade-offs:**

| Aspect          | Sites.Read.All             | Sites.Selected                 |
| --------------- | -------------------------- | ------------------------------ |
| Site dropdown   | Works (shows user's sites) | Empty (requires manual siteId) |
| Admin setup     | None required              | Must grant app access per site |
| User experience | Self-service               | Admin-controlled               |

**Granting Sites.Selected access (PowerShell):**

```powershell
# Connect to SharePoint Online
Connect-SPOService -Url https://contoso-admin.sharepoint.com

# Grant app access to specific site
Grant-SPOSiteDesignRights -Identity <site-url> -Principals <app-id> -Rights Read
```

## API Endpoints Used

### For Dropdown Selectors

| Selector | Endpoint                              | Scope Required |
| -------- | ------------------------------------- | -------------- |
| Sites    | `GET /sites?search=*`                 | Sites.Read.All |
| Drives   | `GET /sites/{siteId}/drives`          | Sites.Read.All |
| Files    | `GET /drives/{driveId}/root/children` | Files.Read.All |
| Sheets   | Download file + parse with exceljs    | Files.Read.All |
| Tables   | `GET .../workbook/tables`             | Files.Read.All |

### For Operations

| Operation   | Endpoint                                 | Scope Required      |
| ----------- | ---------------------------------------- | ------------------- |
| Read rows   | `GET .../items/{fileId}/content`         | Files.Read.All      |
| Append rows | `GET` + `PUT .../items/{fileId}/content` | Files.ReadWrite.All |
| Update cell | `GET` + `PUT .../items/{fileId}/content` | Files.ReadWrite.All |
| Get sheets  | `GET .../items/{fileId}/content`         | Files.Read.All      |

## Security Best Practices

### For n8n Administrators

1. **Use service accounts**: Create dedicated accounts for n8n with appropriate SharePoint permissions
2. **Limit SharePoint access**: Control which sites/libraries the service account can access via SharePoint permissions
3. **Audit credential usage**: Monitor which workflows use the credential
4. **Rotate credentials**: Periodically disconnect and reconnect to refresh tokens

### For Workflow Builders

1. **Least privilege principle**: Use accounts with minimum required SharePoint access
2. **Avoid admin accounts**: Don't connect credentials with SharePoint admin access
3. **Validate inputs**: When using dynamic file/site IDs, validate they match expected patterns

## Credential Upgrade Path

When upgrading from an earlier version with broader scopes:

| Scenario                        | Action Required                         |
| ------------------------------- | --------------------------------------- |
| Existing credentials            | Continue working with original scopes   |
| New credentials                 | Automatically use reduced scopes        |
| Want reduced scopes on existing | Disconnect and reconnect the credential |

Reducing scopes does not break existing functionality since the removed scopes (`Sites.ReadWrite.All`, `Files.Read.All`) were never required.

## Comparison with Native n8n SharePoint Node

| Aspect          | This Node                           | Native n8n SharePoint                  |
| --------------- | ----------------------------------- | -------------------------------------- |
| Excel method    | Download/upload via Files API       | WAC (Web Application Companion) tokens |
| Scopes          | Files.ReadWrite.All, Sites.Read.All | Similar                                |
| Offline editing | Yes (full file download)            | No (requires active session)           |
| Large files     | Limited by memory                   | Streaming                              |

## References

- [Microsoft Graph Permissions Reference](https://learn.microsoft.com/en-us/graph/permissions-reference)
- [Delegated vs Application Permissions](https://learn.microsoft.com/en-us/azure/active-directory/develop/permissions-consent-overview)
- [Sites.Selected Permission](https://learn.microsoft.com/en-us/graph/api/site-get?view=graph-rest-1.0&tabs=http#permissions)
