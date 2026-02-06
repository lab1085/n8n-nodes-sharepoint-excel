# n8n-nodes-sharepoint-excel

This is an n8n community node. It lets you use **Microsoft SharePoint Excel** files in your n8n workflows.

This node provides read and write operations for Excel files stored in SharePoint document libraries via Microsoft Graph API. Unlike the native n8n Microsoft Excel node that uses WAC (Web Application Companion) tokens, this node downloads and uploads the entire Excel file using the `exceljs` library, **bypassing WAC token limitations** that often cause issues with SharePoint-hosted files.

[n8n](https://n8n.io/) is a [fair-code licensed](https://docs.n8n.io/reference/license/) workflow automation platform.

[Installation](#installation)
[Operations](#operations)
[Credentials](#credentials)
[Compatibility](#compatibility)
[Usage](#usage)
[Documentation](#documentation)
[Resources](#resources)

## Installation

Follow the [installation guide](https://docs.n8n.io/integrations/community-nodes/installation/) in the n8n community nodes documentation.

## Operations

### Sheet

| Operation       | Description                                                       |
| --------------- | ----------------------------------------------------------------- |
| **Get Sheets**  | List all worksheets in the workbook                               |
| **Get Rows**    | Read rows from a sheet with configurable header row and start row |
| **Append Rows** | Add new rows to the end of a sheet                                |
| **Update**      | Update a specific cell by reference (e.g., A1, B5)                |
| **Upsert Rows** | Insert or update rows based on a key column                       |
| **Clear**       | Clear all data from a sheet                                       |
| **Delete**      | Delete a sheet from the workbook                                  |

### Table

| Operation       | Description                           |
| --------------- | ------------------------------------- |
| **Get Rows**    | Retrieve all rows from an Excel table |
| **Get Columns** | Get column definitions from a table   |
| **Lookup**      | Find a row by column value            |

### Workbook

| Operation         | Description                        |
| ----------------- | ---------------------------------- |
| **Get Sheets**    | List all sheets in the workbook    |
| **Add Sheet**     | Create a new sheet in the workbook |
| **Delete**        | Delete the workbook file           |
| **Get Workbooks** | List all Excel files in the drive  |

## Credentials

This node requires **Microsoft Graph OAuth2** credentials.

### Prerequisites

1. An Azure account with access to [Azure Portal](https://portal.azure.com)
2. A Microsoft 365 account with SharePoint access

### Setup

1. **Register an application in Azure AD:**
   - Go to Azure Portal > Azure Active Directory > App registrations
   - Click "New registration"
   - Name your application (e.g., "n8n SharePoint Excel")
   - Set the redirect URI to your n8n OAuth callback URL

2. **Configure API permissions:**
   - Add the following Microsoft Graph **delegated** permissions:
     - `Sites.Read.All` - Read sites (for browsing sites and drives)
     - `Files.ReadWrite.All` - Read and write files
   - Grant admin consent for these permissions
   - See [Security and Permissions](docs/security-and-permissions.md) for details

3. **Create a client secret:**
   - Go to "Certificates & secrets"
   - Create a new client secret and copy the value

4. **Configure in n8n:**
   - Add new credentials of type "Microsoft Graph OAuth2 API"
   - Enter your Client ID and Client Secret
   - Complete the OAuth2 authorization flow

## Compatibility

- Tested with n8n version 1.x and above
- Requires Microsoft Graph API v1.0
- Works with `.xlsx` files only

## Usage

### Why use this node instead of the native Microsoft Excel node?

The native n8n Microsoft Excel 365 node uses WAC (Web Application Companion) tokens to interact with Excel files. This works well for OneDrive but often fails with SharePoint-hosted files due to:

- Token expiration issues
- Permission complexities with SharePoint sites
- WAC service availability problems

This node takes a different approach: it **downloads the entire Excel file**, modifies it locally using the `exceljs` library, and **uploads it back**. This bypasses all WAC-related issues.

### Finding Required IDs

The node provides searchable dropdowns to select:

- **Site** - Your SharePoint site
- **Drive** - The document library containing the file
- **File** - The Excel file to operate on
- **Sheet/Table** - The specific sheet or table within the file

You can also enter IDs manually:

- **Site ID** format: `contoso.sharepoint.com,site-guid,web-guid`
- **Drive ID** format: `b!xxxxx...`
- **File ID** format: Item ID from SharePoint

### Reading Data

When reading rows, you can configure:

- **Header Row** - Which row contains column headers (default: 1)
- **Start Row** - First data row to read (default: 2)
- **Max Rows** - Limit the number of rows returned (0 = all)

### Writing Data

When appending or upserting rows:

- Provide data as a JSON object with column headers as keys
- For multiple rows, provide an array of objects
- The node automatically matches columns to existing headers

## Documentation

| Document                                                     | Description                                                    |
| ------------------------------------------------------------ | -------------------------------------------------------------- |
| [Security and Permissions](docs/security-and-permissions.md) | OAuth scopes, delegated permissions, enterprise considerations |
| [File Locking Behavior](docs/file-locking-behavior.md)       | Why files get locked and how to handle it                      |

## Resources

- [n8n community nodes documentation](https://docs.n8n.io/integrations/#community-nodes)
- [Microsoft Graph API documentation](https://learn.microsoft.com/en-us/graph/overview)
- [SharePoint REST API reference](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service)

## License

[MIT](LICENSE.md)
