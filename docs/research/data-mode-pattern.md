# n8n Data Mode Pattern Research

Research into how n8n implements sophisticated data input modes for row operations (append, upsert, update).

## The Three Modes

| Mode                               | Value     | Purpose                                                                   |
| ---------------------------------- | --------- | ------------------------------------------------------------------------- |
| **Auto-Map Input Data to Columns** | `autoMap` | Incoming JSON property names are matched directly to sheet column headers |
| **Map Each Column Below**          | `manual`  | User explicitly maps each input field to a destination column via UI      |
| **Raw**                            | `raw`     | Takes JSON data as-is, expects pre-formatted structure                    |

## Implementation Pattern

### 1. Mode Selection Property

Standard options property to select the data mode:

```typescript
{
  displayName: 'Data Mode',
  name: 'dataMode',
  type: 'options',
  noDataExpression: true,
  options: [
    {
      name: 'Auto-Map Input Data to Columns',
      value: 'autoMap',
      description: 'Use when node input properties match destination column names',
    },
    {
      name: 'Map Each Column Below',
      value: 'manual',
      description: 'Set the value for each destination column',
    },
    {
      name: 'Raw',
      value: 'raw',
      description: 'Send raw data as JSON',
    },
  ],
  default: 'autoMap',
}
```

### 2. Resource Mapper Property

The `resourceMapper` property type provides the drag-and-drop column mapping UI for "Map Each Column Below" mode:

```typescript
{
  displayName: 'Columns',
  name: 'columns',
  type: 'resourceMapper',
  noDataExpression: true,
  default: {
    mappingMode: 'defineBelow',
    value: null,
  },
  typeOptions: {
    resourceMapper: {
      resourceMapperMethod: 'getMappingColumns',
      mode: 'add',  // or 'update', 'upsert'
      fieldWords: {
        singular: 'column',
        plural: 'columns',
      },
      addAllFields: true,
      multiKeyMatch: false,
    },
  },
  displayOptions: {
    show: {
      dataMode: ['manual'],
    },
  },
}
```

### 3. getMappingColumns Method

Must be implemented in `methods.resourceMapping` to dynamically fetch columns:

```typescript
methods = {
	resourceMapping: {
		async getMappingColumns(this: ILoadOptionsFunctions): Promise<ResourceMapperFields> {
			// 1. Get sheet parameters
			const sheetName = this.getNodeParameter('sheetName', 0);

			// 2. Download/query the sheet to get headers
			const headers = await getSheetHeaders(this, sheetName);

			// 3. Return column metadata
			return {
				fields: headers.map((header, index) => ({
					id: header,
					displayName: header,
					type: 'string',
					required: false,
					defaultMatch: index === 0, // First column as default match key
					canBeUsedToMatch: true,
				})),
			};
		},
	},
};
```

### 4. Execution Code Handling

```typescript
async execute(this: IExecuteFunctions) {
  const dataMode = this.getNodeParameter('dataMode', itemIndex) as string;

  let rowData: IDataObject;

  switch (dataMode) {
    case 'autoMap':
      // Use input item directly - property names must match column names
      rowData = items[itemIndex].json;
      break;

    case 'manual':
      // Get mapped values from resourceMapper
      const mappingData = this.getNodeParameter('columns', itemIndex) as ResourceMapperValue;
      rowData = mappingData.value as IDataObject;
      break;

    case 'raw':
      // Parse raw JSON input
      const rawData = this.getNodeParameter('rawData', itemIndex) as string;
      rowData = JSON.parse(rawData);
      break;
  }

  // Proceed with rowData...
}
```

## ResourceMapperValue Type

The resourceMapper returns a structured object:

```typescript
interface ResourceMapperValue {
	mappingMode: 'autoMap' | 'defineBelow';
	schema: ResourceMapperField[];
	value: IDataObject | null;
}

interface ResourceMapperField {
	id: string;
	displayName: string;
	type: FieldType;
	required: boolean;
	defaultMatch: boolean;
	canBeUsedToMatch: boolean;
}
```

## Mode Comparison

| Aspect               | Auto-Map                      | Manual                                   | Raw            |
| -------------------- | ----------------------------- | ---------------------------------------- | -------------- |
| **UI Complexity**    | None                          | Full mapping interface                   | JSON editor    |
| **Configuration**    | Auto-detected                 | Per-column mapping                       | Direct JSON    |
| **Column Discovery** | Automatic at runtime          | Dynamic on setup                         | N/A            |
| **Use Case**         | Input already matches columns | Names don't match or need transformation | Advanced users |
| **Error Handling**   | Ignores extra fields          | User controls explicitly                 | Strict         |

## Examples in n8n Codebase

- **Google Sheets**: `packages/nodes-base/nodes/Google/Sheet/v2/methods/resourceMapping.ts`
- **Postgres**: `packages/nodes-base/nodes/Postgres/v2/methods/resourceMapping.ts`
- **Microsoft Excel**: `packages/nodes-base/nodes/Microsoft/Excel/v2/`

## Implementation Plan for SharePoint Excel

### Current State

- `appendRows` uses `rowData` as raw JSON string
- `upsertRows` uses `rowData` as raw JSON string + `keyColumn` option

### Target State

1. Add `dataMode` property with three options
2. Implement `getMappingColumns` in `methods.resourceMapping`:
   - Download workbook via ExcelJS
   - Extract header row from target sheet
   - Return column metadata
3. Add `resourceMapper` property for manual mode
4. Keep `rowData` JSON property for raw mode
5. Update `appendRows` and `upsertRows` handlers to support all three modes

### Considerations

- **Performance**: `getMappingColumns` will download the entire Excel file just to get headers
  - Could cache workbook during the same execution
  - Headers are typically in row 1 but should be configurable
- **Column Types**: ExcelJS can detect cell types, could map to n8n field types
- **Match Columns**: For upsert, need to identify which column(s) are the unique key

## References

- [n8n UI Elements Documentation](https://docs.n8n.io/integrations/creating-nodes/build/reference/ui-elements/)
- [n8n Data Mapping UI](https://docs.n8n.io/data/data-mapping/data-mapping-ui/)
- [n8n Standard Parameters Reference](https://docs.n8n.io/integrations/creating-nodes/build/reference/node-base-files/standard-parameters/)
