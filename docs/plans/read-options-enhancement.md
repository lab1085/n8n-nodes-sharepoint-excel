# Plan: Add Read Options to readRows and getTableRows

## Overview

Enhance `readRows` (sheet) and `getTableRows` (table) operations with options matching the native n8n Excel node pattern:

| Option | Description |
|--------|-------------|
| **Return All** | Toggle: when ON returns all rows, when OFF shows Limit field |
| **Limit** | Max rows to return (shown when Return All is OFF) |
| **RAW Data** | Toggle: return arrays instead of keyed objects |
| **Data Property** | Property name for raw data output (shown when RAW Data is ON) |
| **Filters > Fields** | Select specific columns to return |

## Current State

### readRows (sheet)
- Has `options.headerRow`, `options.startRow`, `options.maxRows`
- Returns keyed objects: `[{Name: "John", Age: 30}, ...]`
- No field filtering

### getTableRows (table)
- No pagination/limit options
- Returns keyed objects
- No field filtering

## Files to Modify

1. `nodes/SharePointExcel/descriptions.ts` - Add new properties
2. `nodes/SharePointExcel/types.ts` - Add option interfaces
3. `nodes/SharePointExcel/actions/sheet/readRows.ts` - Implement options
4. `nodes/SharePointExcel/actions/table/getRows.ts` - Implement options

## Implementation Steps

### Step 1: types.ts

Add/update interfaces:

```typescript
export interface ReadRowsOptions {
  headerRow?: number;
  startRow?: number;
  maxRows?: number;       // Keep for backward compatibility
  returnAll?: boolean;    // New
  limit?: number;         // New (alternative to maxRows)
  rawData?: boolean;      // New
  dataProperty?: string;  // New
  fields?: string[];      // New
}

export interface GetTableRowsOptions {
  returnAll?: boolean;
  limit?: number;
  rawData?: boolean;
  dataProperty?: string;
  fields?: string[];
}
```

### Step 2: descriptions.ts

Add new properties for readRows and getTableRows:

#### Return All Toggle
```typescript
export const returnAllProperty: INodeProperties = {
  displayName: 'Return All',
  name: 'returnAll',
  type: 'boolean',
  default: true,
  description: 'Whether to return all results or limit the number',
  displayOptions: {
    show: {
      operation: ['readRows', 'getTableRows'],
    },
  },
};
```

#### Limit (shown when Return All is OFF)
```typescript
export const limitProperty: INodeProperties = {
  displayName: 'Limit',
  name: 'limit',
  type: 'number',
  default: 100,
  description: 'Max number of rows to return',
  typeOptions: {
    minValue: 1,
  },
  displayOptions: {
    show: {
      operation: ['readRows', 'getTableRows'],
      returnAll: [false],
    },
  },
};
```

#### RAW Data Toggle (in Options collection)
```typescript
{
  displayName: 'RAW Data',
  name: 'rawData',
  type: 'boolean',
  default: false,
  description: 'Whether to return data as arrays instead of keyed objects',
}
```

#### Data Property (shown when RAW Data is ON)
```typescript
{
  displayName: 'Data Property',
  name: 'dataProperty',
  type: 'string',
  default: 'data',
  description: 'Property name to use for the raw data output',
  displayOptions: {
    show: {
      '/options.rawData': [true],
    },
  },
}
```

#### Fields Filter (in Options collection)
```typescript
{
  displayName: 'Fields',
  name: 'fields',
  type: 'string',
  default: '',
  placeholder: 'Name, Email, Status',
  description: 'Comma-separated list of column names to return (empty = all)',
}
```

#### Update readRowsOptions
Move `maxRows` to be hidden (backward compat) and add new options.

#### Create getTableRowsOptions
New options collection for table operations.

### Step 3: readRows.ts

Refactor to handle new options:

```typescript
export async function execute(...): Promise<INodeExecutionData[]> {
  // Get parameters
  const returnAll = this.getNodeParameter('returnAll', 0, true) as boolean;
  const limit = returnAll ? 0 : this.getNodeParameter('limit', 0, 100) as number;
  const options = this.getNodeParameter('options', 0, {}) as ReadRowsOptions;

  // Backward compat: use maxRows if limit not set
  const maxRows = limit || options.maxRows || 0;

  const rawData = options.rawData || false;
  const dataProperty = options.dataProperty || 'data';
  const fieldsFilter = options.fields
    ? options.fields.split(',').map(f => f.trim()).filter(Boolean)
    : [];

  // ... load workbook, get headers ...

  // Filter headers if fields specified
  const outputHeaders = fieldsFilter.length > 0
    ? headers.filter(h => fieldsFilter.includes(h))
    : headers;

  // Read rows
  if (rawData) {
    // RAW mode: return array of arrays
    const rawRows: (string | number | boolean | null)[][] = [];
    // ... iterate rows, push arrays ...
    return [{ json: { [dataProperty]: rawRows, headers: outputHeaders } }];
  } else {
    // Normal mode: return keyed objects
    // ... current logic with field filtering ...
  }
}
```

### Step 4: getRows.ts (table)

Similar refactor:

```typescript
export async function execute(...): Promise<INodeExecutionData[]> {
  const returnAll = this.getNodeParameter('returnAll', 0, true) as boolean;
  const limit = returnAll ? 0 : this.getNodeParameter('limit', 0, 100) as number;
  const options = this.getNodeParameter('options', 0, {}) as GetTableRowsOptions;

  const rawData = options.rawData || false;
  const dataProperty = options.dataProperty || 'data';
  const fieldsFilter = options.fields
    ? options.fields.split(',').map(f => f.trim()).filter(Boolean)
    : [];

  // ... get columns, get rows ...

  // Apply limit
  const limitedRows = limit > 0 ? rows.slice(0, limit) : rows;

  // Filter columns
  const columnIndices = fieldsFilter.length > 0
    ? headers.map((h, i) => fieldsFilter.includes(h) ? i : -1).filter(i => i >= 0)
    : headers.map((_, i) => i);

  if (rawData) {
    // RAW mode
    const rawRows = limitedRows.map(row =>
      columnIndices.map(i => row.values[0][i])
    );
    return [{ json: { [dataProperty]: rawRows, headers: columnIndices.map(i => headers[i]) } }];
  } else {
    // Normal mode with filtering
    // ...
  }
}
```

### Step 5: Update properties array

Add new properties to the exported `properties` array in correct order.

## Output Formats

### Normal Mode (rawData: false)
```json
[
  { "json": { "Name": "John", "Email": "john@example.com" } },
  { "json": { "Name": "Jane", "Email": "jane@example.com" } }
]
```

### RAW Mode (rawData: true, dataProperty: "data")
```json
[
  {
    "json": {
      "headers": ["Name", "Email"],
      "data": [
        ["John", "john@example.com"],
        ["Jane", "jane@example.com"]
      ]
    }
  }
]
```

## Backward Compatibility

- `maxRows` option kept but hidden, still works if set
- Default `returnAll: true` matches current behavior (return all rows)
- Default `rawData: false` keeps current keyed object format
- Empty `fields` filter returns all columns (current behavior)

## Verification

1. Build: `bun run build`
2. Lint: `bun run lint`
3. Manual testing:
   - Test readRows with Return All ON/OFF
   - Test readRows with RAW Data ON/OFF
   - Test readRows with Fields filter
   - Test getTableRows with same options
   - Verify backward compatibility (workflows using maxRows still work)
