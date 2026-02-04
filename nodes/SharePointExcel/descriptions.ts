import type { INodeProperties } from 'n8n-workflow';

// Source selector
export const sourceProperty: INodeProperties = {
	displayName: 'Source',
	name: 'source',
	type: 'options',
	noDataExpression: true,
	options: [
		{
			name: 'SharePoint',
			value: 'sharepoint',
			description: 'Excel file stored in a SharePoint document library',
		},
		{
			name: 'OneDrive',
			value: 'onedrive',
			description: 'Excel file stored in OneDrive',
		},
	],
	default: 'sharepoint',
};

// Resource selector
export const resourceProperty: INodeProperties = {
	displayName: 'Resource',
	name: 'resource',
	type: 'options',
	noDataExpression: true,
	options: [
		{
			name: 'Table',
			value: 'table',
			description: 'Represents an Excel table',
		},
		{
			name: 'Workbook',
			value: 'workbook',
			description: 'A workbook is the top level object which contains one or more worksheets',
		},
		{
			name: 'Sheet',
			value: 'sheet',
			description: 'A sheet is a grid of cells which can contain data, tables, charts, etc',
		},
	],
	default: 'sheet',
};

// Table Operations
export const tableOperations: INodeProperties = {
	displayName: 'Operation',
	name: 'operation',
	type: 'options',
	noDataExpression: true,
	displayOptions: {
		show: {
			resource: ['table'],
		},
	},
	options: [
		{
			name: 'Get Columns',
			value: 'getColumns',
			description: 'Get column definitions from a table',
			action: 'Get columns from table',
		},
		{
			name: 'Get Rows',
			value: 'getTableRows',
			description: 'Retrieve rows from a table',
			action: 'Get rows from table',
		},
		{
			name: 'Lookup',
			value: 'lookup',
			description: 'Find a row by column value',
			action: 'Lookup row in table',
		},
	],
	default: 'getTableRows',
};

// Workbook Operations
export const workbookOperations: INodeProperties = {
	displayName: 'Operation',
	name: 'operation',
	type: 'options',
	noDataExpression: true,
	displayOptions: {
		show: {
			resource: ['workbook'],
		},
	},
	options: [
		{
			name: 'Add Sheet',
			value: 'addSheet',
			description: 'Add a new sheet to the workbook',
			action: 'Add sheet to workbook',
		},
		{
			name: 'Delete',
			value: 'deleteWorkbook',
			description: 'Delete the workbook file',
			action: 'Delete workbook',
		},
		{
			name: 'Get Sheets',
			value: 'getWorkbookSheets',
			description: 'Get list of sheets in the workbook',
			action: 'Get sheets from workbook',
		},
		{
			name: 'Get Workbooks',
			value: 'getWorkbooks',
			description: 'List Excel files in the drive',
			action: 'Get workbooks from drive',
		},
	],
	default: 'getWorkbookSheets',
};

// Sheet Operations
export const sheetOperations: INodeProperties = {
	displayName: 'Operation',
	name: 'operation',
	type: 'options',
	noDataExpression: true,
	displayOptions: {
		show: {
			resource: ['sheet'],
		},
	},
	options: [
		{
			name: 'Append Rows',
			value: 'appendRows',
			description: 'Append rows to the end of a sheet',
			action: 'Append rows to sheet',
		},
		{
			name: 'Clear',
			value: 'clearSheet',
			description: 'Clear all data from a sheet',
			action: 'Clear sheet',
		},
		{
			name: 'Delete',
			value: 'deleteSheet',
			description: 'Delete a sheet from the workbook',
			action: 'Delete sheet',
		},
		{
			name: 'Get Rows',
			value: 'readRows',
			description: 'Retrieve rows from a sheet',
			action: 'Get rows from sheet',
		},
		{
			name: 'Get Sheets',
			value: 'getSheets',
			description: 'Get list of sheets in the workbook',
			action: 'Get sheets',
		},
		{
			name: 'Update',
			value: 'updateCell',
			description: 'Update a cell in a sheet',
			action: 'Update cell in sheet',
		},
		{
			name: 'Upsert Rows',
			value: 'upsertRows',
			description: 'Append or update rows based on a key column',
			action: 'Upsert rows in sheet',
		},
	],
	default: 'readRows',
};

// Site ID - SharePoint only (resourceLocator with search)
export const siteIdProperty: INodeProperties = {
	displayName: 'Site',
	name: 'siteId',
	type: 'resourceLocator',
	required: true,
	default: { mode: 'list', value: '' },
	description: 'The SharePoint site',
	displayOptions: {
		show: {
			source: ['sharepoint'],
		},
	},
	modes: [
		{
			displayName: 'From List',
			name: 'list',
			type: 'list',
			placeholder: 'Search for a site...',
			typeOptions: {
				searchListMethod: 'searchSites',
				searchable: true,
			},
		},
		{
			displayName: 'By ID',
			name: 'id',
			type: 'string',
			placeholder: 'contoso.sharepoint.com,site-guid,web-guid',
			validation: [
				{
					type: 'regex',
					properties: {
						regex: '.+',
						errorMessage: 'Site ID cannot be empty',
					},
				},
			],
		},
	],
};

// Drive ID - resourceLocator with list and manual ID
export const driveIdProperty: INodeProperties = {
	displayName: 'Drive',
	name: 'driveId',
	type: 'resourceLocator',
	required: true,
	default: { mode: 'list', value: '' },
	description: 'The document library or drive containing the file',
	modes: [
		{
			displayName: 'From List',
			name: 'list',
			type: 'list',
			placeholder: 'Select a drive...',
			typeOptions: {
				searchListMethod: 'getDrives',
				searchable: false,
			},
		},
		{
			displayName: 'By ID',
			name: 'id',
			type: 'string',
			placeholder: 'b!xxxxx',
			validation: [
				{
					type: 'regex',
					properties: {
						regex: '.+',
						errorMessage: 'Drive ID cannot be empty',
					},
				},
			],
		},
	],
};

// File ID - resourceLocator with list and manual ID
export const fileIdProperty: INodeProperties = {
	displayName: 'File',
	name: 'fileId',
	type: 'resourceLocator',
	required: true,
	default: { mode: 'list', value: '' },
	description: 'The Excel file (.xlsx) to operate on',
	displayOptions: {
		hide: {
			operation: ['getWorkbooks'],
		},
	},
	modes: [
		{
			displayName: 'From List',
			name: 'list',
			type: 'list',
			placeholder: 'Select a file...',
			typeOptions: {
				searchListMethod: 'getFiles',
				searchable: false,
			},
		},
		{
			displayName: 'By ID',
			name: 'id',
			type: 'string',
			placeholder: '01ABCDEF...',
			validation: [
				{
					type: 'regex',
					properties: {
						regex: '.+',
						errorMessage: 'File ID cannot be empty',
					},
				},
			],
		},
	],
};

// Sheet name - resourceLocator with list and manual input
export const sheetNameProperty: INodeProperties = {
	displayName: 'Sheet',
	name: 'sheetName',
	type: 'resourceLocator',
	required: true,
	default: { mode: 'list', value: '' },
	description: 'The worksheet to operate on',
	modes: [
		{
			displayName: 'From List',
			name: 'list',
			type: 'list',
			placeholder: 'Select a sheet...',
			typeOptions: {
				searchListMethod: 'getSheets',
				searchable: false,
			},
		},
		{
			displayName: 'By Name',
			name: 'name',
			type: 'string',
			placeholder: 'Sheet1',
			validation: [
				{
					type: 'regex',
					properties: {
						regex: '.+',
						errorMessage: 'Sheet name cannot be empty',
					},
				},
			],
		},
	],
	displayOptions: {
		show: {
			resource: ['sheet'],
		},
		hide: {
			operation: ['getSheets'],
		},
	},
};

// Table name - resourceLocator for table operations
export const tableNameProperty: INodeProperties = {
	displayName: 'Table',
	name: 'tableName',
	type: 'resourceLocator',
	required: true,
	default: { mode: 'list', value: '' },
	description: 'The table to operate on',
	modes: [
		{
			displayName: 'From List',
			name: 'list',
			type: 'list',
			placeholder: 'Select a table...',
			typeOptions: {
				searchListMethod: 'getTables',
				searchable: false,
			},
		},
		{
			displayName: 'By Name',
			name: 'name',
			type: 'string',
			placeholder: 'Table1',
			validation: [
				{
					type: 'regex',
					properties: {
						regex: '.+',
						errorMessage: 'Table name cannot be empty',
					},
				},
			],
		},
	],
	displayOptions: {
		show: {
			resource: ['table'],
		},
	},
};

// Options for readRows
export const readRowsOptions: INodeProperties = {
	displayName: 'Options',
	name: 'options',
	type: 'collection',
	placeholder: 'Add Option',
	default: {},
	displayOptions: {
		show: {
			operation: ['readRows'],
		},
	},
	options: [
		{
			displayName: 'Header Row',
			name: 'headerRow',
			type: 'number',
			default: 1,
			description: 'Row number containing headers (1-indexed)',
		},
		{
			displayName: 'Start Row',
			name: 'startRow',
			type: 'number',
			default: 2,
			description: 'First data row to read (1-indexed)',
		},
		{
			displayName: 'Max Rows',
			name: 'maxRows',
			type: 'number',
			default: 0,
			description: 'Maximum rows to return (0 = all)',
		},
	],
};

// Row data for appendRows
export const rowDataProperty: INodeProperties = {
	displayName: 'Row Data',
	name: 'rowData',
	type: 'json',
	required: true,
	default: '{}',
	description:
		'JSON object with column headers as keys, or array of objects for multiple rows',
	displayOptions: {
		show: {
			operation: ['appendRows'],
		},
	},
};

// Cell reference for updateCell
export const cellRefProperty: INodeProperties = {
	displayName: 'Cell Reference',
	name: 'cellRef',
	type: 'string',
	required: true,
	default: 'A1',
	placeholder: 'A1',
	description: 'Cell to update (e.g., A1, B5, C10)',
	displayOptions: {
		show: {
			operation: ['updateCell'],
		},
	},
};

// Cell value for updateCell
export const cellValueProperty: INodeProperties = {
	displayName: 'Value',
	name: 'cellValue',
	type: 'string',
	required: true,
	default: '',
	description: 'New value for the cell',
	displayOptions: {
		show: {
			operation: ['updateCell'],
		},
	},
};

// Row data for upsertRows
export const upsertRowDataProperty: INodeProperties = {
	displayName: 'Row Data',
	name: 'rowData',
	type: 'json',
	required: true,
	default: '{}',
	description:
		'JSON object with column headers as keys, or array of objects for multiple rows',
	displayOptions: {
		show: {
			operation: ['upsertRows'],
		},
	},
};

// Options for upsertRows
export const upsertRowsOptions: INodeProperties = {
	displayName: 'Options',
	name: 'options',
	type: 'collection',
	placeholder: 'Add Option',
	default: {},
	displayOptions: {
		show: {
			operation: ['upsertRows'],
		},
	},
	options: [
		{
			displayName: 'Key Column',
			name: 'keyColumn',
			type: 'string',
			default: '',
			description: 'Column name to use as the unique key for matching rows',
		},
		{
			displayName: 'Header Row',
			name: 'headerRow',
			type: 'number',
			default: 1,
			description: 'Row number containing headers (1-indexed)',
		},
	],
};

// New sheet name for addSheet
export const newSheetNameProperty: INodeProperties = {
	displayName: 'Sheet Name',
	name: 'newSheetName',
	type: 'string',
	required: true,
	default: '',
	placeholder: 'NewSheet',
	description: 'Name for the new sheet',
	displayOptions: {
		show: {
			operation: ['addSheet'],
		},
	},
};

// Lookup column for table lookup
export const lookupColumnProperty: INodeProperties = {
	displayName: 'Lookup Column',
	name: 'lookupColumn',
	type: 'string',
	required: true,
	default: '',
	description: 'Column name to search in',
	displayOptions: {
		show: {
			operation: ['lookup'],
		},
	},
};

// Lookup value for table lookup
export const lookupValueProperty: INodeProperties = {
	displayName: 'Lookup Value',
	name: 'lookupValue',
	type: 'string',
	required: true,
	default: '',
	description: 'Value to search for',
	displayOptions: {
		show: {
			operation: ['lookup'],
		},
	},
};

// All properties in order
export const properties: INodeProperties[] = [
	sourceProperty,
	resourceProperty,
	tableOperations,
	workbookOperations,
	sheetOperations,
	siteIdProperty,
	driveIdProperty,
	fileIdProperty,
	sheetNameProperty,
	tableNameProperty,
	readRowsOptions,
	rowDataProperty,
	upsertRowDataProperty,
	upsertRowsOptions,
	cellRefProperty,
	cellValueProperty,
	newSheetNameProperty,
	lookupColumnProperty,
	lookupValueProperty,
];
