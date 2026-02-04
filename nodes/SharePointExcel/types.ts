import type { IDataObject, IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';

// Resource locator value from n8n UI
export interface ResourceLocatorValue {
	mode: string;
	value: string;
}

// Sources supported by this node
export type Source = 'sharepoint';

// Resources available
export type Resource = 'sheet' | 'table' | 'workbook';

// Sheet operations
export type SheetOperation =
	| 'getSheets'
	| 'readRows'
	| 'appendRows'
	| 'updateCell'
	| 'upsertRows'
	| 'clearSheet'
	| 'deleteSheet';

// Table operations
export type TableOperation = 'getTableRows' | 'getColumns' | 'lookup';

// Workbook operations
export type WorkbookOperation = 'getWorkbookSheets' | 'addSheet' | 'deleteWorkbook' | 'getWorkbooks';

// All operations union
export type Operation = SheetOperation | TableOperation | WorkbookOperation;

// Context passed to operation handlers
export interface OperationContext {
	source: Source;
	resource: Resource;
	operation: Operation;
	basePath: string;
	driveId: string;
	fileId: string;
	siteId?: string;
}

// Handler function signature for operations
export type OperationHandler = (
	this: IExecuteFunctions,
	items: INodeExecutionData[],
	context: OperationContext,
) => Promise<INodeExecutionData[]>;

// Graph API response types
export interface GraphSite {
	id: string;
	displayName: string;
	webUrl: string;
}

export interface GraphDrive {
	id: string;
	name: string;
}

export interface GraphDriveItem {
	id: string;
	name: string;
	file?: object;
}

export interface GraphError {
	error: {
		code?: string;
		message?: string;
	};
}

// Excel sheet info returned by getSheets
export interface SheetInfo {
	name: string;
	id: number;
	rowCount: number;
	columnCount: number;
}

// Options for readRows operation
export interface ReadRowsOptions {
	headerRow?: number;
	startRow?: number;
	rawData?: boolean;
	dataProperty?: string;
	fields?: string;
}

// Options for getTableRows operation
export interface GetTableRowsOptions {
	rawData?: boolean;
	dataProperty?: string;
	fields?: string;
}

// Options for upsertRows operation
export interface UpsertRowsOptions {
	headerRow?: number;
	keyColumn: string;
}

// Row data type
export type RowData = IDataObject;

// Data input modes for append/upsert operations
export type DataMode = 'autoMap' | 'manual' | 'raw';

// Resource mapper value from n8n UI
export interface ResourceMapperValue {
	mappingMode: 'defineBelow' | 'autoMapInputData';
	value: Record<string, unknown> | null;
	matchingColumns?: string[];
	schema?: ResourceMapperField[];
}

// Resource mapper field schema
export interface ResourceMapperField {
	id: string;
	displayName: string;
	required: boolean;
	defaultMatch: boolean;
	display: boolean;
	type: 'string' | 'number' | 'boolean' | 'dateTime';
	canBeUsedToMatch: boolean;
}
