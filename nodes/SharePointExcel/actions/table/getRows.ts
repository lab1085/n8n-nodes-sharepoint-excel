import type { IDataObject, IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import { graphRequest } from '../../api';
import type { GetTableRowsOptions, OperationContext, ResourceLocatorValue } from '../../types';

interface TableRow {
	index: number;
	values: (string | number | boolean | null)[][];
}

interface TableColumn {
	name: string;
	index: number;
}

export async function execute(
	this: IExecuteFunctions,
	_items: INodeExecutionData[],
	context: OperationContext,
): Promise<INodeExecutionData[]> {
	const returnData: INodeExecutionData[] = [];

	const tableNameParam = this.getNodeParameter('tableName', 0) as string | ResourceLocatorValue;
	const tableName = typeof tableNameParam === 'object' ? tableNameParam.value : tableNameParam;

	// Get returnAll and limit parameters
	const returnAll = this.getNodeParameter('returnAll', 0, true) as boolean;
	const limit = returnAll ? 0 : (this.getNodeParameter('limit', 0, 50) as number);

	const options = this.getNodeParameter('options', 0, {}) as GetTableRowsOptions;
	const rawData = options.rawData || false;
	const dataProperty = options.dataProperty || 'data';
	const fieldsFilter = options.fields
		? new Set(
				options.fields
					.split(',')
					.map((f) => f.trim())
					.filter(Boolean),
			)
		: new Set<string>();

	// Get table columns first to build headers
	const columnsEndpoint = `${context.basePath}/workbook/tables/${encodeURIComponent(tableName)}/columns`;
	const columnsResponse = await graphRequest.call(this, 'GET', columnsEndpoint);
	const columns = (columnsResponse as { value?: TableColumn[] }).value || [];
	const allHeaders = columns.map((c) => c.name);

	// Filter headers if fields specified
	const columnIndices: number[] = [];
	const outputHeaders: string[] = [];

	allHeaders.forEach((header, idx) => {
		if (fieldsFilter.size === 0 || fieldsFilter.has(header)) {
			columnIndices.push(idx);
			outputHeaders.push(header);
		}
	});

	// Get table rows via Graph API
	const rowsEndpoint = `${context.basePath}/workbook/tables/${encodeURIComponent(tableName)}/rows`;
	const rowsResponse = await graphRequest.call(this, 'GET', rowsEndpoint);
	const rows = (rowsResponse as { value?: TableRow[] }).value || [];

	// Apply limit
	const limitedRows = limit > 0 ? rows.slice(0, limit) : rows;

	if (rawData) {
		// RAW mode: return single item with arrays
		const rawRows: (string | number | boolean | null)[][] = [];

		for (const row of limitedRows) {
			const values = row.values[0] || [];
			const rowArray: (string | number | boolean | null)[] = [];
			let hasData = false;

			for (const idx of columnIndices) {
				const value = values[idx] ?? null;
				rowArray.push(value);
				if (value !== null && value !== undefined && value !== '') {
					hasData = true;
				}
			}

			if (hasData) {
				rawRows.push(rowArray);
			}
		}

		if (rawRows.length === 0) {
			returnData.push({ json: { message: 'No rows found in table' } });
		} else {
			returnData.push({
				json: {
					headers: outputHeaders,
					[dataProperty]: rawRows,
				},
			});
		}
	} else {
		// Normal mode: return keyed objects
		for (const row of limitedRows) {
			const rowData: IDataObject = {};
			const values = row.values[0] || [];
			let hasData = false;

			columnIndices.forEach((idx, i) => {
				const value = values[idx] ?? null;
				rowData[outputHeaders[i]] = value;
				if (value !== null && value !== undefined && value !== '') {
					hasData = true;
				}
			});

			if (hasData) {
				returnData.push({ json: rowData });
			}
		}

		if (returnData.length === 0) {
			returnData.push({
				json: { message: 'No rows found in table' } as IDataObject,
			});
		}
	}

	return returnData;
}
