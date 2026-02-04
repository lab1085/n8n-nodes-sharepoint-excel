import type { IDataObject, IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import { graphRequest } from '../../api';
import type { OperationContext, ResourceLocatorValue } from '../../types';

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
	const lookupColumn = this.getNodeParameter('lookupColumn', 0) as string;
	const lookupValue = this.getNodeParameter('lookupValue', 0) as string;

	// Get table columns first to build headers and find lookup column index
	const columnsEndpoint = `${context.basePath}/workbook/tables/${encodeURIComponent(tableName)}/columns`;
	const columnsResponse = await graphRequest.call(this, 'GET', columnsEndpoint);
	const columns = (columnsResponse as { value?: TableColumn[] }).value || [];
	const headers = columns.map((c) => c.name);

	const lookupColIndex = headers.indexOf(lookupColumn);
	if (lookupColIndex === -1) {
		throw new NodeOperationError(
			this.getNode(),
			`Column "${lookupColumn}" not found in table "${tableName}"`,
		);
	}

	// Get table rows via Graph API
	const rowsEndpoint = `${context.basePath}/workbook/tables/${encodeURIComponent(tableName)}/rows`;
	const rowsResponse = await graphRequest.call(this, 'GET', rowsEndpoint);

	const rows = (rowsResponse as { value?: TableRow[] }).value || [];

	// Find matching rows
	for (const row of rows) {
		const values = row.values[0] || [];
		const cellValue = values[lookupColIndex];

		// Compare as strings for flexibility
		if (String(cellValue) === String(lookupValue)) {
			const rowData: IDataObject = {};
			headers.forEach((header, idx) => {
				rowData[header] = values[idx] ?? null;
			});
			returnData.push({ json: rowData });
		}
	}

	if (returnData.length === 0) {
		returnData.push({
			json: {
				message: `No rows found where "${lookupColumn}" equals "${lookupValue}"`,
			} as IDataObject,
		});
	}

	return returnData;
}
