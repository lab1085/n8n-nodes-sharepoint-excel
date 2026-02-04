import type { IDataObject, IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
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

	const tableNameParam = this.getNodeParameter('tableName', 0) as
		| string
		| ResourceLocatorValue;
	const tableName =
		typeof tableNameParam === 'object' ? tableNameParam.value : tableNameParam;

	// Get table columns first to build headers
	const columnsEndpoint = `${context.basePath}/workbook/tables/${encodeURIComponent(tableName)}/columns`;
	const columnsResponse = await graphRequest.call(this, 'GET', columnsEndpoint);
	const columns = (columnsResponse as { value?: TableColumn[] }).value || [];
	const headers = columns.map((c) => c.name);

	// Get table rows via Graph API
	const rowsEndpoint = `${context.basePath}/workbook/tables/${encodeURIComponent(tableName)}/rows`;
	const rowsResponse = await graphRequest.call(this, 'GET', rowsEndpoint);

	const rows = (rowsResponse as { value?: TableRow[] }).value || [];

	for (const row of rows) {
		const rowData: IDataObject = {};
		const values = row.values[0] || [];

		headers.forEach((header, idx) => {
			rowData[header] = values[idx] ?? null;
		});

		returnData.push({ json: rowData });
	}

	if (returnData.length === 0) {
		returnData.push({
			json: { message: 'No rows found in table' } as IDataObject,
		});
	}

	return returnData;
}
