import type { IDataObject, IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import { graphRequest } from '../../api';
import type { OperationContext, ResourceLocatorValue } from '../../types';

interface TableColumn {
	id: string;
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

	// Get table columns via Graph API
	const endpoint = `${context.basePath}/workbook/tables/${encodeURIComponent(tableName)}/columns`;
	const response = await graphRequest.call(this, 'GET', endpoint);

	const columns = (response as { value?: TableColumn[] }).value || [];

	for (const column of columns) {
		returnData.push({
			json: {
				id: column.id,
				name: column.name,
				index: column.index,
			} as IDataObject,
		});
	}

	if (returnData.length === 0) {
		returnData.push({
			json: { message: 'No columns found in table' } as IDataObject,
		});
	}

	return returnData;
}
