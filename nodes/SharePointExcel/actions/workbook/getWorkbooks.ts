import type { IDataObject, IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import { graphRequest } from '../../api';
import type { GraphDriveItem, OperationContext } from '../../types';

export async function execute(
	this: IExecuteFunctions,
	_items: INodeExecutionData[],
	context: OperationContext,
): Promise<INodeExecutionData[]> {
	const returnData: INodeExecutionData[] = [];

	// List files in the drive root
	const endpoint = `/drives/${context.driveId}/root/children`;
	const response = await graphRequest.call(this, 'GET', endpoint);

	const items = (response as { value?: GraphDriveItem[] }).value || [];

	// Filter to only Excel files
	const excelFiles = items.filter(
		(item) => item.file && item.name.toLowerCase().endsWith('.xlsx'),
	);

	for (const file of excelFiles) {
		returnData.push({
			json: {
				id: file.id,
				name: file.name,
			} as IDataObject,
		});
	}

	if (returnData.length === 0) {
		returnData.push({
			json: { message: 'No Excel files found in drive' } as IDataObject,
		});
	}

	return returnData;
}
