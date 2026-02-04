import type { IDataObject, IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import { deleteFile } from '../../api';
import type { OperationContext } from '../../types';

export async function execute(
	this: IExecuteFunctions,
	_items: INodeExecutionData[],
	context: OperationContext,
): Promise<INodeExecutionData[]> {
	// Delete the file via Graph API
	await deleteFile.call(this, context.basePath);

	return [
		{
			json: {
				success: true,
				deleted: true,
				fileId: context.fileId,
			} as IDataObject,
		},
	];
}
