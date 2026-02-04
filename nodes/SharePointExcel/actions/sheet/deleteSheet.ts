import type { IDataObject, IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import { loadWorkbook, saveWorkbook } from '../../api';
import type { OperationContext, ResourceLocatorValue } from '../../types';

export async function execute(
	this: IExecuteFunctions,
	_items: INodeExecutionData[],
	context: OperationContext,
): Promise<INodeExecutionData[]> {
	const sheetNameParam = this.getNodeParameter('sheetName', 0) as string | ResourceLocatorValue;
	const sheetName = typeof sheetNameParam === 'object' ? sheetNameParam.value : sheetNameParam;

	const workbook = await loadWorkbook.call(this, context.basePath);
	const worksheet = workbook.getWorksheet(sheetName);

	if (!worksheet) {
		throw new NodeOperationError(this.getNode(), `Sheet "${sheetName}" not found in workbook`);
	}

	// Check if this is the last sheet
	if (workbook.worksheets.length <= 1) {
		throw new NodeOperationError(this.getNode(), 'Cannot delete the last sheet in a workbook');
	}

	// Remove the worksheet
	workbook.removeWorksheet(worksheet.id);

	// Upload back
	await saveWorkbook.call(this, context.basePath, workbook);

	return [
		{
			json: {
				success: true,
				deletedSheet: sheetName,
				remainingSheets: workbook.worksheets.map((ws) => ws.name),
			} as IDataObject,
		},
	];
}
