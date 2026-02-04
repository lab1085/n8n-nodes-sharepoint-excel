import type { IDataObject, IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import { loadWorkbook, saveWorkbook } from '../../api';
import type { OperationContext } from '../../types';

export async function execute(
	this: IExecuteFunctions,
	_items: INodeExecutionData[],
	context: OperationContext,
): Promise<INodeExecutionData[]> {
	const newSheetName = this.getNodeParameter('newSheetName', 0) as string;

	if (!newSheetName || newSheetName.trim() === '') {
		throw new NodeOperationError(this.getNode(), 'Sheet name cannot be empty');
	}

	const workbook = await loadWorkbook.call(this, context.basePath);

	// Check if sheet name already exists
	const existingSheet = workbook.getWorksheet(newSheetName);
	if (existingSheet) {
		throw new NodeOperationError(
			this.getNode(),
			`Sheet "${newSheetName}" already exists in the workbook`,
		);
	}

	// Add new worksheet
	const newSheet = workbook.addWorksheet(newSheetName);

	// Upload back
	await saveWorkbook.call(this, context.basePath, workbook);

	return [
		{
			json: {
				success: true,
				sheetName: newSheetName,
				sheetId: newSheet.id,
				totalSheets: workbook.worksheets.length,
			} as IDataObject,
		},
	];
}
