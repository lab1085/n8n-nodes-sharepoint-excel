import type { IDataObject, IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import { getWorksheet, loadWorkbook, saveWorkbook } from '../../api';
import type { OperationContext, ResourceLocatorValue } from '../../types';

export async function execute(
	this: IExecuteFunctions,
	_items: INodeExecutionData[],
	context: OperationContext,
): Promise<INodeExecutionData[]> {
	const sheetNameParam = this.getNodeParameter('sheetName', 0) as string | ResourceLocatorValue;
	const sheetName = typeof sheetNameParam === 'object' ? sheetNameParam.value : sheetNameParam;

	const workbook = await loadWorkbook.call(this, context.basePath);
	const worksheet = getWorksheet(workbook, sheetName, this);

	// Get counts before clearing
	const rowCount = worksheet.rowCount;
	const columnCount = worksheet.columnCount;

	// Clear all rows in one call (keeping the worksheet)
	if (rowCount > 0) {
		worksheet.spliceRows(1, rowCount);
	}

	// Upload back
	await saveWorkbook.call(this, context.basePath, workbook);

	return [
		{
			json: {
				success: true,
				sheet: sheetName,
				clearedRows: rowCount,
				clearedColumns: columnCount,
			} as IDataObject,
		},
	];
}
