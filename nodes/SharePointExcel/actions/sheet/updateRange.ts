import type { IDataObject, IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import { getWorksheet, loadWorkbook, saveWorkbook } from '../../api';
import type { OperationContext, ResourceLocatorValue } from '../../types';

/**
 * Update a cell (or range) in a sheet
 * Operation value: 'updateCell' (kept for backward compatibility)
 */
export async function execute(
	this: IExecuteFunctions,
	_items: INodeExecutionData[],
	context: OperationContext,
): Promise<INodeExecutionData[]> {
	const sheetNameParam = this.getNodeParameter('sheetName', 0) as
		| string
		| ResourceLocatorValue;
	const sheetName =
		typeof sheetNameParam === 'object' ? sheetNameParam.value : sheetNameParam;
	const cellRef = this.getNodeParameter('cellRef', 0) as string;
	const cellValue = this.getNodeParameter('cellValue', 0) as string;

	const workbook = await loadWorkbook.call(this, context.basePath);
	const worksheet = getWorksheet(workbook, sheetName, this);

	// Update the cell
	const cell = worksheet.getCell(cellRef);
	const oldValue = cell.value;
	cell.value = cellValue;

	// Upload back
	await saveWorkbook.call(this, context.basePath, workbook);

	return [
		{
			json: {
				success: true,
				cell: cellRef,
				oldValue,
				newValue: cellValue,
				sheet: sheetName,
			} as IDataObject,
		},
	];
}
