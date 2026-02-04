import type { IDataObject, IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import { getWorksheet, loadWorkbook, saveWorkbook } from '../../api';
import type { OperationContext, ResourceLocatorValue, RowData } from '../../types';

export async function execute(
	this: IExecuteFunctions,
	items: INodeExecutionData[],
	context: OperationContext,
): Promise<INodeExecutionData[]> {
	const returnData: INodeExecutionData[] = [];

	const sheetNameParam = this.getNodeParameter('sheetName', 0) as
		| string
		| ResourceLocatorValue;
	const sheetName =
		typeof sheetNameParam === 'object' ? sheetNameParam.value : sheetNameParam;

	// Process each input item
	for (let i = 0; i < items.length; i++) {
		const rowDataParam = this.getNodeParameter('rowData', i) as string;
		let rowsToAdd: RowData[];

		try {
			const parsed = JSON.parse(rowDataParam);
			rowsToAdd = Array.isArray(parsed) ? parsed : [parsed];
		} catch (err) {
			throw new NodeOperationError(
				this.getNode(),
				`Invalid JSON in Row Data: ${(err as Error).message}`,
				{ itemIndex: i },
			);
		}

		// Download current file
		const workbook = await loadWorkbook.call(this, context.basePath);
		const worksheet = getWorksheet(workbook, sheetName, this, i);

		// Get existing headers from row 1
		const headers: string[] = [];
		const headerRow = worksheet.getRow(1);
		headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
			headers[colNumber] = String(cell.value);
		});

		// Check if headers exist, if not create them from the first row's keys
		const hasHeaders = headers.some((h) => h && h.trim() !== '');
		const headerMap: Record<string, number> = {};

		if (!hasHeaders && rowsToAdd.length > 0) {
			// Create headers from the keys of the first row
			const keys = Object.keys(rowsToAdd[0]);
			keys.forEach((key, idx) => {
				const colNumber = idx + 1; // Excel columns are 1-indexed
				headerRow.getCell(colNumber).value = key;
				headers[colNumber] = key;
				headerMap[key] = colNumber;
			});
		} else {
			// Build header-to-column map from existing headers
			headers.forEach((h, idx) => {
				if (h) headerMap[h] = idx;
			});
		}

		// Add rows
		for (const rowData of rowsToAdd) {
			const newRow: (string | number | boolean | null)[] = [];
			for (const [key, value] of Object.entries(rowData)) {
				const colIdx = headerMap[key];
				if (colIdx !== undefined) {
					newRow[colIdx] = value as string | number | boolean | null;
				}
			}
			worksheet.addRow(newRow);
		}

		// Upload back
		await saveWorkbook.call(this, context.basePath, workbook);

		returnData.push({
			json: {
				success: true,
				rowsAdded: rowsToAdd.length,
				sheet: sheetName,
			} as IDataObject,
		});
	}

	return returnData;
}
