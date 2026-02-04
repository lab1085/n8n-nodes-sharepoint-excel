import type { IDataObject, IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import { getWorksheet, loadWorkbook, saveWorkbook } from '../../api';
import type { OperationContext, ResourceLocatorValue, RowData } from '../../types';

interface UpsertOptions {
	keyColumn?: string;
	headerRow?: number;
}

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
	const options = this.getNodeParameter('options', 0) as UpsertOptions;
	const keyColumn = options.keyColumn;
	const headerRow = options.headerRow || 1;

	if (!keyColumn) {
		throw new NodeOperationError(
			this.getNode(),
			'Key Column is required for upsert operation',
		);
	}

	// Process each input item
	for (let i = 0; i < items.length; i++) {
		const rowDataParam = this.getNodeParameter('rowData', i) as string;
		let rowsToUpsert: RowData[];

		try {
			const parsed = JSON.parse(rowDataParam);
			rowsToUpsert = Array.isArray(parsed) ? parsed : [parsed];
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

		// Get existing headers
		const headers: string[] = [];
		const headerRowData = worksheet.getRow(headerRow);
		headerRowData.eachCell({ includeEmpty: false }, (cell, colNumber) => {
			headers[colNumber] = String(cell.value);
		});

		// Build header-to-column map
		const headerMap: Record<string, number> = {};
		headers.forEach((h, idx) => {
			if (h) headerMap[h] = idx;
		});

		// Find key column index
		const keyColIndex = headerMap[keyColumn];
		if (keyColIndex === undefined) {
			throw new NodeOperationError(
				this.getNode(),
				`Key column "${keyColumn}" not found in headers`,
				{ itemIndex: i },
			);
		}

		// Build index of existing rows by key value
		const existingRows: Map<string | number | boolean, number> = new Map();
		for (let rowNum = headerRow + 1; rowNum <= worksheet.rowCount; rowNum++) {
			const row = worksheet.getRow(rowNum);
			const keyCell = row.getCell(keyColIndex);
			if (keyCell.value !== null && keyCell.value !== undefined) {
				existingRows.set(keyCell.value as string | number | boolean, rowNum);
			}
		}

		let updatedCount = 0;
		let appendedCount = 0;

		// Process each row to upsert
		for (const rowData of rowsToUpsert) {
			const keyValue = rowData[keyColumn];

			if (keyValue === undefined || keyValue === null) {
				throw new NodeOperationError(
					this.getNode(),
					`Row data missing key column "${keyColumn}" value`,
					{ itemIndex: i },
				);
			}

			const existingRowNum = existingRows.get(keyValue as string | number | boolean);

			if (existingRowNum !== undefined) {
				// Update existing row
				const row = worksheet.getRow(existingRowNum);
				for (const [key, value] of Object.entries(rowData)) {
					const colIdx = headerMap[key];
					if (colIdx !== undefined) {
						row.getCell(colIdx).value = value as string | number | boolean | null;
					}
				}
				updatedCount++;
			} else {
				// Append new row
				const newRow: (string | number | boolean | null)[] = [];
				for (const [key, value] of Object.entries(rowData)) {
					const colIdx = headerMap[key];
					if (colIdx !== undefined) {
						newRow[colIdx] = value as string | number | boolean | null;
					}
				}
				worksheet.addRow(newRow);
				appendedCount++;
			}
		}

		// Upload back
		await saveWorkbook.call(this, context.basePath, workbook);

		returnData.push({
			json: {
				success: true,
				rowsUpdated: updatedCount,
				rowsAppended: appendedCount,
				sheet: sheetName,
			} as IDataObject,
		});
	}

	return returnData;
}
