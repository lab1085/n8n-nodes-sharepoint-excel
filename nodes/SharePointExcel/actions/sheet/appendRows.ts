import type { IDataObject, IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import { getWorksheet, loadWorkbook, saveWorkbook } from '../../api';
import type {
	DataMode,
	OperationContext,
	ResourceLocatorValue,
	ResourceMapperValue,
	RowData,
} from '../../types';

interface AppendOptions {
	headerRow?: number;
}

/**
 * Extract row data based on the selected data mode
 */
function getRowData(
	context: IExecuteFunctions,
	itemIndex: number,
	item: INodeExecutionData,
	dataMode: DataMode,
): RowData[] {
	switch (dataMode) {
		case 'autoMap': {
			// Use input JSON directly
			return [item.json as RowData];
		}
		case 'manual': {
			// Get from resourceMapper
			const columns = context.getNodeParameter(
				'columns',
				itemIndex,
			) as ResourceMapperValue;
			if (!columns.value || Object.keys(columns.value).length === 0) {
				throw new NodeOperationError(
					context.getNode(),
					'No column values provided in manual mapping mode',
					{ itemIndex },
				);
			}
			return [columns.value as RowData];
		}
		case 'raw': {
			// Parse JSON from rowData parameter
			const rowDataParam = context.getNodeParameter('rowData', itemIndex) as string;
			try {
				const parsed = JSON.parse(rowDataParam);
				return Array.isArray(parsed) ? parsed : [parsed];
			} catch (err) {
				throw new NodeOperationError(
					context.getNode(),
					`Invalid JSON in Row Data: ${(err as Error).message}`,
					{ itemIndex },
				);
			}
		}
		default:
			throw new NodeOperationError(
				context.getNode(),
				`Unknown data mode: ${dataMode}`,
				{ itemIndex },
			);
	}
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

	const dataMode = this.getNodeParameter('dataMode', 0, 'autoMap') as DataMode;
	const options = this.getNodeParameter('options', 0, {}) as AppendOptions;
	const headerRow = options.headerRow || 1;

	// Download workbook once
	const workbook = await loadWorkbook.call(this, context.basePath);
	const worksheet = getWorksheet(workbook, sheetName, this, 0);

	// Get existing headers
	const headers: string[] = [];
	const headerRowData = worksheet.getRow(headerRow);
	headerRowData.eachCell({ includeEmpty: false }, (cell, colNumber) => {
		headers[colNumber] = String(cell.value);
	});

	// Check if headers exist
	const hasHeaders = headers.some((h) => h && h.trim() !== '');
	const headerMap: Record<string, number> = {};

	// Track total rows added
	let totalRowsAdded = 0;

	// Process each input item
	for (let i = 0; i < items.length; i++) {
		const rowsToAdd = getRowData(this, i, items[i], dataMode);

		// If no headers yet, create them from the first row's keys
		if (!hasHeaders && i === 0 && rowsToAdd.length > 0) {
			const keys = Object.keys(rowsToAdd[0]);
			keys.forEach((key, idx) => {
				const colNumber = idx + 1; // Excel columns are 1-indexed
				headerRowData.getCell(colNumber).value = key;
				headers[colNumber] = key;
				headerMap[key] = colNumber;
			});
		} else if (i === 0) {
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
			totalRowsAdded++;
		}
	}

	// Upload back once
	await saveWorkbook.call(this, context.basePath, workbook);

	returnData.push({
		json: {
			success: true,
			rowsAdded: totalRowsAdded,
			sheet: sheetName,
		} as IDataObject,
	});

	return returnData;
}
