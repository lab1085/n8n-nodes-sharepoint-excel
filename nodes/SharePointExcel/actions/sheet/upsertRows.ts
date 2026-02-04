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

interface UpsertOptions {
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

/**
 * Get key columns for matching rows
 */
function getKeyColumns(
	context: IExecuteFunctions,
	dataMode: DataMode,
): string[] {
	if (dataMode === 'manual') {
		// Get from resourceMapper matchingColumns
		const columns = context.getNodeParameter('columns', 0) as ResourceMapperValue;
		const matchingColumns = columns.matchingColumns || [];
		if (matchingColumns.length === 0) {
			throw new NodeOperationError(
				context.getNode(),
				'At least one matching column must be selected in manual mode',
			);
		}
		return matchingColumns;
	} else {
		// Get from keyColumn parameter (autoMap/raw modes)
		const keyColumn = context.getNodeParameter('keyColumn', 0, '') as string;
		if (!keyColumn) {
			throw new NodeOperationError(
				context.getNode(),
				'Key Column is required for upsert operation',
			);
		}
		return [keyColumn];
	}
}

/**
 * Create composite key from row data
 */
function createKey(rowData: RowData, keyColumns: string[]): string {
	return keyColumns.map((col) => String(rowData[col] ?? '')).join('|');
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
	const options = this.getNodeParameter('options', 0, {}) as UpsertOptions;
	const headerRow = options.headerRow || 1;

	// Get key columns based on data mode
	const keyColumns = getKeyColumns(this, dataMode);

	// Download workbook once
	const workbook = await loadWorkbook.call(this, context.basePath);
	const worksheet = getWorksheet(workbook, sheetName, this, 0);

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

	// Validate key columns exist in headers
	for (const keyCol of keyColumns) {
		if (headerMap[keyCol] === undefined) {
			throw new NodeOperationError(
				this.getNode(),
				`Key column "${keyCol}" not found in headers`,
			);
		}
	}

	// Build index of existing rows by composite key
	const existingRows: Map<string, number> = new Map();
	for (let rowNum = headerRow + 1; rowNum <= worksheet.rowCount; rowNum++) {
		const row = worksheet.getRow(rowNum);
		const rowData: RowData = {};

		// Extract key column values
		for (const keyCol of keyColumns) {
			const colIdx = headerMap[keyCol];
			const cell = row.getCell(colIdx);
			rowData[keyCol] = cell.value as string | number | boolean | null;
		}

		const key = createKey(rowData, keyColumns);
		if (key && key !== keyColumns.map(() => '').join('|')) {
			existingRows.set(key, rowNum);
		}
	}

	let updatedCount = 0;
	let appendedCount = 0;

	// Process each input item
	for (let i = 0; i < items.length; i++) {
		const rowsToUpsert = getRowData(this, i, items[i], dataMode);

		for (const rowData of rowsToUpsert) {
			// Validate key values exist
			for (const keyCol of keyColumns) {
				if (rowData[keyCol] === undefined || rowData[keyCol] === null) {
					throw new NodeOperationError(
						this.getNode(),
						`Row data missing key column "${keyCol}" value`,
						{ itemIndex: i },
					);
				}
			}

			const key = createKey(rowData, keyColumns);
			const existingRowNum = existingRows.get(key);

			if (existingRowNum !== undefined) {
				// Update existing row
				const row = worksheet.getRow(existingRowNum);
				for (const [colName, value] of Object.entries(rowData)) {
					const colIdx = headerMap[colName];
					if (colIdx !== undefined) {
						row.getCell(colIdx).value = value as string | number | boolean | null;
					}
				}
				updatedCount++;
			} else {
				// Append new row
				const newRow: (string | number | boolean | null)[] = [];
				for (const [colName, value] of Object.entries(rowData)) {
					const colIdx = headerMap[colName];
					if (colIdx !== undefined) {
						newRow[colIdx] = value as string | number | boolean | null;
					}
				}
				worksheet.addRow(newRow);
				appendedCount++;
			}
		}
	}

	// Upload back once
	await saveWorkbook.call(this, context.basePath, workbook);

	returnData.push({
		json: {
			success: true,
			rowsUpdated: updatedCount,
			rowsAppended: appendedCount,
			sheet: sheetName,
		} as IDataObject,
	});

	return returnData;
}
