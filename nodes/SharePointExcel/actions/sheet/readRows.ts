import type { IDataObject, IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import { getWorksheet, loadWorkbook } from '../../api';
import type { OperationContext, ReadRowsOptions, ResourceLocatorValue } from '../../types';

export async function execute(
	this: IExecuteFunctions,
	_items: INodeExecutionData[],
	context: OperationContext,
): Promise<INodeExecutionData[]> {
	const returnData: INodeExecutionData[] = [];

	const sheetNameParam = this.getNodeParameter('sheetName', 0) as
		| string
		| ResourceLocatorValue;
	const sheetName =
		typeof sheetNameParam === 'object' ? sheetNameParam.value : sheetNameParam;

	// Get returnAll and limit parameters
	const returnAll = this.getNodeParameter('returnAll', 0, true) as boolean;
	const limit = returnAll ? 0 : (this.getNodeParameter('limit', 0, 50) as number);

	const options = this.getNodeParameter('options', 0) as ReadRowsOptions;
	const headerRow = options.headerRow || 1;
	const startRow = options.startRow || 2;
	const rawData = options.rawData || false;
	const dataProperty = options.dataProperty || 'data';
	const fieldsFilter = options.fields
		? new Set(options.fields.split(',').map((f) => f.trim()).filter(Boolean))
		: new Set<string>();

	const workbook = await loadWorkbook.call(this, context.basePath);
	const worksheet = getWorksheet(workbook, sheetName, this);

	// Get headers from header row
	const allHeaders: Map<number, string> = new Map();
	const headerRowData = worksheet.getRow(headerRow);
	headerRowData.eachCell({ includeEmpty: false }, (cell, colNumber) => {
		allHeaders.set(colNumber, String(cell.value || `Column${colNumber}`));
	});

	// Filter headers if fields specified
	const filteredColumns: number[] = [];
	const outputHeaders: string[] = [];

	allHeaders.forEach((header, colNumber) => {
		if (fieldsFilter.size === 0 || fieldsFilter.has(header)) {
			filteredColumns.push(colNumber);
			outputHeaders.push(header);
		}
	});

	// Read data rows
	if (rawData) {
		// RAW mode: return single item with arrays
		const rawRows: (string | number | boolean | null)[][] = [];
		let rowCount = 0;

		for (let rowNum = startRow; rowNum <= worksheet.rowCount; rowNum++) {
			if (limit > 0 && rowCount >= limit) break;

			const row = worksheet.getRow(rowNum);
			const rowArray: (string | number | boolean | null)[] = [];
			let hasData = false;

			for (const colNumber of filteredColumns) {
				const cell = row.getCell(colNumber);
				const value = cell.value as string | number | boolean | null;
				rowArray.push(value ?? null);
				if (value !== null && value !== undefined) {
					hasData = true;
				}
			}

			if (hasData) {
				rawRows.push(rowArray);
				rowCount++;
			}
		}

		if (rawRows.length === 0) {
			returnData.push({ json: { message: 'No data found in sheet' } });
		} else {
			returnData.push({
				json: {
					headers: outputHeaders,
					[dataProperty]: rawRows,
				},
			});
		}
	} else {
		// Normal mode: return keyed objects
		let rowCount = 0;

		for (let rowNum = startRow; rowNum <= worksheet.rowCount; rowNum++) {
			if (limit > 0 && rowCount >= limit) break;

			const row = worksheet.getRow(rowNum);
			const rowData: IDataObject = {};
			let hasData = false;

			for (let i = 0; i < filteredColumns.length; i++) {
				const colNumber = filteredColumns[i];
				const header = outputHeaders[i];
				const cell = row.getCell(colNumber);
				const value = cell.value as string | number | boolean | null;

				if (value !== null && value !== undefined) {
					rowData[header] = value;
					hasData = true;
				}
			}

			if (hasData) {
				returnData.push({ json: rowData });
				rowCount++;
			}
		}

		if (returnData.length === 0) {
			returnData.push({ json: { message: 'No data found in sheet' } });
		}
	}

	return returnData;
}
