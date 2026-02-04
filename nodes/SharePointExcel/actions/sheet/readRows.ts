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
	const options = this.getNodeParameter('options', 0) as ReadRowsOptions;
	const headerRow = options.headerRow || 1;
	const startRow = options.startRow || 2;
	const maxRows = options.maxRows || 0;

	const workbook = await loadWorkbook.call(this, context.basePath);
	const worksheet = getWorksheet(workbook, sheetName, this);

	// Get headers from header row
	const headers: string[] = [];
	const headerRowData = worksheet.getRow(headerRow);
	headerRowData.eachCell({ includeEmpty: false }, (cell, colNumber) => {
		headers[colNumber] = String(cell.value || `Column${colNumber}`);
	});

	// Read data rows
	let rowCount = 0;
	for (let rowNum = startRow; rowNum <= worksheet.rowCount; rowNum++) {
		if (maxRows > 0 && rowCount >= maxRows) break;

		const row = worksheet.getRow(rowNum);
		const rowData: IDataObject = {};
		let hasData = false;

		row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
			const header = headers[colNumber] || `Column${colNumber}`;
			rowData[header] = cell.value as string | number | boolean;
			hasData = true;
		});

		if (hasData) {
			returnData.push({ json: rowData });
			rowCount++;
		}
	}

	if (returnData.length === 0) {
		returnData.push({ json: { message: 'No data found in sheet' } });
	}

	return returnData;
}
