import * as ExcelJS from 'exceljs';
import type {
	ILoadOptionsFunctions,
	ResourceMapperFields,
} from 'n8n-workflow';
import type { ResourceLocatorValue, ResourceMapperField } from './types';

const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
const CREDENTIAL_NAME = 'microsoftGraphOAuth2Api';

/**
 * Extract value from resource locator parameter
 */
function getResourceValue(param: string | ResourceLocatorValue): string {
	return typeof param === 'object' ? param.value : param;
}

/**
 * Infer field type from ExcelJS cell value
 */
function inferFieldType(
	value: ExcelJS.CellValue,
): ResourceMapperField['type'] {
	if (value === null || value === undefined) {
		return 'string';
	}
	if (typeof value === 'number') {
		return 'number';
	}
	if (typeof value === 'boolean') {
		return 'boolean';
	}
	if (value instanceof Date) {
		return 'dateTime';
	}
	return 'string';
}

/**
 * Get mapping columns from Excel sheet headers for resourceMapper UI
 */
export async function getMappingColumns(
	this: ILoadOptionsFunctions,
): Promise<ResourceMapperFields> {
	const fields: ResourceMapperField[] = [];

	try {
		const siteIdParam = this.getNodeParameter('siteId') as
			| string
			| ResourceLocatorValue;
		const siteId = getResourceValue(siteIdParam);

		const driveIdParam = this.getNodeParameter('driveId') as
			| string
			| ResourceLocatorValue;
		const driveId = getResourceValue(driveIdParam);

		const fileIdParam = this.getNodeParameter('fileId') as
			| string
			| ResourceLocatorValue;
		const fileId = getResourceValue(fileIdParam);

		const sheetNameParam = this.getNodeParameter('sheetName') as
			| string
			| ResourceLocatorValue;
		const sheetName = getResourceValue(sheetNameParam);

		if (!siteId || !driveId || !fileId || !sheetName) {
			return { fields };
		}

		// Get header row from options (default: 1)
		let headerRow = 1;
		try {
			const options = this.getNodeParameter('options', {}) as {
				headerRow?: number;
			};
			headerRow = options.headerRow || 1;
		} catch {
			// Options might not be available yet
		}

		// Download the file
		const endpoint = `${GRAPH_BASE_URL}/sites/${siteId}/drives/${driveId}/items/${fileId}/content`;
		const response = await this.helpers.httpRequestWithAuthentication.call(
			this,
			CREDENTIAL_NAME,
			{
				method: 'GET',
				url: endpoint,
				encoding: 'arraybuffer',
				json: false,
			},
		);

		// Parse with ExcelJS
		const workbook = new ExcelJS.Workbook();
		await workbook.xlsx.load(response as ArrayBuffer);

		const worksheet = workbook.getWorksheet(sheetName);
		if (!worksheet) {
			return { fields };
		}

		// Get headers from the specified row
		const headerRowData = worksheet.getRow(headerRow);
		const dataRow = worksheet.getRow(headerRow + 1);

		headerRowData.eachCell({ includeEmpty: false }, (cell, colNumber) => {
			const headerValue = String(cell.value || '').trim();
			if (headerValue) {
				// Try to infer type from the first data row
				const dataCell = dataRow.getCell(colNumber);
				const fieldType = inferFieldType(dataCell.value);

				fields.push({
					id: headerValue,
					displayName: headerValue,
					required: false,
					defaultMatch: false,
					display: true,
					type: fieldType,
					canBeUsedToMatch: true,
				});
			}
		});
	} catch (err) {
		this.logger.error('Failed to load mapping columns', {
			error: err instanceof Error ? err.message : String(err),
		});
	}

	return { fields };
}
