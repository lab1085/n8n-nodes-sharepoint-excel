import * as ExcelJS from 'exceljs';
import type {
	IDataObject,
	IExecuteFunctions,
	IHttpRequestMethods,
	IHttpRequestOptions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
} from 'n8n-workflow';
import { NodeConnectionTypes, NodeOperationError } from 'n8n-workflow';

export class SharePointExcel implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'SharePoint Excel',
		name: 'sharePointExcel',
		icon: 'file:excel.svg',
		group: ['transform'],
		version: 1,
		subtitle: '={{$parameter["operation"]}}',
		description:
			'Read and write Excel files in SharePoint or OneDrive (bypasses WAC token issues).',
		defaults: {
			name: 'SharePoint Excel',
		},
		inputs: [NodeConnectionTypes.Main],
		outputs: [NodeConnectionTypes.Main],
		usableAsTool: true,
		credentials: [
			{
				name: 'microsoftGraphOAuth2Api',
				required: true,
			},
		],
		properties: [
			// Source selector
			{
				displayName: 'Source',
				name: 'source',
				type: 'options',
				noDataExpression: true,
				options: [
					{
						name: 'SharePoint',
						value: 'sharepoint',
						description: 'Excel file stored in a SharePoint document library',
					},
					{
						name: 'OneDrive',
						value: 'onedrive',
						description: 'Excel file stored in OneDrive',
					},
				],
				default: 'sharepoint',
			},

			// Operation selector
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				noDataExpression: true,
				options: [
					{
						name: 'Append Rows',
						value: 'appendRows',
						description: 'Add rows to the end of a sheet',
						action: 'Append rows to sheet',
					},
					{
						name: 'Get Sheets',
						value: 'getSheets',
						description: 'Get list of sheets in the workbook',
						action: 'Get sheet names',
					},
					{
						name: 'Read Rows',
						value: 'readRows',
						description: 'Read rows from a sheet',
						action: 'Read rows from sheet',
					},
					{
						name: 'Update Cell',
						value: 'updateCell',
						description: 'Update a specific cell value',
						action: 'Update cell value',
					},
				],
				default: 'readRows',
			},

			// Site ID - SharePoint only
			{
				displayName: 'Site ID',
				name: 'siteId',
				type: 'string',
				required: true,
				default: '',
				placeholder: 'contoso.sharepoint.com,site-guid,web-guid',
				description:
					'The SharePoint site ID. Find it via Graph Explorer: GET /sites/{hostname}:/{site-path}.',
				displayOptions: {
					show: {
						source: ['sharepoint'],
					},
				},
			},

			// Drive ID - shown for both but with different descriptions
			{
				displayName: 'Drive ID',
				name: 'driveId',
				type: 'string',
				required: true,
				default: '',
				placeholder: 'b!xxxxx',
				description:
					'The drive ID. For SharePoint: GET /sites/{siteId}/drives. For OneDrive: GET /me/drives.',
			},

			// File ID
			{
				displayName: 'File ID',
				name: 'fileId',
				type: 'string',
				required: true,
				default: '',
				placeholder: '01ABCDEF...',
				description: 'The Excel file item ID. Find via: GET /drives/{driveId}/root/children.',
			},

			// Sheet name (not for getSheets)
			{
				displayName: 'Sheet Name',
				name: 'sheetName',
				type: 'string',
				required: true,
				default: 'Sheet1',
				description: 'Name of the worksheet',
				displayOptions: {
					hide: {
						operation: ['getSheets'],
					},
				},
			},

			// Options for readRows
			{
				displayName: 'Options',
				name: 'options',
				type: 'collection',
				placeholder: 'Add Option',
				default: {},
				displayOptions: {
					show: {
						operation: ['readRows'],
					},
				},
				options: [
					{
						displayName: 'Header Row',
						name: 'headerRow',
						type: 'number',
						default: 1,
						description: 'Row number containing headers (1-indexed)',
					},
					{
						displayName: 'Start Row',
						name: 'startRow',
						type: 'number',
						default: 2,
						description: 'First data row to read (1-indexed)',
					},
					{
						displayName: 'Max Rows',
						name: 'maxRows',
						type: 'number',
						default: 0,
						description: 'Maximum rows to return (0 = all)',
					},
				],
			},

			// Row data for appendRows
			{
				displayName: 'Row Data',
				name: 'rowData',
				type: 'json',
				required: true,
				default: '{}',
				description:
					'JSON object with column headers as keys, or array of objects for multiple rows',
				displayOptions: {
					show: {
						operation: ['appendRows'],
					},
				},
			},

			// Cell reference for updateCell
			{
				displayName: 'Cell Reference',
				name: 'cellRef',
				type: 'string',
				required: true,
				default: 'A1',
				placeholder: 'A1',
				description: 'Cell to update (e.g., A1, B5, C10)',
				displayOptions: {
					show: {
						operation: ['updateCell'],
					},
				},
			},

			// Cell value for updateCell
			{
				displayName: 'Value',
				name: 'cellValue',
				type: 'string',
				required: true,
				default: '',
				description: 'New value for the cell',
				displayOptions: {
					show: {
						operation: ['updateCell'],
					},
				},
			},
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const returnData: INodeExecutionData[] = [];

		const source = this.getNodeParameter('source', 0) as string;
		const operation = this.getNodeParameter('operation', 0) as string;
		const driveId = this.getNodeParameter('driveId', 0) as string;
		const fileId = this.getNodeParameter('fileId', 0) as string;

		// Build base path based on source
		let basePath: string;
		if (source === 'sharepoint') {
			const siteId = this.getNodeParameter('siteId', 0) as string;
			basePath = `/sites/${siteId}/drives/${driveId}/items/${fileId}`;
		} else {
			// OneDrive - uses drive directly
			basePath = `/drives/${driveId}/items/${fileId}`;
		}

		// Helper: Make Graph API request
		const graphRequest = async (
			method: IHttpRequestMethods,
			endpoint: string,
			body?: Buffer | IDataObject,
			isBuffer = false,
		) => {
			const url = `https://graph.microsoft.com/v1.0${endpoint}`;

			
			try {
				let response;

				if (isBuffer && method === 'PUT' && Buffer.isBuffer(body)) {
					// Use requestOAuth2 for binary uploads (httpRequestWithAuthentication doesn't handle buffers correctly)
					const options = {
						method,
						uri: url,
						body: body,
						headers: {
							'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
							'Content-Length': body.length,
						},
						encoding: null,
						json: false,
					};
					// eslint-disable-next-line @n8n/community-nodes/no-deprecated-workflow-functions
					response = await this.helpers.requestOAuth2.call(
						this,
						'microsoftGraphOAuth2Api',
						options,
					);
				} else if (isBuffer && method === 'GET') {
					// Use httpRequestWithAuthentication for downloads
					const options: IHttpRequestOptions = {
						method,
						url,
						encoding: 'arraybuffer',
						json: false,
						timeout: 30000,
					};
					response = await this.helpers.httpRequestWithAuthentication.call(
						this,
						'microsoftGraphOAuth2Api',
						options,
					);
				} else {
					// Regular JSON requests
					const options: IHttpRequestOptions = {
						method,
						url,
						json: true,
						timeout: 30000,
					};
					if (body && !Buffer.isBuffer(body)) {
						options.body = body as IDataObject;
					}
					response = await this.helpers.httpRequestWithAuthentication.call(
						this,
						'microsoftGraphOAuth2Api',
						options,
					);
				}

				
				// Check if response is an error object from Graph API
				if (response && typeof response === 'object' && 'error' in response) {
					const errorResponse = response as { error: { message?: string; code?: string } };
					throw new NodeOperationError(
						this.getNode(),
						errorResponse.error.message || 'Graph API request failed',
						{ description: `Error code: ${errorResponse.error.code || 'unknown'}` },
					);
				}

				return response;
			} catch (err) {
								// Re-throw NodeOperationError as-is
				if (err instanceof NodeOperationError) {
					throw err;
				}
				// Wrap other errors
				const error = err as Error;
				throw new NodeOperationError(
					this.getNode(),
					`Graph API request failed: ${error.message}`,
				);
			}
		};

		// Download Excel file as ArrayBuffer
		const downloadExcel = async (): Promise<ArrayBuffer> => {
			const response = await graphRequest('GET', `${basePath}/content`, undefined, true);
			return response as ArrayBuffer;
		};

		// Upload Excel file
		const uploadExcel = async (data: Buffer | ArrayBuffer): Promise<void> => {
			const buffer = Buffer.isBuffer(data) ? data : Buffer.from(data);
			await graphRequest('PUT', `${basePath}/content`, buffer, true);
		};

		try {
			if (operation === 'getSheets') {
				// Download and parse to get sheet names
				const buffer = await downloadExcel();
				const workbook = new ExcelJS.Workbook();
				await workbook.xlsx.load(buffer);

				const sheets = workbook.worksheets.map((ws) => ({
					name: ws.name,
					id: ws.id,
					rowCount: ws.rowCount,
					columnCount: ws.columnCount,
				}));

				returnData.push({ json: { sheets } });
			}

			if (operation === 'readRows') {
				const sheetName = this.getNodeParameter('sheetName', 0) as string;
				const options = this.getNodeParameter('options', 0) as IDataObject;
				const headerRow = (options.headerRow as number) || 1;
				const startRow = (options.startRow as number) || 2;
				const maxRows = (options.maxRows as number) || 0;

				const buffer = await downloadExcel();
				const workbook = new ExcelJS.Workbook();
				await workbook.xlsx.load(buffer);

				const worksheet = workbook.getWorksheet(sheetName);
				if (!worksheet) {
					throw new NodeOperationError(
						this.getNode(),
						`Sheet "${sheetName}" not found in workbook`,
					);
				}

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
			}

			if (operation === 'appendRows') {
				const sheetName = this.getNodeParameter('sheetName', 0) as string;

				// Process each input item
				for (let i = 0; i < items.length; i++) {
					const rowDataParam = this.getNodeParameter('rowData', i) as string;
					let rowsToAdd: IDataObject[];

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
					const buffer = await downloadExcel();
					const workbook = new ExcelJS.Workbook();
					await workbook.xlsx.load(buffer);

					const worksheet = workbook.getWorksheet(sheetName);
					if (!worksheet) {
						throw new NodeOperationError(this.getNode(), `Sheet "${sheetName}" not found`, {
							itemIndex: i,
						});
					}

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
					const newBuffer = await workbook.xlsx.writeBuffer();
					await uploadExcel(newBuffer);

					returnData.push({
						json: {
							success: true,
							rowsAdded: rowsToAdd.length,
							sheet: sheetName,
						},
					});
				}
			}

			if (operation === 'updateCell') {
				const sheetName = this.getNodeParameter('sheetName', 0) as string;
				const cellRef = this.getNodeParameter('cellRef', 0) as string;
				const cellValue = this.getNodeParameter('cellValue', 0) as string;

				const buffer = await downloadExcel();
				const workbook = new ExcelJS.Workbook();
				await workbook.xlsx.load(buffer);

				const worksheet = workbook.getWorksheet(sheetName);
				if (!worksheet) {
					throw new NodeOperationError(this.getNode(), `Sheet "${sheetName}" not found`);
				}

				// Update the cell
				const cell = worksheet.getCell(cellRef);
				const oldValue = cell.value;
				cell.value = cellValue;

				// Upload back
				const newBuffer = await workbook.xlsx.writeBuffer();
				await uploadExcel(newBuffer);

				returnData.push({
					json: {
						success: true,
						cell: cellRef,
						oldValue,
						newValue: cellValue,
						sheet: sheetName,
					},
				});
			}
		} catch (err) {
			// Re-throw NodeOperationError as-is to preserve error details
			if (err instanceof NodeOperationError) {
				if (this.continueOnFail()) {
					returnData.push({
						json: { error: err.message },
						pairedItem: { item: 0 },
					});
				} else {
					throw err;
				}
			} else {
				const error = err as Error;
				if (this.continueOnFail()) {
					returnData.push({
						json: { error: error.message },
						pairedItem: { item: 0 },
					});
				} else {
					throw new NodeOperationError(this.getNode(), error.message);
				}
			}
		}

		return [returnData];
	}
}
