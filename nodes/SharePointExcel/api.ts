import * as ExcelJS from 'exceljs';
import type {
	IDataObject,
	IExecuteFunctions,
	IHttpRequestMethods,
	IHttpRequestOptions,
} from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import type { GraphError } from './types';

const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
const CREDENTIAL_NAME = 'microsoftGraphOAuth2Api';
const REQUEST_TIMEOUT = 30000;
const EXCEL_CONTENT_TYPE =
	'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/**
 * Make a Graph API request
 */
export async function graphRequest(
	this: IExecuteFunctions,
	method: IHttpRequestMethods,
	endpoint: string,
	body?: Buffer | IDataObject,
	isBuffer = false,
): Promise<unknown> {
	const url = `${GRAPH_BASE_URL}${endpoint}`;

	try {
		let response;

		if (isBuffer && method === 'PUT' && Buffer.isBuffer(body)) {
			// Use requestOAuth2 for binary uploads (httpRequestWithAuthentication doesn't handle buffers correctly)
			const options = {
				method,
				uri: url,
				body: body,
				headers: {
					'Content-Type': EXCEL_CONTENT_TYPE,
					'Content-Length': body.length,
				},
				encoding: null,
				json: false,
			};
			// eslint-disable-next-line @n8n/community-nodes/no-deprecated-workflow-functions
			response = await this.helpers.requestOAuth2.call(
				this,
				CREDENTIAL_NAME,
				options,
			);
		} else if (isBuffer && method === 'GET') {
			// Use httpRequestWithAuthentication for downloads
			const options: IHttpRequestOptions = {
				method,
				url,
				encoding: 'arraybuffer',
				json: false,
				timeout: REQUEST_TIMEOUT,
			};
			response = await this.helpers.httpRequestWithAuthentication.call(
				this,
				CREDENTIAL_NAME,
				options,
			);
		} else {
			// Regular JSON requests
			const options: IHttpRequestOptions = {
				method,
				url,
				json: true,
				timeout: REQUEST_TIMEOUT,
			};
			if (body && !Buffer.isBuffer(body)) {
				options.body = body as IDataObject;
			}
			response = await this.helpers.httpRequestWithAuthentication.call(
				this,
				CREDENTIAL_NAME,
				options,
			);
		}

		// Check if response is an error object from Graph API
		if (response && typeof response === 'object' && 'error' in response) {
			const errorResponse = response as GraphError;
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
}

/**
 * Download Excel file as ArrayBuffer
 */
export async function download(
	this: IExecuteFunctions,
	basePath: string,
): Promise<ArrayBuffer> {
	const response = await graphRequest.call(
		this,
		'GET',
		`${basePath}/content`,
		undefined,
		true,
	);
	return response as ArrayBuffer;
}

/**
 * Upload Excel file
 */
export async function upload(
	this: IExecuteFunctions,
	basePath: string,
	data: Buffer | ArrayBuffer,
): Promise<void> {
	const buffer = Buffer.isBuffer(data) ? data : Buffer.from(data);
	await graphRequest.call(this, 'PUT', `${basePath}/content`, buffer, true);
}

/**
 * Download and parse Excel workbook
 */
export async function loadWorkbook(
	this: IExecuteFunctions,
	basePath: string,
): Promise<ExcelJS.Workbook> {
	const buffer = await download.call(this, basePath);
	const workbook = new ExcelJS.Workbook();
	await workbook.xlsx.load(buffer);
	return workbook;
}

/**
 * Save workbook and upload
 */
export async function saveWorkbook(
	this: IExecuteFunctions,
	basePath: string,
	workbook: ExcelJS.Workbook,
): Promise<void> {
	const buffer = await workbook.xlsx.writeBuffer();
	await upload.call(this, basePath, buffer);
}

/**
 * Get worksheet by name, throwing if not found
 */
export function getWorksheet(
	workbook: ExcelJS.Workbook,
	sheetName: string,
	node: IExecuteFunctions,
	itemIndex?: number,
): ExcelJS.Worksheet {
	const worksheet = workbook.getWorksheet(sheetName);
	if (!worksheet) {
		throw new NodeOperationError(
			node.getNode(),
			`Sheet "${sheetName}" not found in workbook`,
			itemIndex !== undefined ? { itemIndex } : undefined,
		);
	}
	return worksheet;
}

/**
 * Delete a file via Graph API
 */
export async function deleteFile(
	this: IExecuteFunctions,
	basePath: string,
): Promise<void> {
	await graphRequest.call(this, 'DELETE', basePath);
}

/**
 * List items in a drive folder
 */
export async function listDriveItems(
	this: IExecuteFunctions,
	driveId: string,
	folderId?: string,
): Promise<unknown> {
	const path = folderId
		? `/drives/${driveId}/items/${folderId}/children`
		: `/drives/${driveId}/root/children`;
	return graphRequest.call(this, 'GET', path);
}
