import * as ExcelJS from 'exceljs';
import type {
	ILoadOptionsFunctions,
	INodeListSearchItems,
	INodeListSearchResult,
} from 'n8n-workflow';
import type {
	GraphDrive,
	GraphDriveItem,
	GraphSite,
	ResourceLocatorValue,
} from './types';

const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
const CREDENTIAL_NAME = 'microsoftGraphOAuth2Api';

/**
 * Extract value from resource locator parameter
 */
function getResourceValue(param: string | ResourceLocatorValue): string {
	return typeof param === 'object' ? param.value : param;
}

/**
 * Search SharePoint sites
 */
export async function searchSites(
	this: ILoadOptionsFunctions,
	filter?: string,
): Promise<INodeListSearchResult> {
	const results: INodeListSearchItems[] = [];

	try {
		// Use wildcard '*' to list all sites when no filter provided
		const searchTerm = filter?.trim() || '*';
		const response = await this.helpers.httpRequestWithAuthentication.call(
			this,
			CREDENTIAL_NAME,
			{
				method: 'GET',
				url: `${GRAPH_BASE_URL}/sites?search=${encodeURIComponent(searchTerm)}`,
				json: true,
			},
		);

		const sites = (response as { value?: GraphSite[] }).value || [];
		for (const site of sites) {
			results.push({
				name: site.displayName,
				value: site.id,
				url: site.webUrl,
			});
		}
	} catch {
		// Return empty results on error
	}

	return { results };
}

/**
 * Get drives for site or OneDrive
 */
export async function getDrives(
	this: ILoadOptionsFunctions,
): Promise<INodeListSearchResult> {
	const results: INodeListSearchItems[] = [];

	try {
		const siteIdParam = this.getNodeParameter('siteId') as
			| string
			| ResourceLocatorValue;
		const siteId = getResourceValue(siteIdParam);

		if (!siteId) {
			return { results };
		}
		const endpoint = `${GRAPH_BASE_URL}/sites/${siteId}/drives`;

		const response = await this.helpers.httpRequestWithAuthentication.call(
			this,
			CREDENTIAL_NAME,
			{
				method: 'GET',
				url: endpoint,
				json: true,
			},
		);

		const drives = (response as { value?: GraphDrive[] }).value || [];
		for (const drive of drives) {
			results.push({
				name: drive.name,
				value: drive.id,
			});
		}
	} catch {
		// Return empty results on error
	}

	return { results };
}

/**
 * Get Excel files in drive
 */
export async function getFiles(
	this: ILoadOptionsFunctions,
): Promise<INodeListSearchResult> {
	const results: INodeListSearchItems[] = [];

	try {
		const driveIdParam = this.getNodeParameter('driveId') as
			| string
			| ResourceLocatorValue;
		const driveId = getResourceValue(driveIdParam);

		if (!driveId) {
			return { results };
		}

		const response = await this.helpers.httpRequestWithAuthentication.call(
			this,
			CREDENTIAL_NAME,
			{
				method: 'GET',
				url: `${GRAPH_BASE_URL}/drives/${driveId}/root/children`,
				json: true,
			},
		);

		const items = (response as { value?: GraphDriveItem[] }).value || [];
		for (const item of items) {
			// Only show Excel files
			if (item.file && item.name.toLowerCase().endsWith('.xlsx')) {
				results.push({
					name: item.name,
					value: item.id,
				});
			}
		}
	} catch {
		// Return empty results on error
	}

	return { results };
}

/**
 * Get sheets from Excel workbook
 */
export async function getSheets(
	this: ILoadOptionsFunctions,
): Promise<INodeListSearchResult> {
	const results: INodeListSearchItems[] = [];

	try {
		const driveIdParam = this.getNodeParameter('driveId') as
			| string
			| ResourceLocatorValue;
		const driveId = getResourceValue(driveIdParam);
		const fileIdParam = this.getNodeParameter('fileId') as
			| string
			| ResourceLocatorValue;
		const fileId = getResourceValue(fileIdParam);

		if (!driveId || !fileId) {
			return { results };
		}

		// Build endpoint for SharePoint
		const siteIdParam = this.getNodeParameter('siteId') as
			| string
			| ResourceLocatorValue;
		const siteId = getResourceValue(siteIdParam);
		const endpoint = `${GRAPH_BASE_URL}/sites/${siteId}/drives/${driveId}/items/${fileId}/content`;

		// Download the file
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

		// Parse with exceljs
		const workbook = new ExcelJS.Workbook();
		await workbook.xlsx.load(response as ArrayBuffer);

		for (const worksheet of workbook.worksheets) {
			results.push({
				name: worksheet.name,
				value: worksheet.name,
			});
		}
	} catch {
		// Return empty results on error
	}

	return { results };
}

/**
 * Get tables from Excel workbook via Graph API
 * Note: ExcelJS doesn't expose table metadata, so we use Graph API for this
 */
export async function getTables(
	this: ILoadOptionsFunctions,
): Promise<INodeListSearchResult> {
	const results: INodeListSearchItems[] = [];

	try {
		const driveIdParam = this.getNodeParameter('driveId') as
			| string
			| ResourceLocatorValue;
		const driveId = getResourceValue(driveIdParam);
		const fileIdParam = this.getNodeParameter('fileId') as
			| string
			| ResourceLocatorValue;
		const fileId = getResourceValue(fileIdParam);

		if (!driveId || !fileId) {
			return { results };
		}

		// Build endpoint for SharePoint - use workbook/tables API
		const siteIdParam = this.getNodeParameter('siteId') as
			| string
			| ResourceLocatorValue;
		const siteId = getResourceValue(siteIdParam);
		const endpoint = `${GRAPH_BASE_URL}/sites/${siteId}/drives/${driveId}/items/${fileId}/workbook/tables`;

		// Get tables via Graph API
		const response = await this.helpers.httpRequestWithAuthentication.call(
			this,
			CREDENTIAL_NAME,
			{
				method: 'GET',
				url: endpoint,
				json: true,
			},
		);

		interface TableInfo {
			id: string;
			name: string;
		}
		const tables = (response as { value?: TableInfo[] }).value || [];
		for (const table of tables) {
			results.push({
				name: table.name,
				value: table.name,
			});
		}
	} catch {
		// Return empty results on error
	}

	return { results };
}
