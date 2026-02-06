import type { IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import { NodeOperationError } from 'n8n-workflow';
import type { OperationContext, Resource, Operation, ResourceLocatorValue } from '../types';
import * as sheet from './sheet';
import * as workbook from './workbook';
import * as table from './table';

/**
 * Build operation context from node parameters
 */
export function buildContext(executeFunctions: IExecuteFunctions): OperationContext {
	const source = 'sharepoint' as const;
	const resource = executeFunctions.getNodeParameter('resource', 0) as Resource;
	const operation = executeFunctions.getNodeParameter('operation', 0) as Operation;

	// Handle resourceLocator format for driveId
	const driveIdParam = executeFunctions.getNodeParameter('driveId', 0) as
		| string
		| ResourceLocatorValue;
	const driveId = typeof driveIdParam === 'object' ? driveIdParam.value : driveIdParam;

	// fileId is not required for getWorkbooks operation
	let fileId: string | undefined;
	if (operation !== 'getWorkbooks') {
		const fileIdParam = executeFunctions.getNodeParameter('fileId', 0) as
			| string
			| ResourceLocatorValue;
		fileId = typeof fileIdParam === 'object' ? fileIdParam.value : fileIdParam;
	}

	// siteId is not required for getWorkbooks operation
	let siteId: string | undefined;
	let basePath = '';
	if (operation !== 'getWorkbooks') {
		const siteIdParam = executeFunctions.getNodeParameter('siteId', 0) as
			| string
			| ResourceLocatorValue;
		siteId = typeof siteIdParam === 'object' ? siteIdParam.value : siteIdParam;
		basePath = `/sites/${siteId}/drives/${driveId}/items/${fileId}`;
	}

	return {
		source,
		resource,
		operation,
		basePath,
		driveId,
		fileId,
		siteId,
	};
}

/**
 * Route execution to the appropriate operation handler
 */
export async function router(
	this: IExecuteFunctions,
	items: INodeExecutionData[],
): Promise<INodeExecutionData[]> {
	const context = buildContext(this);
	const { resource, operation } = context;

	// Route to appropriate handler
	if (resource === 'sheet') {
		switch (operation) {
			case 'getSheets':
				return sheet.getSheets.execute.call(this, items, context);
			case 'readRows':
				return sheet.readRows.execute.call(this, items, context);
			case 'appendRows':
				return sheet.appendRows.execute.call(this, items, context);
			case 'updateCell':
				return sheet.updateRange.execute.call(this, items, context);
			case 'upsertRows':
				return sheet.upsertRows.execute.call(this, items, context);
			case 'clearSheet':
				return sheet.clearSheet.execute.call(this, items, context);
			case 'deleteSheet':
				return sheet.deleteSheet.execute.call(this, items, context);
			default:
				throw new NodeOperationError(this.getNode(), `Unknown sheet operation: ${operation}`);
		}
	}

	if (resource === 'table') {
		switch (operation) {
			case 'getColumns':
				return table.getColumns.execute.call(this, items, context);
			case 'getTableRows':
				return table.getRows.execute.call(this, items, context);
			case 'lookup':
				return table.lookup.execute.call(this, items, context);
			default:
				throw new NodeOperationError(this.getNode(), `Unknown table operation: ${operation}`);
		}
	}

	if (resource === 'workbook') {
		switch (operation) {
			case 'addSheet':
				return workbook.addSheet.execute.call(this, items, context);
			case 'deleteWorkbook':
				return workbook.deleteWorkbook.execute.call(this, items, context);
			case 'getWorkbooks':
				return workbook.getWorkbooks.execute.call(this, items, context);
			default:
				throw new NodeOperationError(this.getNode(), `Unknown workbook operation: ${operation}`);
		}
	}

	throw new NodeOperationError(this.getNode(), `Unknown resource: ${resource}`);
}
