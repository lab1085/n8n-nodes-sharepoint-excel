import type {
	IExecuteFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
} from 'n8n-workflow';
import { NodeConnectionTypes, NodeOperationError } from 'n8n-workflow';
import { router } from './actions/router';
import { properties } from './descriptions';
import {
	searchSites,
	getDrives,
	getFiles,
	getSheets,
	getTables,
} from './listSearch';
import { getMappingColumns } from './resourceMapping';

export class SharePointExcel implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Microsoft SharePoint Excel',
		name: 'sharePointExcel',
		icon: 'file:excel.svg',
		group: ['transform'],
		version: 1,
		subtitle: '={{$parameter["resource"] + ": " + $parameter["operation"]}}',
		description:
			'Read and write Excel files in SharePoint (bypasses WAC token issues).',
		defaults: {
			name: 'Microsoft SharePoint Excel',
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
		properties,
	};

	methods = {
		listSearch: {
			searchSites,
			getDrives,
			getFiles,
			getSheets,
			getTables,
		},
		resourceMapping: {
			getMappingColumns,
		},
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		let returnData: INodeExecutionData[] = [];

		try {
			returnData = await router.call(this, items);
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
