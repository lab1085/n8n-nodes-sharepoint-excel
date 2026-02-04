import type { IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import { loadWorkbook } from '../../api';
import type { OperationContext, SheetInfo } from '../../types';

export async function execute(
	this: IExecuteFunctions,
	_items: INodeExecutionData[],
	context: OperationContext,
): Promise<INodeExecutionData[]> {
	const workbook = await loadWorkbook.call(this, context.basePath);

	const sheets: SheetInfo[] = workbook.worksheets.map((ws) => ({
		name: ws.name,
		id: ws.id,
		rowCount: ws.rowCount,
		columnCount: ws.columnCount,
	}));

	return [{ json: { sheets } }];
}
