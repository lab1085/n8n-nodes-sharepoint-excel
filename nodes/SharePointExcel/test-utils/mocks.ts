import { vi } from 'vitest';
import type { IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import type { OperationContext } from '../types';

/**
 * Creates a mock IExecuteFunctions for testing n8n node operations
 */
export function createMockExecuteFunctions(params: Record<string, unknown> = {}) {
	const getNodeParameter = vi.fn((name: string, _index: number, fallback?: unknown) => {
		if (name in params) {
			return params[name];
		}
		return fallback;
	});

	return {
		getNodeParameter,
		getNode: vi.fn(() => ({ name: 'TestNode' })),
		continueOnFail: vi.fn(() => false),
		getInputData: vi.fn(() => [{ json: {} }] as INodeExecutionData[]),
	} as unknown as IExecuteFunctions;
}

/**
 * Creates a mock OperationContext
 */
export function createMockContext(overrides: Partial<OperationContext> = {}): OperationContext {
	return {
		source: 'sharepoint',
		resource: 'sheet',
		operation: 'readRows',
		basePath: '/sites/test-site/drives/test-drive/items/test-file',
		driveId: 'test-drive',
		fileId: 'test-file',
		...overrides,
	};
}

/**
 * Creates a mock exceljs Cell
 */
export function createMockCell(value: string | number | boolean | null) {
	return { value };
}

/**
 * Creates a mock exceljs Row
 */
export function createMockRow(cells: Record<number, string | number | boolean | null>) {
	return {
		eachCell: vi.fn((opts: { includeEmpty: boolean }, callback: (cell: { value: unknown }, colNumber: number) => void) => {
			Object.entries(cells).forEach(([col, value]) => {
				if (!opts.includeEmpty && (value === null || value === undefined)) return;
				callback({ value }, Number(col));
			});
		}),
		getCell: vi.fn((colNumber: number) => createMockCell(cells[colNumber] ?? null)),
	};
}

/**
 * Creates a mock exceljs Worksheet
 */
export function createMockWorksheet(data: {
	headers: Record<number, string>;
	rows: Record<number, string | number | boolean | null>[];
}) {
	const headerRow = createMockRow(data.headers);
	const dataRows = data.rows.map((row) => createMockRow(row));

	return {
		rowCount: data.rows.length + 1, // +1 for header row
		getRow: vi.fn((rowNum: number) => {
			if (rowNum === 1) return headerRow;
			return dataRows[rowNum - 2] || createMockRow({});
		}),
	};
}

/**
 * Creates a mock exceljs Workbook
 */
export function createMockWorkbook(worksheets: Record<string, ReturnType<typeof createMockWorksheet>>) {
	return {
		getWorksheet: vi.fn((name: string) => worksheets[name]),
		xlsx: {
			load: vi.fn(),
		},
	};
}
