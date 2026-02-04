import { vi } from 'vitest';
import type {
	IExecuteFunctions,
	ILoadOptionsFunctions,
	INodeExecutionData,
} from 'n8n-workflow';
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
 * Creates a mock ILoadOptionsFunctions for testing listSearch methods
 */
export function createMockLoadOptionsFunctions(params: Record<string, unknown> = {}) {
	const getNodeParameter = vi.fn((name: string) => {
		if (name in params) {
			return params[name];
		}
		return undefined;
	});

	// Create the mock function that will be called via .call()
	const httpRequestMock = vi.fn();

	// The actual function needs a .call method that invokes the mock
	const httpRequestWithAuthentication = Object.assign(
		vi.fn(),
		{ call: httpRequestMock },
	);

	return {
		getNodeParameter,
		getNode: vi.fn(() => ({ name: 'TestNode' })),
		logger: {
			error: vi.fn(),
			info: vi.fn(),
			debug: vi.fn(),
			warn: vi.fn(),
		},
		helpers: {
			httpRequestWithAuthentication,
		},
		// Expose for easy access in tests
		_httpRequestWithAuthentication: httpRequestMock,
	} as unknown as ILoadOptionsFunctions & {
		_httpRequestWithAuthentication: ReturnType<typeof vi.fn>;
	};
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

/** Options for worksheet mock */
export interface MockWorksheetOptions {
	headers: Record<number, string>;
	rows: Record<number, string | number | boolean | null>[];
	/** Override rowCount (defaults to rows.length + 1) */
	rowCount?: number;
	/** Override columnCount (defaults to Object.keys(headers).length) */
	columnCount?: number;
}

/**
 * Creates a mock exceljs Worksheet
 */
export function createMockWorksheet(data: MockWorksheetOptions) {
	const headerRow = createMockRow(data.headers);
	const dataRows = data.rows.map((row) => createMockRow(row));

	const rowCount = data.rowCount ?? data.rows.length + 1;
	const columnCount = data.columnCount ?? Object.keys(data.headers).length;

	return {
		rowCount,
		columnCount,
		getRow: vi.fn((rowNum: number) => {
			if (rowNum === 1) return headerRow;
			return dataRows[rowNum - 2] || createMockRow({});
		}),
		spliceRows: vi.fn(),
		addRow: vi.fn(),
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

/** Options for setupSheetMocks */
export interface SetupSheetMocksOptions {
	sheetName: string;
	rowCount: number;
	columnCount: number;
	headers?: Record<number, string>;
	rows?: Record<number, string | number | boolean | null>[];
}

/**
 * Sets up common mocks for sheet operations.
 * Call this after vi.mock() and importing the mocked modules.
 *
 * @example
 * ```typescript
 * vi.mock('../../api', () => ({
 *   loadWorkbook: vi.fn(),
 *   saveWorkbook: vi.fn(),
 *   getWorksheet: vi.fn(),
 * }));
 *
 * import { loadWorkbook, saveWorkbook, getWorksheet } from '../../api';
 *
 * it('test', async () => {
 *   const { worksheet, workbook } = setupSheetMocks(
 *     { loadWorkbook, saveWorkbook, getWorksheet },
 *     { sheetName: 'Sheet1', rowCount: 3, columnCount: 2 }
 *   );
 *   // ... test code
 * });
 * ```
 */
export function setupSheetMocks(
	mocks: {
		loadWorkbook: unknown;
		saveWorkbook: unknown;
		getWorksheet: unknown;
	},
	options: SetupSheetMocksOptions,
) {
	const worksheet = createMockWorksheet({
		headers: options.headers ?? {},
		rows: options.rows ?? [],
		rowCount: options.rowCount,
		columnCount: options.columnCount,
	});

	const workbook = createMockWorkbook({ [options.sheetName]: worksheet });

	vi.mocked(mocks.loadWorkbook as ReturnType<typeof vi.fn>).mockResolvedValue(workbook as never);
	vi.mocked(mocks.getWorksheet as ReturnType<typeof vi.fn>).mockReturnValue(worksheet as never);
	vi.mocked(mocks.saveWorkbook as ReturnType<typeof vi.fn>).mockResolvedValue(undefined);

	return { worksheet, workbook };
}
