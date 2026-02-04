import { describe, it, expect, vi, beforeEach } from 'vitest';
import { NodeOperationError } from 'n8n-workflow';
import { execute } from './upsertRows';
import {
	createMockExecuteFunctions,
	createMockContext,
	createMockWorkbook,
} from '../../test-utils/mocks';

// Mock the api module
vi.mock('../../api', () => ({
	loadWorkbook: vi.fn(),
	saveWorkbook: vi.fn(),
	getWorksheet: vi.fn(),
}));

import { loadWorkbook, saveWorkbook, getWorksheet } from '../../api';

/** Options for upsert worksheet mock */
interface UpsertMockOptions {
	sheetName?: string;
	headers: Record<number, string>;
	existingRows?: Record<number, string | number | boolean | null>[];
}

/** Result from setupUpsertMocks */
interface UpsertMockResult {
	worksheet: ReturnType<typeof createUpsertWorksheet>;
	workbook: ReturnType<typeof createMockWorkbook>;
	cellUpdates: Map<number, Map<number, unknown>>;
	addedRows: unknown[][];
}

/** Creates a worksheet mock with trackable cell updates for upsert testing */
function createUpsertWorksheet(
	headers: Record<number, string>,
	existingRows: Record<number, string | number | boolean | null>[],
	cellUpdates: Map<number, Map<number, unknown>>,
	addedRows: unknown[][],
) {
	const headerRow = {
		eachCell: vi.fn(
			(
				_opts: { includeEmpty: boolean },
				callback: (cell: { value: unknown }, colNumber: number) => void,
			) => {
				Object.entries(headers).forEach(([col, value]) => {
					callback({ value }, Number(col));
				});
			},
		),
		getCell: vi.fn((colNumber: number) => ({
			value: headers[colNumber] ?? null,
		})),
	};

	const createDataRow = (rowNum: number, data: Record<number, unknown>) => {
		if (!cellUpdates.has(rowNum)) {
			cellUpdates.set(rowNum, new Map());
		}
		const rowUpdates = cellUpdates.get(rowNum)!;

		return {
			eachCell: vi.fn(
				(
					_opts: { includeEmpty: boolean },
					callback: (cell: { value: unknown }, colNumber: number) => void,
				) => {
					Object.entries(data).forEach(([col, value]) => {
						if (value !== null && value !== undefined) {
							callback({ value }, Number(col));
						}
					});
				},
			),
			getCell: vi.fn((colNumber: number) => {
				const cellValue = rowUpdates.get(colNumber) ?? data[colNumber] ?? null;
				return {
					get value() {
						return cellValue;
					},
					set value(v: unknown) {
						rowUpdates.set(colNumber, v);
					},
				};
			}),
		};
	};

	return {
		rowCount: existingRows.length + 1,
		columnCount: Object.keys(headers).length,
		getRow: vi.fn((rowNum: number) => {
			if (rowNum === 1) return headerRow;
			const dataIndex = rowNum - 2;
			if (dataIndex >= 0 && dataIndex < existingRows.length) {
				return createDataRow(rowNum, existingRows[dataIndex]);
			}
			return createDataRow(rowNum, {});
		}),
		addRow: vi.fn((row: unknown[]) => {
			addedRows.push([...row]);
		}),
		spliceRows: vi.fn(),
	};
}

/** Sets up all mocks needed for upsert operation tests */
function setupUpsertMocks(options: UpsertMockOptions): UpsertMockResult {
	const { sheetName = 'Sheet1', headers, existingRows = [] } = options;

	const cellUpdates: Map<number, Map<number, unknown>> = new Map();
	const addedRows: unknown[][] = [];

	const worksheet = createUpsertWorksheet(headers, existingRows, cellUpdates, addedRows);
	const workbook = createMockWorkbook({
		[sheetName]: worksheet as unknown as Parameters<typeof createMockWorkbook>[0][string],
	});

	vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);
	vi.mocked(getWorksheet).mockReturnValue(worksheet as never);
	vi.mocked(saveWorkbook).mockResolvedValue(undefined);

	return { worksheet, workbook, cellUpdates, addedRows };
}

describe('upsertRows', () => {
	beforeEach(() => {
		vi.clearAllMocks();
	});

	describe('basic functionality', () => {
		it('updates existing row and saves the workbook', async () => {
			setupUpsertMocks({
				headers: { 1: 'ID', 2: 'Name', 3: 'Email' },
				existingRows: [{ 1: 'A1', 2: 'John', 3: 'john@example.com' }],
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				keyColumn: 'ID',
				options: {},
			});

			const context = createMockContext({ operation: 'upsertRows' });
			const result = await execute.call(
				mockFunctions,
				[{ json: { ID: 'A1', Name: 'John Updated', Email: 'john.updated@example.com' } }],
				context,
			);

			expect(loadWorkbook).toHaveBeenCalledTimes(1);
			expect(saveWorkbook).toHaveBeenCalledTimes(1);
			expect(result[0].json).toEqual({
				success: true,
				rowsUpdated: 1,
				rowsAppended: 0,
				sheet: 'Sheet1',
			});
		});

		it('handles resourceLocator format for sheetName', async () => {
			const { workbook } = setupUpsertMocks({
				sheetName: 'MySheet',
				headers: { 1: 'ID', 2: 'Name' },
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: { mode: 'list', value: 'MySheet' },
				dataMode: 'autoMap',
				keyColumn: 'ID',
				options: {},
			});

			const context = createMockContext({ operation: 'upsertRows' });
			await execute.call(mockFunctions, [{ json: { ID: 'A1', Name: 'Test' } }], context);

			expect(getWorksheet).toHaveBeenCalledWith(workbook, 'MySheet', mockFunctions, 0);
		});

		it('uses correct basePath from context', async () => {
			const { workbook } = setupUpsertMocks({
				headers: { 1: 'ID', 2: 'Name' },
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				keyColumn: 'ID',
				options: {},
			});

			const customBasePath = '/sites/custom-site/drives/custom-drive/items/custom-file';
			const context = createMockContext({
				operation: 'upsertRows',
				basePath: customBasePath,
			});

			await execute.call(mockFunctions, [{ json: { ID: 'A1', Name: 'Test' } }], context);

			expect(loadWorkbook).toHaveBeenCalledWith(customBasePath);
			expect(saveWorkbook).toHaveBeenCalledWith(customBasePath, workbook);
		});
	});

	describe('update behavior', () => {
		it('updates all columns for matched row', async () => {
			const { cellUpdates } = setupUpsertMocks({
				headers: { 1: 'ID', 2: 'Name', 3: 'Email' },
				existingRows: [{ 1: 'A1', 2: 'John', 3: 'john@example.com' }],
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				keyColumn: 'ID',
				options: {},
			});

			const context = createMockContext({ operation: 'upsertRows' });
			await execute.call(
				mockFunctions,
				[{ json: { ID: 'A1', Name: 'John Updated', Email: 'new@example.com' } }],
				context,
			);

			const row2Updates = cellUpdates.get(2);
			expect(row2Updates?.get(1)).toBe('A1');
			expect(row2Updates?.get(2)).toBe('John Updated');
			expect(row2Updates?.get(3)).toBe('new@example.com');
		});

		it('only updates columns that exist in headers', async () => {
			const { cellUpdates } = setupUpsertMocks({
				headers: { 1: 'ID', 2: 'Name' },
				existingRows: [{ 1: 'A1', 2: 'John' }],
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				keyColumn: 'ID',
				options: {},
			});

			const context = createMockContext({ operation: 'upsertRows' });
			await execute.call(
				mockFunctions,
				[{ json: { ID: 'A1', Name: 'John Updated', ExtraField: 'ignored' } }],
				context,
			);

			const row2Updates = cellUpdates.get(2);
			expect(row2Updates?.get(1)).toBe('A1');
			expect(row2Updates?.get(2)).toBe('John Updated');
			expect(row2Updates?.get(3)).toBeUndefined();
		});

		it('updates multiple rows in batch', async () => {
			const { cellUpdates } = setupUpsertMocks({
				headers: { 1: 'ID', 2: 'Name' },
				existingRows: [
					{ 1: 'A1', 2: 'John' },
					{ 1: 'A2', 2: 'Jane' },
					{ 1: 'A3', 2: 'Bob' },
				],
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				keyColumn: 'ID',
				options: {},
			});

			const context = createMockContext({ operation: 'upsertRows' });
			const result = await execute.call(
				mockFunctions,
				[{ json: { ID: 'A1', Name: 'John Updated' } }, { json: { ID: 'A3', Name: 'Bob Updated' } }],
				context,
			);

			expect(result[0].json.rowsUpdated).toBe(2);
			expect(cellUpdates.get(2)?.get(2)).toBe('John Updated');
			expect(cellUpdates.get(4)?.get(2)).toBe('Bob Updated');
		});
	});

	describe('insert behavior', () => {
		it('appends new row when key not found', async () => {
			const { worksheet, addedRows } = setupUpsertMocks({
				headers: { 1: 'ID', 2: 'Name', 3: 'Email' },
				existingRows: [{ 1: 'A1', 2: 'John', 3: 'john@example.com' }],
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				keyColumn: 'ID',
				options: {},
			});

			const context = createMockContext({ operation: 'upsertRows' });
			const result = await execute.call(
				mockFunctions,
				[{ json: { ID: 'A2', Name: 'New Person', Email: 'new@example.com' } }],
				context,
			);

			expect(result[0].json.rowsAppended).toBe(1);
			expect(worksheet.addRow).toHaveBeenCalledTimes(1);
			expect(addedRows[0][1]).toBe('A2');
			expect(addedRows[0][2]).toBe('New Person');
			expect(addedRows[0][3]).toBe('new@example.com');
		});

		it('writes data to correct columns (1-indexed)', async () => {
			const { addedRows } = setupUpsertMocks({
				headers: { 1: 'ID', 2: 'Name', 3: 'Email' },
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				keyColumn: 'ID',
				options: {},
			});

			const context = createMockContext({ operation: 'upsertRows' });
			await execute.call(
				mockFunctions,
				[{ json: { ID: 'A1', Name: 'Test', Email: 'test@example.com' } }],
				context,
			);

			expect(addedRows[0][0]).toBeUndefined();
			expect(addedRows[0][1]).toBe('A1');
			expect(addedRows[0][2]).toBe('Test');
			expect(addedRows[0][3]).toBe('test@example.com');
		});

		it('handles partial data (not all columns provided)', async () => {
			const { addedRows } = setupUpsertMocks({
				headers: { 1: 'ID', 2: 'Name', 3: 'Email' },
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				keyColumn: 'ID',
				options: {},
			});

			const context = createMockContext({ operation: 'upsertRows' });
			await execute.call(mockFunctions, [{ json: { ID: 'A1', Name: 'Test' } }], context);

			expect(addedRows[0][1]).toBe('A1');
			expect(addedRows[0][2]).toBe('Test');
			expect(addedRows[0][3]).toBeUndefined();
		});
	});

	describe('key column matching', () => {
		it('matches single key column (autoMap mode)', async () => {
			const { cellUpdates, addedRows } = setupUpsertMocks({
				headers: { 1: 'ID', 2: 'Name' },
				existingRows: [
					{ 1: 'A1', 2: 'John' },
					{ 1: 'A2', 2: 'Jane' },
				],
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				keyColumn: 'ID',
				options: {},
			});

			const context = createMockContext({ operation: 'upsertRows' });
			const result = await execute.call(
				mockFunctions,
				[{ json: { ID: 'A2', Name: 'Jane Updated' } }, { json: { ID: 'A3', Name: 'New' } }],
				context,
			);

			expect(result[0].json.rowsUpdated).toBe(1);
			expect(result[0].json.rowsAppended).toBe(1);
			expect(cellUpdates.get(3)?.get(2)).toBe('Jane Updated');
			expect(addedRows[0][1]).toBe('A3');
		});

		it('matches composite key from multiple columns (manual mode)', async () => {
			const { cellUpdates } = setupUpsertMocks({
				headers: { 1: 'FirstName', 2: 'LastName', 3: 'Email' },
				existingRows: [
					{ 1: 'John', 2: 'Doe', 3: 'john.doe@example.com' },
					{ 1: 'Jane', 2: 'Doe', 3: 'jane.doe@example.com' },
				],
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'manual',
				columns: {
					mappingMode: 'defineBelow',
					value: { FirstName: 'Jane', LastName: 'Doe', Email: 'jane.updated@example.com' },
					matchingColumns: ['FirstName', 'LastName'],
				},
				options: {},
			});

			const context = createMockContext({ operation: 'upsertRows' });
			const result = await execute.call(mockFunctions, [{ json: {} }], context);

			expect(result[0].json.rowsUpdated).toBe(1);
			expect(cellUpdates.get(3)?.get(3)).toBe('jane.updated@example.com');
		});

		it('handles mix of updates and inserts in same batch', async () => {
			const { cellUpdates, addedRows } = setupUpsertMocks({
				headers: { 1: 'ID', 2: 'Name' },
				existingRows: [{ 1: 'A1', 2: 'John' }],
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				keyColumn: 'ID',
				options: {},
			});

			const context = createMockContext({ operation: 'upsertRows' });
			const result = await execute.call(
				mockFunctions,
				[
					{ json: { ID: 'A1', Name: 'John Updated' } },
					{ json: { ID: 'A2', Name: 'New Person 1' } },
					{ json: { ID: 'A3', Name: 'New Person 2' } },
				],
				context,
			);

			expect(result[0].json).toEqual({
				success: true,
				rowsUpdated: 1,
				rowsAppended: 2,
				sheet: 'Sheet1',
			});
			expect(cellUpdates.get(2)?.get(2)).toBe('John Updated');
			expect(addedRows).toHaveLength(2);
		});
	});

	describe('data modes', () => {
		it('autoMap mode uses input JSON directly', async () => {
			const { addedRows } = setupUpsertMocks({
				headers: { 1: 'ID', 2: 'Name' },
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				keyColumn: 'ID',
				options: {},
			});

			const context = createMockContext({ operation: 'upsertRows' });
			await execute.call(
				mockFunctions,
				[{ json: { ID: 'AutoMapID', Name: 'AutoMap Value' } }],
				context,
			);

			expect(addedRows[0][1]).toBe('AutoMapID');
			expect(addedRows[0][2]).toBe('AutoMap Value');
		});

		it('manual mode uses resourceMapper columns value', async () => {
			const { addedRows } = setupUpsertMocks({
				headers: { 1: 'ID', 2: 'Name' },
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'manual',
				columns: {
					mappingMode: 'defineBelow',
					value: { ID: 'ManualID', Name: 'Manual Value' },
					matchingColumns: ['ID'],
				},
				options: {},
			});

			const context = createMockContext({ operation: 'upsertRows' });
			await execute.call(mockFunctions, [{ json: {} }], context);

			expect(addedRows[0][1]).toBe('ManualID');
			expect(addedRows[0][2]).toBe('Manual Value');
		});

		it('raw mode parses JSON string', async () => {
			const { addedRows } = setupUpsertMocks({
				headers: { 1: 'ID', 2: 'Name' },
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'raw',
				keyColumn: 'ID',
				rowData: '{"ID": "RawID", "Name": "Raw Value"}',
				options: {},
			});

			const context = createMockContext({ operation: 'upsertRows' });
			await execute.call(mockFunctions, [{ json: {} }], context);

			expect(addedRows[0][1]).toBe('RawID');
			expect(addedRows[0][2]).toBe('Raw Value');
		});

		it('raw mode handles array of objects', async () => {
			const { addedRows } = setupUpsertMocks({
				headers: { 1: 'ID', 2: 'Name' },
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'raw',
				keyColumn: 'ID',
				rowData: '[{"ID": "R1", "Name": "Row 1"}, {"ID": "R2", "Name": "Row 2"}]',
				options: {},
			});

			const context = createMockContext({ operation: 'upsertRows' });
			const result = await execute.call(mockFunctions, [{ json: {} }], context);

			expect(result[0].json.rowsAppended).toBe(2);
			expect(addedRows[0][1]).toBe('R1');
			expect(addedRows[1][1]).toBe('R2');
		});
	});

	describe('custom headerRow option', () => {
		it('reads headers from specified row', async () => {
			const { worksheet, cellUpdates } = setupUpsertMocks({
				headers: { 1: 'ID', 2: 'Name' },
				existingRows: [{ 1: 'A1', 2: 'John' }],
			});

			// Override getRow for custom header row at row 3
			worksheet.getRow = vi.fn((rowNum: number) => {
				if (!cellUpdates.has(rowNum)) {
					cellUpdates.set(rowNum, new Map());
				}
				const rowUpdates = cellUpdates.get(rowNum)!;

				if (rowNum === 3) {
					return {
						eachCell: vi.fn(
							(
								_opts: { includeEmpty: boolean },
								callback: (cell: { value: unknown }, col: number) => void,
							) => {
								callback({ value: 'ID' }, 1);
								callback({ value: 'Name' }, 2);
							},
						),
						getCell: vi.fn((col: number) => ({
							value: col === 1 ? 'ID' : 'Name',
						})),
					};
				}
				if (rowNum === 4) {
					return {
						eachCell: vi.fn(
							(
								_opts: { includeEmpty: boolean },
								callback: (cell: { value: unknown }, col: number) => void,
							) => {
								callback({ value: 'A1' }, 1);
								callback({ value: 'John' }, 2);
							},
						),
						getCell: vi.fn((col: number) => ({
							get value(): string | number | boolean | null {
								return (rowUpdates.get(col) as string) ?? (col === 1 ? 'A1' : 'John');
							},
							set value(v: string | number | boolean | null) {
								rowUpdates.set(col, v);
							},
						})),
					};
				}
				return {
					eachCell: vi.fn(),
					getCell: vi.fn(() => ({ value: null as string | number | boolean | null })),
				};
			}) as typeof worksheet.getRow;

			// Need to update rowCount for custom header scenario
			(worksheet as { rowCount: number }).rowCount = 4;

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				keyColumn: 'ID',
				options: { headerRow: 3 },
			});

			const context = createMockContext({ operation: 'upsertRows' });
			const result = await execute.call(
				mockFunctions,
				[{ json: { ID: 'A1', Name: 'John Updated' } }],
				context,
			);

			expect(worksheet.getRow).toHaveBeenCalledWith(3);
			expect(result[0].json.rowsUpdated).toBe(1);
			expect(cellUpdates.get(4)?.get(2)).toBe('John Updated');
		});
	});

	describe('edge cases', () => {
		it('handles empty sheet (headers only, no data rows)', async () => {
			const { addedRows } = setupUpsertMocks({
				headers: { 1: 'ID', 2: 'Name' },
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				keyColumn: 'ID',
				options: {},
			});

			const context = createMockContext({ operation: 'upsertRows' });
			const result = await execute.call(
				mockFunctions,
				[{ json: { ID: 'A1', Name: 'New' } }],
				context,
			);

			expect(result[0].json).toEqual({
				success: true,
				rowsUpdated: 0,
				rowsAppended: 1,
				sheet: 'Sheet1',
			});
			expect(addedRows).toHaveLength(1);
		});

		it('skips empty key values when indexing existing rows', async () => {
			const { addedRows } = setupUpsertMocks({
				headers: { 1: 'ID', 2: 'Name' },
				existingRows: [
					{ 1: '', 2: 'Empty ID Row' },
					{ 1: 'A1', 2: 'John' },
				],
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				keyColumn: 'ID',
				options: {},
			});

			const context = createMockContext({ operation: 'upsertRows' });
			const result = await execute.call(
				mockFunctions,
				[{ json: { ID: 'A2', Name: 'New Person' } }],
				context,
			);

			expect(result[0].json.rowsAppended).toBe(1);
			expect(addedRows[0][1]).toBe('A2');
		});
	});

	describe('error handling', () => {
		it('throws when keyColumn parameter missing (autoMap/raw)', async () => {
			setupUpsertMocks({
				headers: { 1: 'ID', 2: 'Name' },
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				options: {},
			});

			const context = createMockContext({ operation: 'upsertRows' });

			await expect(
				execute.call(mockFunctions, [{ json: { ID: 'A1', Name: 'Test' } }], context),
			).rejects.toThrow('Key Column is required for upsert operation');
		});

		it('throws when no matchingColumns selected (manual)', async () => {
			setupUpsertMocks({
				headers: { 1: 'ID', 2: 'Name' },
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'manual',
				columns: {
					mappingMode: 'defineBelow',
					value: { ID: 'A1', Name: 'Test' },
					matchingColumns: [],
				},
				options: {},
			});

			const context = createMockContext({ operation: 'upsertRows' });

			await expect(execute.call(mockFunctions, [{ json: {} }], context)).rejects.toThrow(
				'At least one matching column must be selected in manual mode',
			);
		});

		it('throws when key column not found in headers', async () => {
			setupUpsertMocks({
				headers: { 1: 'ID', 2: 'Name' },
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				keyColumn: 'NonExistentColumn',
				options: {},
			});

			const context = createMockContext({ operation: 'upsertRows' });

			await expect(
				execute.call(mockFunctions, [{ json: { NonExistentColumn: 'A1' } }], context),
			).rejects.toThrow('Key column "NonExistentColumn" not found in headers');
		});

		it('throws when row data missing key column value', async () => {
			setupUpsertMocks({
				headers: { 1: 'ID', 2: 'Name' },
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				keyColumn: 'ID',
				options: {},
			});

			const context = createMockContext({ operation: 'upsertRows' });

			await expect(
				execute.call(mockFunctions, [{ json: { Name: 'Test' } }], context),
			).rejects.toThrow('Row data missing key column "ID" value');
		});

		it('throws when columns value empty (manual mode)', async () => {
			setupUpsertMocks({
				headers: { 1: 'ID', 2: 'Name' },
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'manual',
				columns: {
					mappingMode: 'defineBelow',
					value: {},
					matchingColumns: ['ID'],
				},
				options: {},
			});

			const context = createMockContext({ operation: 'upsertRows' });

			await expect(execute.call(mockFunctions, [{ json: {} }], context)).rejects.toThrow(
				'No column values provided in manual mapping mode',
			);
		});

		it('throws for invalid JSON in raw mode', async () => {
			setupUpsertMocks({
				headers: { 1: 'ID', 2: 'Name' },
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'raw',
				keyColumn: 'ID',
				rowData: 'not valid json {{{',
				options: {},
			});

			const context = createMockContext({ operation: 'upsertRows' });

			await expect(execute.call(mockFunctions, [{ json: {} }], context)).rejects.toThrow(
				'Invalid JSON in Row Data',
			);
		});

		it('throws for unknown data mode', async () => {
			setupUpsertMocks({
				headers: { 1: 'ID', 2: 'Name' },
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'invalidMode',
				keyColumn: 'ID',
				options: {},
			});

			const context = createMockContext({ operation: 'upsertRows' });

			await expect(execute.call(mockFunctions, [{ json: { ID: 'A1' } }], context)).rejects.toThrow(
				'Unknown data mode: invalidMode',
			);
		});

		it('throws when sheet not found', async () => {
			setupUpsertMocks({
				headers: { 1: 'ID' },
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'NonExistent',
				dataMode: 'autoMap',
				keyColumn: 'ID',
				options: {},
			});

			vi.mocked(getWorksheet).mockImplementation(() => {
				throw new NodeOperationError(
					mockFunctions.getNode(),
					'Sheet "NonExistent" not found in workbook',
				);
			});

			const context = createMockContext({ operation: 'upsertRows' });

			await expect(execute.call(mockFunctions, [{ json: { ID: 'A1' } }], context)).rejects.toThrow(
				'Sheet "NonExistent" not found in workbook',
			);
		});

		it('throws when loadWorkbook fails', async () => {
			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				keyColumn: 'ID',
				options: {},
			});

			vi.mocked(loadWorkbook).mockRejectedValue(
				new NodeOperationError(mockFunctions.getNode(), 'Graph API request failed: Access denied'),
			);

			const context = createMockContext({ operation: 'upsertRows' });

			await expect(execute.call(mockFunctions, [{ json: { ID: 'A1' } }], context)).rejects.toThrow(
				'Graph API request failed: Access denied',
			);
		});

		it('throws when saveWorkbook fails', async () => {
			setupUpsertMocks({
				headers: { 1: 'ID', 2: 'Name' },
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				keyColumn: 'ID',
				options: {},
			});

			vi.mocked(saveWorkbook).mockRejectedValue(
				new NodeOperationError(mockFunctions.getNode(), 'Graph API request failed: File locked'),
			);

			const context = createMockContext({ operation: 'upsertRows' });

			await expect(
				execute.call(mockFunctions, [{ json: { ID: 'A1', Name: 'Test' } }], context),
			).rejects.toThrow('Graph API request failed: File locked');
		});
	});
});
