import { describe, it, expect, vi, beforeEach } from 'vitest';
import { NodeOperationError } from 'n8n-workflow';
import { execute } from './appendRows';
import {
	createMockExecuteFunctions,
	createMockContext,
	setupSheetMocks,
} from '../../test-utils/mocks';

// Mock the api module
vi.mock('../../api', () => ({
	loadWorkbook: vi.fn(),
	saveWorkbook: vi.fn(),
	getWorksheet: vi.fn(),
}));

import { loadWorkbook, saveWorkbook, getWorksheet } from '../../api';

const apiMocks = { loadWorkbook, saveWorkbook, getWorksheet };

describe('appendRows', () => {
	beforeEach(() => {
		vi.clearAllMocks();
	});

	describe('basic functionality', () => {
		it('calls saveWorkbook after appending', async () => {
			setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 1,
				columnCount: 1,
				headers: { 1: 'Name' },
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				options: {},
			});

			const context = createMockContext({ operation: 'appendRows' });
			await execute.call(mockFunctions, [{ json: { Name: 'ValueA' } }], context);

			expect(saveWorkbook).toHaveBeenCalledTimes(1);
		});

		it('handles resourceLocator format for sheetName', async () => {
			const { workbook } = setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 1,
				columnCount: 1,
				headers: { 1: 'Name' },
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: { mode: 'list', value: 'Sheet1' },
				dataMode: 'autoMap',
				options: {},
			});

			const context = createMockContext({ operation: 'appendRows' });
			await execute.call(mockFunctions, [{ json: { Name: 'ValueA' } }], context);

			expect(getWorksheet).toHaveBeenCalledWith(workbook, 'Sheet1', mockFunctions, 0);
		});

		it('uses correct basePath from context', async () => {
			const { workbook } = setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 1,
				columnCount: 1,
				headers: { 1: 'Name' },
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				options: {},
			});

			const context = createMockContext({
				operation: 'appendRows',
				basePath: '/sites/test-site/drives/test-drive/items/test-file',
			});
			await execute.call(mockFunctions, [{ json: { Name: 'ValueA' } }], context);

			expect(loadWorkbook).toHaveBeenCalledWith(
				'/sites/test-site/drives/test-drive/items/test-file',
			);
			expect(saveWorkbook).toHaveBeenCalledWith(
				'/sites/test-site/drives/test-drive/items/test-file',
				workbook,
			);
		});
	});

	describe('column mapping with existing headers', () => {
		it('writes data to correct columns (A, B) not shifted (B, C)', async () => {
			const { worksheet } = setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 2,
				columnCount: 2,
				headers: { 1: 'Name', 2: 'Email' },
				rows: [{ 1: 'Existing', 2: 'existing@example.com' }],
			});

			const addedRows: unknown[][] = [];
			worksheet.addRow = vi.fn((row: unknown[]) => {
				addedRows.push([...row]);
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				options: {},
			});

			const context = createMockContext({ operation: 'appendRows' });
			const inputItems = [{ json: { Name: 'ValueA', Email: 'a@example.com' } }];

			const result = await execute.call(mockFunctions, inputItems, context);

			expect(result[0].json).toEqual({
				success: true,
				rowsAdded: 1,
				sheet: 'Sheet1',
			});

			expect(worksheet.addRow).toHaveBeenCalledTimes(1);

			// Data should be at indices 1 and 2 (1-indexed for ExcelJS)
			expect(addedRows[0][1]).toBe('ValueA');
			expect(addedRows[0][2]).toBe('a@example.com');

			// Index 0 should be empty/undefined (not contain data)
			expect(addedRows[0][0]).toBeUndefined();
		});

		it('handles partial data (not all columns provided)', async () => {
			const { worksheet } = setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 1,
				columnCount: 3,
				headers: { 1: 'Name', 2: 'Email', 3: 'Age' },
			});

			const addedRows: unknown[][] = [];
			worksheet.addRow = vi.fn((row: unknown[]) => {
				addedRows.push([...row]);
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				options: {},
			});

			const context = createMockContext({ operation: 'appendRows' });
			const inputItems = [{ json: { Name: 'ValueA' } }];

			await execute.call(mockFunctions, inputItems, context);

			expect(addedRows[0][1]).toBe('ValueA');
			expect(addedRows[0][2]).toBeUndefined();
			expect(addedRows[0][3]).toBeUndefined();
		});

		it('ignores extra fields not in headers', async () => {
			const { worksheet } = setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 1,
				columnCount: 2,
				headers: { 1: 'Name', 2: 'Email' },
			});

			const addedRows: unknown[][] = [];
			worksheet.addRow = vi.fn((row: unknown[]) => {
				addedRows.push([...row]);
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				options: {},
			});

			const context = createMockContext({ operation: 'appendRows' });
			const inputItems = [
				{ json: { Name: 'ValueA', Email: 'a@example.com', ExtraField: 'ignored' } },
			];

			await execute.call(mockFunctions, inputItems, context);

			expect(addedRows[0][1]).toBe('ValueA');
			expect(addedRows[0][2]).toBe('a@example.com');
			expect(addedRows[0][3]).toBeUndefined();
		});
	});

	describe('custom headerRow option', () => {
		it('reads headers from specified row', async () => {
			const { worksheet } = setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 3,
				columnCount: 2,
			});

			// Override getRow for custom header row behavior
			worksheet.getRow = vi.fn((rowNum: number) => {
				if (rowNum === 3) {
					return {
						eachCell: vi.fn(
							(
								_opts: { includeEmpty: boolean },
								callback: (cell: { value: unknown }, col: number) => void,
							) => {
								callback({ value: 'Name' }, 1);
								callback({ value: 'Email' }, 2);
							},
						),
						getCell: vi.fn(),
					};
				}
				return { eachCell: vi.fn(), getCell: vi.fn(() => ({ value: null })) };
			});

			const addedRows: unknown[][] = [];
			worksheet.addRow = vi.fn((row: unknown[]) => {
				addedRows.push([...row]);
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				options: { headerRow: 3 },
			});

			const context = createMockContext({ operation: 'appendRows' });
			await execute.call(
				mockFunctions,
				[{ json: { Name: 'ValueA', Email: 'a@example.com' } }],
				context,
			);

			expect(worksheet.getRow).toHaveBeenCalledWith(3);
			expect(addedRows[0][1]).toBe('ValueA');
			expect(addedRows[0][2]).toBe('a@example.com');
		});
	});

	describe('creating headers on empty sheet', () => {
		it('creates headers and writes data to correct columns', async () => {
			const { worksheet } = setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 0,
				columnCount: 0,
			});

			const headerCellsWritten: Record<number, string> = {};
			const mockHeaderRow = {
				eachCell: vi.fn(),
				getCell: vi.fn((colNumber: number) => ({
					get value() {
						return headerCellsWritten[colNumber];
					},
					set value(v: string) {
						headerCellsWritten[colNumber] = v;
					},
				})),
			};

			// Override getRow for custom header row behavior
			worksheet.getRow = vi.fn((rowNum: number) => {
				if (rowNum === 1) return mockHeaderRow;
				return { eachCell: vi.fn(), getCell: vi.fn(() => ({ value: null })) };
			});

			const addedRows: unknown[][] = [];
			worksheet.addRow = vi.fn((row: unknown[]) => {
				addedRows.push([...row]);
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				options: {},
			});

			const context = createMockContext({ operation: 'appendRows' });
			const inputItems = [{ json: { Name: 'ValueA', Email: 'a@example.com' } }];

			await execute.call(mockFunctions, inputItems, context);

			expect(headerCellsWritten[1]).toBe('Name');
			expect(headerCellsWritten[2]).toBe('Email');
			expect(addedRows[0][1]).toBe('ValueA');
			expect(addedRows[0][2]).toBe('a@example.com');
		});
	});

	describe('multiple input items', () => {
		it('appends multiple rows correctly', async () => {
			const { worksheet } = setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 1,
				columnCount: 2,
				headers: { 1: 'Name', 2: 'Email' },
			});

			const addedRows: unknown[][] = [];
			worksheet.addRow = vi.fn((row: unknown[]) => {
				addedRows.push([...row]);
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				options: {},
			});

			const context = createMockContext({ operation: 'appendRows' });
			const inputItems = [
				{ json: { Name: 'Row1', Email: 'row1@example.com' } },
				{ json: { Name: 'Row2', Email: 'row2@example.com' } },
				{ json: { Name: 'Row3', Email: 'row3@example.com' } },
			];

			const result = await execute.call(mockFunctions, inputItems, context);

			expect(result[0].json.rowsAdded).toBe(3);
			expect(addedRows).toHaveLength(3);
			expect(addedRows[0][1]).toBe('Row1');
			expect(addedRows[1][1]).toBe('Row2');
			expect(addedRows[2][1]).toBe('Row3');
		});
	});

	describe('manual data mode', () => {
		it('uses resourceMapper columns value', async () => {
			const { worksheet } = setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 1,
				columnCount: 2,
				headers: { 1: 'Name', 2: 'Email' },
			});

			const addedRows: unknown[][] = [];
			worksheet.addRow = vi.fn((row: unknown[]) => {
				addedRows.push([...row]);
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'manual',
				columns: {
					mappingMode: 'defineBelow',
					value: { Name: 'ManualValue', Email: 'manual@example.com' },
				},
				options: {},
			});

			const context = createMockContext({ operation: 'appendRows' });
			await execute.call(mockFunctions, [{ json: {} }], context);

			expect(addedRows[0][1]).toBe('ManualValue');
			expect(addedRows[0][2]).toBe('manual@example.com');
		});

		it('throws error when columns value is empty', async () => {
			setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 1,
				columnCount: 1,
				headers: { 1: 'Name' },
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'manual',
				columns: {
					mappingMode: 'defineBelow',
					value: {},
				},
				options: {},
			});

			const context = createMockContext({ operation: 'appendRows' });

			await expect(execute.call(mockFunctions, [{ json: {} }], context)).rejects.toThrow(
				'No column values provided in manual mapping mode',
			);
		});
	});

	describe('unknown data mode', () => {
		it('throws error for invalid data mode', async () => {
			setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 1,
				columnCount: 1,
				headers: { 1: 'Name' },
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'invalidMode',
				options: {},
			});

			const context = createMockContext({ operation: 'appendRows' });

			await expect(
				execute.call(mockFunctions, [{ json: { Name: 'ValueA' } }], context),
			).rejects.toThrow('Unknown data mode: invalidMode');
		});
	});

	describe('raw data mode', () => {
		it('parses JSON and writes to correct columns', async () => {
			const { worksheet } = setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 1,
				columnCount: 2,
				headers: { 1: 'Name', 2: 'Email' },
			});

			const addedRows: unknown[][] = [];
			worksheet.addRow = vi.fn((row: unknown[]) => {
				addedRows.push([...row]);
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'raw',
				rowData: '{"Name": "RawValue", "Email": "raw@example.com"}',
				options: {},
			});

			const context = createMockContext({ operation: 'appendRows' });
			await execute.call(mockFunctions, [{ json: {} }], context);

			expect(addedRows[0][1]).toBe('RawValue');
			expect(addedRows[0][2]).toBe('raw@example.com');
		});

		it('throws error on invalid JSON', async () => {
			setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 1,
				columnCount: 1,
				headers: { 1: 'Name' },
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'raw',
				rowData: 'not valid json {{{',
				options: {},
			});

			const context = createMockContext({ operation: 'appendRows' });

			await expect(execute.call(mockFunctions, [{ json: {} }], context)).rejects.toThrow(
				'Invalid JSON in Row Data',
			);
		});

		it('handles array of objects in raw mode', async () => {
			const { worksheet } = setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 1,
				columnCount: 2,
				headers: { 1: 'Name', 2: 'Email' },
			});

			const addedRows: unknown[][] = [];
			worksheet.addRow = vi.fn((row: unknown[]) => {
				addedRows.push([...row]);
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'raw',
				rowData:
					'[{"Name": "Item1", "Email": "item1@example.com"}, {"Name": "Item2", "Email": "item2@example.com"}]',
				options: {},
			});

			const context = createMockContext({ operation: 'appendRows' });
			const result = await execute.call(mockFunctions, [{ json: {} }], context);

			expect(result[0].json.rowsAdded).toBe(2);
			expect(addedRows[0][1]).toBe('Item1');
			expect(addedRows[1][1]).toBe('Item2');
		});
	});

	describe('error handling', () => {
		it('throws error when sheet not found', async () => {
			setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 1,
				columnCount: 1,
				headers: { 1: 'Name' },
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'NonExistent',
				dataMode: 'autoMap',
				options: {},
			});

			vi.mocked(getWorksheet).mockImplementation(() => {
				throw new NodeOperationError(
					mockFunctions.getNode(),
					'Sheet "NonExistent" not found in workbook',
				);
			});

			const context = createMockContext({ operation: 'appendRows' });

			await expect(
				execute.call(mockFunctions, [{ json: { Name: 'ValueA' } }], context),
			).rejects.toThrow('Sheet "NonExistent" not found in workbook');
		});

		it('throws error when loadWorkbook fails', async () => {
			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				options: {},
			});

			vi.mocked(loadWorkbook).mockRejectedValue(
				new NodeOperationError(mockFunctions.getNode(), 'Graph API request failed: Access denied'),
			);

			const context = createMockContext({ operation: 'appendRows' });

			await expect(
				execute.call(mockFunctions, [{ json: { Name: 'ValueA' } }], context),
			).rejects.toThrow('Graph API request failed: Access denied');
		});

		it('throws error when saveWorkbook fails', async () => {
			setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 1,
				columnCount: 1,
				headers: { 1: 'Name' },
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				dataMode: 'autoMap',
				options: {},
			});

			vi.mocked(saveWorkbook).mockRejectedValue(
				new NodeOperationError(mockFunctions.getNode(), 'Graph API request failed: File locked'),
			);

			const context = createMockContext({ operation: 'appendRows' });

			await expect(
				execute.call(mockFunctions, [{ json: { Name: 'ValueA' } }], context),
			).rejects.toThrow('Graph API request failed: File locked');
		});
	});
});
