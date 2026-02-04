import { describe, it, expect, vi, beforeEach } from 'vitest';
import { NodeOperationError } from 'n8n-workflow';
import { execute } from './clearSheet';
import {
	createMockExecuteFunctions,
	createMockContext,
	setupSheetMocks,
	createMockWorkbook,
} from '../../test-utils/mocks';

// Mock the api module
vi.mock('../../api', () => ({
	loadWorkbook: vi.fn(),
	saveWorkbook: vi.fn(),
	getWorksheet: vi.fn(),
}));

import { loadWorkbook, saveWorkbook, getWorksheet } from '../../api';

const apiMocks = { loadWorkbook, saveWorkbook, getWorksheet };

describe('clearSheet', () => {
	beforeEach(() => {
		vi.clearAllMocks();
	});

	describe('basic functionality', () => {
		it('clears all rows and saves the workbook', async () => {
			const { worksheet } = setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 3,
				columnCount: 2,
				headers: { 1: 'Name', 2: 'Email' },
				rows: [
					{ 1: 'John', 2: 'john@example.com' },
					{ 1: 'Jane', 2: 'jane@example.com' },
				],
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
			});

			const context = createMockContext({ operation: 'clearSheet' });
			const result = await execute.call(mockFunctions, [{ json: {} }], context);

			expect(loadWorkbook).toHaveBeenCalledTimes(1);
			expect(worksheet.spliceRows).toHaveBeenCalledWith(1, 3);
			expect(saveWorkbook).toHaveBeenCalledTimes(1);
			expect(result).toHaveLength(1);
			expect(result[0].json).toEqual({
				success: true,
				sheet: 'Sheet1',
				clearedRows: 3,
				clearedColumns: 2,
			});
		});

		it('handles resourceLocator format for sheetName', async () => {
			const { workbook } = setupSheetMocks(apiMocks, {
				sheetName: 'MySheet',
				rowCount: 1,
				columnCount: 1,
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: { mode: 'list', value: 'MySheet' },
			});

			const context = createMockContext({ operation: 'clearSheet' });
			const result = await execute.call(mockFunctions, [{ json: {} }], context);

			expect(getWorksheet).toHaveBeenCalledWith(workbook, 'MySheet', mockFunctions);
			expect(result[0].json.sheet).toBe('MySheet');
		});

		it('uses correct basePath from context', async () => {
			const { workbook } = setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 1,
				columnCount: 1,
			});

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
			});

			const customBasePath = '/sites/custom-site/drives/custom-drive/items/custom-file';
			const context = createMockContext({
				operation: 'clearSheet',
				basePath: customBasePath,
			});

			await execute.call(mockFunctions, [{ json: {} }], context);

			expect(loadWorkbook).toHaveBeenCalledWith(customBasePath);
			expect(saveWorkbook).toHaveBeenCalledWith(customBasePath, workbook);
		});
	});

	describe('row clearing behavior', () => {
		it('handles empty sheet (0 rows)', async () => {
			const { worksheet } = setupSheetMocks(apiMocks, {
				sheetName: 'EmptySheet',
				rowCount: 0,
				columnCount: 0,
			});
			const mockFunctions = createMockExecuteFunctions({ sheetName: 'EmptySheet' });
			const context = createMockContext({ operation: 'clearSheet' });

			const result = await execute.call(mockFunctions, [{ json: {} }], context);

			expect(worksheet.spliceRows).not.toHaveBeenCalled();
			expect(result[0].json).toEqual({
				success: true,
				sheet: 'EmptySheet',
				clearedRows: 0,
				clearedColumns: 0,
			});
		});

		it('handles sheet with single row', async () => {
			const { worksheet } = setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 1,
				columnCount: 1,
			});
			const mockFunctions = createMockExecuteFunctions({ sheetName: 'Sheet1' });
			const context = createMockContext({ operation: 'clearSheet' });

			const result = await execute.call(mockFunctions, [{ json: {} }], context);

			expect(worksheet.spliceRows).toHaveBeenCalledWith(1, 1);
			expect(result[0].json).toEqual({
				success: true,
				sheet: 'Sheet1',
				clearedRows: 1,
				clearedColumns: 1,
			});
		});

		it('handles large sheets', async () => {
			const { worksheet } = setupSheetMocks(apiMocks, {
				sheetName: 'LargeSheet',
				rowCount: 10000,
				columnCount: 50,
			});
			const mockFunctions = createMockExecuteFunctions({ sheetName: 'LargeSheet' });
			const context = createMockContext({ operation: 'clearSheet' });

			const result = await execute.call(mockFunctions, [{ json: {} }], context);

			expect(worksheet.spliceRows).toHaveBeenCalledWith(1, 10000);
			expect(result[0].json).toEqual({
				success: true,
				sheet: 'LargeSheet',
				clearedRows: 10000,
				clearedColumns: 50,
			});
		});
	});

	describe('edge cases', () => {
		it('ignores input items (operation is sheet-level)', async () => {
			setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 1,
				columnCount: 1,
			});
			const mockFunctions = createMockExecuteFunctions({ sheetName: 'Sheet1' });
			const context = createMockContext({ operation: 'clearSheet' });

			const multipleItems = [
				{ json: { data: 'item1' } },
				{ json: { data: 'item2' } },
				{ json: { data: 'item3' } },
			];

			const result = await execute.call(mockFunctions, multipleItems, context);

			expect(result).toHaveLength(1);
			expect(loadWorkbook).toHaveBeenCalledTimes(1);
			expect(saveWorkbook).toHaveBeenCalledTimes(1);
		});
	});

	describe('error handling', () => {
		it('throws error when sheet not found', async () => {
			const workbook = createMockWorkbook({});
			const mockFunctions = createMockExecuteFunctions({ sheetName: 'NonExistent' });
			const context = createMockContext({ operation: 'clearSheet' });

			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);
			vi.mocked(getWorksheet).mockImplementation(() => {
				throw new NodeOperationError(
					mockFunctions.getNode(),
					'Sheet "NonExistent" not found in workbook',
				);
			});

			await expect(execute.call(mockFunctions, [{ json: {} }], context)).rejects.toThrow(
				'Sheet "NonExistent" not found in workbook',
			);
		});

		it('throws error when loadWorkbook fails', async () => {
			const mockFunctions = createMockExecuteFunctions({ sheetName: 'Sheet1' });
			const context = createMockContext({ operation: 'clearSheet' });

			vi.mocked(loadWorkbook).mockRejectedValue(
				new NodeOperationError(mockFunctions.getNode(), 'Graph API request failed: Access denied'),
			);

			await expect(execute.call(mockFunctions, [{ json: {} }], context)).rejects.toThrow(
				'Graph API request failed: Access denied',
			);
		});

		it('throws error when saveWorkbook fails', async () => {
			setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 1,
				columnCount: 1,
			});
			const mockFunctions = createMockExecuteFunctions({ sheetName: 'Sheet1' });
			const context = createMockContext({ operation: 'clearSheet' });

			vi.mocked(saveWorkbook).mockRejectedValue(
				new NodeOperationError(mockFunctions.getNode(), 'Graph API request failed: File is locked'),
			);

			await expect(execute.call(mockFunctions, [{ json: {} }], context)).rejects.toThrow(
				'Graph API request failed: File is locked',
			);
		});
	});
});
