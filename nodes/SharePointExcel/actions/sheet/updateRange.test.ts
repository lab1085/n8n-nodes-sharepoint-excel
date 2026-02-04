import { describe, it, expect, vi, beforeEach } from 'vitest';
import { NodeOperationError } from 'n8n-workflow';
import { execute } from './updateRange';
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

describe('updateRange', () => {
	beforeEach(() => {
		vi.clearAllMocks();
	});

	describe('basic functionality', () => {
		it('updates a cell and saves the workbook', async () => {
			const { worksheet } = setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 3,
				columnCount: 2,
			});

			const mockCell = { value: 'oldValue' };
			worksheet.getCell = vi.fn(() => mockCell);

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				cellRef: 'A1',
				cellValue: 'newValue',
			});

			const context = createMockContext({ operation: 'updateCell' });
			const result = await execute.call(mockFunctions, [{ json: {} }], context);

			expect(loadWorkbook).toHaveBeenCalledTimes(1);
			expect(saveWorkbook).toHaveBeenCalledTimes(1);
			expect(result).toHaveLength(1);
			expect(result[0].json).toEqual({
				success: true,
				cell: 'A1',
				oldValue: 'oldValue',
				newValue: 'newValue',
				sheet: 'Sheet1',
			});
		});

		it('handles resourceLocator format for sheetName', async () => {
			const { worksheet, workbook } = setupSheetMocks(apiMocks, {
				sheetName: 'MySheet',
				rowCount: 1,
				columnCount: 1,
			});

			worksheet.getCell = vi.fn(() => ({ value: null }));

			const mockFunctions = createMockExecuteFunctions({
				sheetName: { mode: 'list', value: 'MySheet' },
				cellRef: 'B2',
				cellValue: 'test',
			});

			const context = createMockContext({ operation: 'updateCell' });
			const result = await execute.call(mockFunctions, [{ json: {} }], context);

			expect(getWorksheet).toHaveBeenCalledWith(workbook, 'MySheet', mockFunctions);
			expect(result[0].json.sheet).toBe('MySheet');
		});

		it('uses correct basePath from context', async () => {
			const { worksheet, workbook } = setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 1,
				columnCount: 1,
			});

			worksheet.getCell = vi.fn(() => ({ value: null }));

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				cellRef: 'A1',
				cellValue: 'test',
			});

			const customBasePath = '/sites/custom-site/drives/custom-drive/items/custom-file';
			const context = createMockContext({
				operation: 'updateCell',
				basePath: customBasePath,
			});

			await execute.call(mockFunctions, [{ json: {} }], context);

			expect(loadWorkbook).toHaveBeenCalledWith(customBasePath);
			expect(saveWorkbook).toHaveBeenCalledWith(customBasePath, workbook);
		});
	});

	describe('edge cases', () => {
		it('handles cell with no previous value (null oldValue)', async () => {
			const { worksheet } = setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 1,
				columnCount: 1,
			});

			const mockCell = { value: null };
			worksheet.getCell = vi.fn(() => mockCell);

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				cellRef: 'A1',
				cellValue: 'firstValue',
			});

			const context = createMockContext({ operation: 'updateCell' });
			const result = await execute.call(mockFunctions, [{ json: {} }], context);

			expect(result[0].json.oldValue).toBeNull();
			expect(result[0].json.newValue).toBe('firstValue');
		});

		it('ignores input items (operation is cell-level)', async () => {
			const { worksheet } = setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 1,
				columnCount: 1,
			});

			worksheet.getCell = vi.fn(() => ({ value: null }));

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				cellRef: 'A1',
				cellValue: 'test',
			});

			const context = createMockContext({ operation: 'updateCell' });

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
			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'NonExistent',
				cellRef: 'A1',
				cellValue: 'test',
			});
			const context = createMockContext({ operation: 'updateCell' });

			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);
			vi.mocked(getWorksheet).mockImplementation(() => {
				throw new NodeOperationError(
					mockFunctions.getNode(),
					'Sheet "NonExistent" not found in workbook',
				);
			});

			await expect(
				execute.call(mockFunctions, [{ json: {} }], context),
			).rejects.toThrow('Sheet "NonExistent" not found in workbook');
		});

		it('throws error when loadWorkbook fails', async () => {
			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				cellRef: 'A1',
				cellValue: 'test',
			});
			const context = createMockContext({ operation: 'updateCell' });

			vi.mocked(loadWorkbook).mockRejectedValue(
				new NodeOperationError(
					mockFunctions.getNode(),
					'Graph API request failed: Access denied',
				),
			);

			await expect(
				execute.call(mockFunctions, [{ json: {} }], context),
			).rejects.toThrow('Graph API request failed: Access denied');
		});

		it('throws error when saveWorkbook fails', async () => {
			const { worksheet } = setupSheetMocks(apiMocks, {
				sheetName: 'Sheet1',
				rowCount: 1,
				columnCount: 1,
			});

			worksheet.getCell = vi.fn(() => ({ value: null }));

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				cellRef: 'A1',
				cellValue: 'test',
			});
			const context = createMockContext({ operation: 'updateCell' });

			vi.mocked(saveWorkbook).mockRejectedValue(
				new NodeOperationError(
					mockFunctions.getNode(),
					'Graph API request failed: File is locked',
				),
			);

			await expect(
				execute.call(mockFunctions, [{ json: {} }], context),
			).rejects.toThrow('Graph API request failed: File is locked');
		});
	});
});
