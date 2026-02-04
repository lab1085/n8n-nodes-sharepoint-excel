import { describe, it, expect, vi, beforeEach } from 'vitest';
import { NodeOperationError } from 'n8n-workflow';
import { execute } from './getSheets';
import { createMockExecuteFunctions, createMockContext } from '../../test-utils/mocks';

// 1. Module mocks
vi.mock('../../api', () => ({
	loadWorkbook: vi.fn(),
}));

import { loadWorkbook } from '../../api';

// 2. Local helpers for getSheets-specific mocks
interface MockWorksheetInfo {
	name: string;
	id: number;
	rowCount: number;
	columnCount: number;
}

function createWorkbookForGetSheets(sheets: MockWorksheetInfo[]) {
	return {
		worksheets: sheets,
	};
}

// 3. Test suite
describe('getSheets', () => {
	beforeEach(() => {
		vi.clearAllMocks();
	});

	describe('basic functionality', () => {
		it('returns all worksheets with correct properties', async () => {
			const workbook = createWorkbookForGetSheets([
				{ name: 'Sheet1', id: 1, rowCount: 10, columnCount: 5 },
				{ name: 'Data', id: 2, rowCount: 100, columnCount: 3 },
				{ name: 'Summary', id: 3, rowCount: 50, columnCount: 8 },
			]);

			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);

			const mockFunctions = createMockExecuteFunctions();
			const context = createMockContext({ operation: 'getSheets' });
			const result = await execute.call(mockFunctions, [{ json: {} }], context);

			expect(result).toHaveLength(1);
			expect(result[0].json).toEqual({
				sheets: [
					{ name: 'Sheet1', id: 1, rowCount: 10, columnCount: 5 },
					{ name: 'Data', id: 2, rowCount: 100, columnCount: 3 },
					{ name: 'Summary', id: 3, rowCount: 50, columnCount: 8 },
				],
			});
		});

		it('uses basePath from context to load workbook', async () => {
			const workbook = createWorkbookForGetSheets([
				{ name: 'Sheet1', id: 1, rowCount: 5, columnCount: 2 },
			]);

			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);

			const mockFunctions = createMockExecuteFunctions();
			const customBasePath = '/sites/custom-site/drives/custom-drive/items/custom-file';
			const context = createMockContext({
				operation: 'getSheets',
				basePath: customBasePath,
			});

			await execute.call(mockFunctions, [{ json: {} }], context);

			expect(loadWorkbook).toHaveBeenCalledWith(customBasePath);
		});
	});

	describe('edge cases', () => {
		it('returns empty array when workbook has no sheets', async () => {
			const workbook = createWorkbookForGetSheets([]);

			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);

			const mockFunctions = createMockExecuteFunctions();
			const context = createMockContext({ operation: 'getSheets' });
			const result = await execute.call(mockFunctions, [{ json: {} }], context);

			expect(result).toHaveLength(1);
			expect(result[0].json).toEqual({ sheets: [] });
		});

		it('handles single sheet workbook', async () => {
			const workbook = createWorkbookForGetSheets([
				{ name: 'OnlySheet', id: 1, rowCount: 25, columnCount: 10 },
			]);

			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);

			const mockFunctions = createMockExecuteFunctions();
			const context = createMockContext({ operation: 'getSheets' });
			const result = await execute.call(mockFunctions, [{ json: {} }], context);

			expect(result).toHaveLength(1);
			expect(result[0].json).toEqual({
				sheets: [{ name: 'OnlySheet', id: 1, rowCount: 25, columnCount: 10 }],
			});
		});
	});

	describe('error handling', () => {
		it('throws when loadWorkbook fails', async () => {
			const mockFunctions = createMockExecuteFunctions();
			const context = createMockContext({ operation: 'getSheets' });

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
	});
});
