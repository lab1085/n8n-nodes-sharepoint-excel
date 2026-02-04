import { describe, it, expect, vi, beforeEach } from 'vitest';
import { NodeOperationError } from 'n8n-workflow';
import { execute } from './deleteSheet';
import { createMockExecuteFunctions, createMockContext } from '../../test-utils/mocks';

// 1. Module mocks
vi.mock('../../api', () => ({
	loadWorkbook: vi.fn(),
	saveWorkbook: vi.fn(),
}));

import { loadWorkbook, saveWorkbook } from '../../api';

// 2. Local helpers for deleteSheet-specific mocks
interface MockWorksheetSimple {
	id: number;
	name: string;
}

function createWorkbookForDelete(sheets: MockWorksheetSimple[]) {
	const worksheets = [...sheets];

	return {
		worksheets,
		getWorksheet: vi.fn((name: string) => worksheets.find((ws) => ws.name === name)),
		removeWorksheet: vi.fn((id: number) => {
			const index = worksheets.findIndex((ws) => ws.id === id);
			if (index !== -1) {
				worksheets.splice(index, 1);
			}
		}),
	};
}

// 3. Test suite
describe('deleteSheet', () => {
	beforeEach(() => {
		vi.clearAllMocks();
	});

	describe('basic functionality', () => {
		it('deletes sheet and saves workbook', async () => {
			const workbook = createWorkbookForDelete([
				{ id: 1, name: 'Sheet1' },
				{ id: 2, name: 'Sheet2' },
			]);
			const mockFunctions = createMockExecuteFunctions({ sheetName: 'Sheet1' });
			const context = createMockContext({ operation: 'deleteSheet' });

			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);
			vi.mocked(saveWorkbook).mockResolvedValue(undefined);

			const result = await execute.call(mockFunctions, [{ json: {} }], context);

			expect(loadWorkbook).toHaveBeenCalledTimes(1);
			expect(workbook.removeWorksheet).toHaveBeenCalledWith(1);
			expect(saveWorkbook).toHaveBeenCalledTimes(1);
			expect(result).toHaveLength(1);
			expect(result[0].json).toEqual({
				success: true,
				deletedSheet: 'Sheet1',
				remainingSheets: ['Sheet2'],
			});
		});

		it('handles resourceLocator format for sheetName', async () => {
			const workbook = createWorkbookForDelete([
				{ id: 1, name: 'DataSheet' },
				{ id: 2, name: 'Other' },
			]);
			const mockFunctions = createMockExecuteFunctions({
				sheetName: { mode: 'list', value: 'DataSheet' },
			});
			const context = createMockContext({ operation: 'deleteSheet' });

			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);
			vi.mocked(saveWorkbook).mockResolvedValue(undefined);

			const result = await execute.call(mockFunctions, [{ json: {} }], context);

			expect(workbook.getWorksheet).toHaveBeenCalledWith('DataSheet');
			expect(result[0].json.deletedSheet).toBe('DataSheet');
		});

		it('uses correct basePath from context', async () => {
			const workbook = createWorkbookForDelete([
				{ id: 1, name: 'Sheet1' },
				{ id: 2, name: 'Sheet2' },
			]);
			const mockFunctions = createMockExecuteFunctions({ sheetName: 'Sheet1' });
			const customBasePath = '/sites/custom-site/drives/custom-drive/items/custom-file';
			const context = createMockContext({
				operation: 'deleteSheet',
				basePath: customBasePath,
			});

			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);
			vi.mocked(saveWorkbook).mockResolvedValue(undefined);

			await execute.call(mockFunctions, [{ json: {} }], context);

			expect(loadWorkbook).toHaveBeenCalledWith(customBasePath);
			expect(saveWorkbook).toHaveBeenCalledWith(customBasePath, workbook);
		});
	});

	describe('edge cases', () => {
		it('ignores input items (operation is sheet-level)', async () => {
			const workbook = createWorkbookForDelete([
				{ id: 1, name: 'Sheet1' },
				{ id: 2, name: 'Sheet2' },
			]);
			const mockFunctions = createMockExecuteFunctions({ sheetName: 'Sheet1' });
			const context = createMockContext({ operation: 'deleteSheet' });

			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);
			vi.mocked(saveWorkbook).mockResolvedValue(undefined);

			// Pass multiple input items
			const multipleItems = [{ json: { a: 1 } }, { json: { b: 2 } }, { json: { c: 3 } }];
			const result = await execute.call(mockFunctions, multipleItems, context);

			// Should return single result regardless of input count
			expect(result).toHaveLength(1);
			expect(result[0].json.success).toBe(true);
		});
	});

	describe('error handling', () => {
		it('throws error when sheet not found', async () => {
			const workbook = createWorkbookForDelete([
				{ id: 1, name: 'Sheet1' },
				{ id: 2, name: 'Sheet2' },
			]);
			const mockFunctions = createMockExecuteFunctions({ sheetName: 'NonExistent' });
			const context = createMockContext({ operation: 'deleteSheet' });

			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);

			await expect(execute.call(mockFunctions, [{ json: {} }], context)).rejects.toThrow(
				'Sheet "NonExistent" not found in workbook',
			);

			expect(saveWorkbook).not.toHaveBeenCalled();
		});

		it('throws error when trying to delete the last sheet', async () => {
			const workbook = createWorkbookForDelete([{ id: 1, name: 'OnlySheet' }]);
			const mockFunctions = createMockExecuteFunctions({ sheetName: 'OnlySheet' });
			const context = createMockContext({ operation: 'deleteSheet' });

			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);

			await expect(execute.call(mockFunctions, [{ json: {} }], context)).rejects.toThrow(
				'Cannot delete the last sheet in a workbook',
			);

			expect(workbook.removeWorksheet).not.toHaveBeenCalled();
			expect(saveWorkbook).not.toHaveBeenCalled();
		});

		it('throws error when loadWorkbook fails', async () => {
			const mockFunctions = createMockExecuteFunctions({ sheetName: 'Sheet1' });
			const context = createMockContext({ operation: 'deleteSheet' });

			vi.mocked(loadWorkbook).mockRejectedValue(
				new NodeOperationError(mockFunctions.getNode(), 'Graph API request failed: Access denied'),
			);

			await expect(execute.call(mockFunctions, [{ json: {} }], context)).rejects.toThrow(
				'Graph API request failed: Access denied',
			);

			expect(saveWorkbook).not.toHaveBeenCalled();
		});

		it('throws error when saveWorkbook fails', async () => {
			const workbook = createWorkbookForDelete([
				{ id: 1, name: 'Sheet1' },
				{ id: 2, name: 'Sheet2' },
			]);
			const mockFunctions = createMockExecuteFunctions({ sheetName: 'Sheet1' });
			const context = createMockContext({ operation: 'deleteSheet' });

			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);
			vi.mocked(saveWorkbook).mockRejectedValue(
				new NodeOperationError(mockFunctions.getNode(), 'Failed to upload workbook'),
			);

			await expect(execute.call(mockFunctions, [{ json: {} }], context)).rejects.toThrow(
				'Failed to upload workbook',
			);

			expect(workbook.removeWorksheet).toHaveBeenCalled();
		});
	});
});
