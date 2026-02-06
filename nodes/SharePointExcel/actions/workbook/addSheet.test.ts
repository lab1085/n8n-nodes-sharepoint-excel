import { describe, it, expect, vi, beforeEach } from 'vitest';
import { NodeOperationError } from 'n8n-workflow';
import { execute } from './addSheet';
import { createMockExecuteFunctions, createMockContext } from '../../test-utils/mocks';

vi.mock('../../api', () => ({
	loadWorkbook: vi.fn(),
	saveWorkbook: vi.fn(),
}));

import { loadWorkbook, saveWorkbook } from '../../api';

function createWorkbookMock(existingSheets: string[] = []) {
	const worksheets = existingSheets.map((name, index) => ({ name, id: index + 1 }));
	let nextId = worksheets.length + 1;

	return {
		getWorksheet: vi.fn((name: string) => worksheets.find((s) => s.name === name)),
		addWorksheet: vi.fn((name: string) => {
			const newSheet = { name, id: nextId++ };
			worksheets.push(newSheet);
			return newSheet;
		}),
		worksheets,
	};
}

describe('addSheet', () => {
	beforeEach(() => {
		vi.clearAllMocks();
	});

	describe('basic functionality', () => {
		it('adds a new sheet and saves the workbook', async () => {
			const workbook = createWorkbookMock(['Sheet1']);
			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);
			vi.mocked(saveWorkbook).mockResolvedValue(undefined);

			const mockFunctions = createMockExecuteFunctions({ newSheetName: 'NewSheet' });
			const context = createMockContext({
				resource: 'workbook',
				operation: 'addSheet',
			});

			const result = await execute.call(mockFunctions, [{ json: {} }], context);

			expect(result).toHaveLength(1);
			expect(result[0].json).toEqual({
				success: true,
				sheetName: 'NewSheet',
				sheetId: 2,
				totalSheets: 2,
			});
			expect(loadWorkbook).toHaveBeenCalledWith(context.basePath);
			expect(workbook.addWorksheet).toHaveBeenCalledWith('NewSheet');
			expect(saveWorkbook).toHaveBeenCalledWith(context.basePath, workbook);
			expect(workbook.addWorksheet).toHaveBeenCalledBefore(vi.mocked(saveWorkbook));
		});
	});

	describe('edge cases', () => {
		it('ignores input items (operation is workbook-level)', async () => {
			const workbook = createWorkbookMock([]);
			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);
			vi.mocked(saveWorkbook).mockResolvedValue(undefined);

			const mockFunctions = createMockExecuteFunctions({ newSheetName: 'NewSheet' });
			const context = createMockContext({
				resource: 'workbook',
				operation: 'addSheet',
			});

			const multipleItems = [{ json: { a: 1 } }, { json: { b: 2 } }, { json: { c: 3 } }];
			const result = await execute.call(mockFunctions, multipleItems, context);

			expect(result).toHaveLength(1);
			expect(workbook.addWorksheet).toHaveBeenCalledTimes(1);
		});
	});

	describe('validation', () => {
		it('throws error when sheet name is empty string', async () => {
			const mockFunctions = createMockExecuteFunctions({ newSheetName: '' });
			const context = createMockContext({
				resource: 'workbook',
				operation: 'addSheet',
			});

			await expect(execute.call(mockFunctions, [{ json: {} }], context)).rejects.toThrow(
				'Sheet name cannot be empty',
			);

			expect(loadWorkbook).not.toHaveBeenCalled();
		});

		it('throws error when sheet name is whitespace only', async () => {
			const mockFunctions = createMockExecuteFunctions({ newSheetName: '   ' });
			const context = createMockContext({
				resource: 'workbook',
				operation: 'addSheet',
			});

			await expect(execute.call(mockFunctions, [{ json: {} }], context)).rejects.toThrow(
				'Sheet name cannot be empty',
			);

			expect(loadWorkbook).not.toHaveBeenCalled();
		});

		it('throws error when sheet already exists', async () => {
			const workbook = createWorkbookMock(['ExistingSheet']);
			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);

			const mockFunctions = createMockExecuteFunctions({ newSheetName: 'ExistingSheet' });
			const context = createMockContext({
				resource: 'workbook',
				operation: 'addSheet',
			});

			await expect(execute.call(mockFunctions, [{ json: {} }], context)).rejects.toThrow(
				'Sheet "ExistingSheet" already exists in the workbook',
			);

			expect(workbook.addWorksheet).not.toHaveBeenCalled();
			expect(saveWorkbook).not.toHaveBeenCalled();
		});
	});

	describe('error handling', () => {
		it('throws error when loadWorkbook fails', async () => {
			const mockFunctions = createMockExecuteFunctions({ newSheetName: 'NewSheet' });
			const context = createMockContext({
				resource: 'workbook',
				operation: 'addSheet',
			});

			vi.mocked(loadWorkbook).mockRejectedValue(
				new NodeOperationError(mockFunctions.getNode(), 'Graph API request failed: Access denied'),
			);

			await expect(execute.call(mockFunctions, [{ json: {} }], context)).rejects.toThrow(
				'Graph API request failed: Access denied',
			);
		});

		it('throws error when saveWorkbook fails', async () => {
			const workbook = createWorkbookMock([]);
			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);

			const mockFunctions = createMockExecuteFunctions({ newSheetName: 'NewSheet' });
			const context = createMockContext({
				resource: 'workbook',
				operation: 'addSheet',
			});

			vi.mocked(saveWorkbook).mockRejectedValue(
				new NodeOperationError(mockFunctions.getNode(), 'Failed to upload workbook'),
			);

			await expect(execute.call(mockFunctions, [{ json: {} }], context)).rejects.toThrow(
				'Failed to upload workbook',
			);
		});
	});
});
