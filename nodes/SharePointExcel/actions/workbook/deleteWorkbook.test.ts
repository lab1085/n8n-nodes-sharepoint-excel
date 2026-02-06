import { describe, it, expect, vi, beforeEach } from 'vitest';
import { NodeOperationError } from 'n8n-workflow';
import { execute } from './deleteWorkbook';
import { createMockExecuteFunctions, createMockContext } from '../../test-utils/mocks';

vi.mock('../../api', () => ({
	deleteFile: vi.fn(),
}));

import { deleteFile } from '../../api';

describe('deleteWorkbook', () => {
	beforeEach(() => {
		vi.clearAllMocks();
	});

	describe('basic functionality', () => {
		it('deletes file and returns success response', async () => {
			vi.mocked(deleteFile).mockResolvedValue(undefined);

			const mockFunctions = createMockExecuteFunctions();
			const context = createMockContext({
				resource: 'workbook',
				operation: 'deleteWorkbook',
				basePath: '/sites/my-site/drives/my-drive/items/my-file',
				fileId: 'file-123',
			});

			const result = await execute.call(mockFunctions, [{ json: {} }], context);

			expect(result).toHaveLength(1);
			expect(result[0].json).toEqual({
				success: true,
				deleted: true,
				fileId: 'file-123',
			});
			expect(deleteFile).toHaveBeenCalledWith(context.basePath);
		});
	});

	describe('edge cases', () => {
		it('ignores input items (operation is workbook-level)', async () => {
			vi.mocked(deleteFile).mockResolvedValue(undefined);

			const mockFunctions = createMockExecuteFunctions();
			const context = createMockContext({
				resource: 'workbook',
				operation: 'deleteWorkbook',
				fileId: 'file-456',
			});

			const multipleItems = [{ json: { a: 1 } }, { json: { b: 2 } }, { json: { c: 3 } }];
			const result = await execute.call(mockFunctions, multipleItems, context);

			expect(result).toHaveLength(1);
			expect(deleteFile).toHaveBeenCalledTimes(1);
		});
	});

	describe('error handling', () => {
		it('throws error when deleteFile fails', async () => {
			const mockFunctions = createMockExecuteFunctions();
			const context = createMockContext({
				resource: 'workbook',
				operation: 'deleteWorkbook',
			});

			vi.mocked(deleteFile).mockRejectedValue(
				new NodeOperationError(mockFunctions.getNode(), 'Graph API request failed: File not found'),
			);

			await expect(execute.call(mockFunctions, [{ json: {} }], context)).rejects.toThrow(
				'Graph API request failed: File not found',
			);
		});
	});
});
