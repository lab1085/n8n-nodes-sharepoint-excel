import { describe, it, expect, vi, beforeEach } from 'vitest';
import { NodeOperationError } from 'n8n-workflow';
import { execute } from './getWorkbooks';
import { createMockExecuteFunctions, createMockContext } from '../../test-utils/mocks';

vi.mock('../../api', () => ({
	graphRequest: vi.fn(),
}));

import { graphRequest } from '../../api';

describe('getWorkbooks', () => {
	beforeEach(() => {
		vi.clearAllMocks();
	});

	describe('basic functionality', () => {
		it('returns Excel files with id and name', async () => {
			vi.mocked(graphRequest).mockResolvedValue({
				value: [
					{ id: 'file-1', name: 'Report.xlsx', file: {} },
					{ id: 'file-2', name: 'Data.xlsx', file: {} },
				],
			});

			const mockFunctions = createMockExecuteFunctions();
			const context = createMockContext({
				resource: 'workbook',
				operation: 'getWorkbooks',
			});

			const result = await execute.call(mockFunctions, [{ json: {} }], context);

			expect(result).toHaveLength(2);
			expect(result[0].json).toEqual({ id: 'file-1', name: 'Report.xlsx' });
			expect(result[1].json).toEqual({ id: 'file-2', name: 'Data.xlsx' });
		});

		it('uses correct endpoint with driveId from context', async () => {
			vi.mocked(graphRequest).mockResolvedValue({ value: [] });

			const mockFunctions = createMockExecuteFunctions();
			const context = createMockContext({
				resource: 'workbook',
				operation: 'getWorkbooks',
				driveId: 'custom-drive-id',
			});

			await execute.call(mockFunctions, [{ json: {} }], context);

			expect(graphRequest).toHaveBeenCalledWith('GET', '/drives/custom-drive-id/root/children');
		});
	});

	describe('filtering behavior', () => {
		it('filters out folders (items without file property)', async () => {
			vi.mocked(graphRequest).mockResolvedValue({
				value: [
					{ id: 'folder-1', name: 'Documents' }, // No file property - folder
					{ id: 'file-1', name: 'Data.xlsx', file: {} },
				],
			});

			const mockFunctions = createMockExecuteFunctions();
			const context = createMockContext({
				resource: 'workbook',
				operation: 'getWorkbooks',
			});

			const result = await execute.call(mockFunctions, [{ json: {} }], context);

			expect(result).toHaveLength(1);
			expect(result[0].json).toEqual({ id: 'file-1', name: 'Data.xlsx' });
		});

		it('filters out non-xlsx files', async () => {
			vi.mocked(graphRequest).mockResolvedValue({
				value: [
					{ id: 'file-1', name: 'document.pdf', file: {} },
					{ id: 'file-2', name: 'spreadsheet.xlsx', file: {} },
					{ id: 'file-3', name: 'word.docx', file: {} },
				],
			});

			const mockFunctions = createMockExecuteFunctions();
			const context = createMockContext({
				resource: 'workbook',
				operation: 'getWorkbooks',
			});

			const result = await execute.call(mockFunctions, [{ json: {} }], context);

			expect(result).toHaveLength(1);
			expect(result[0].json).toEqual({ id: 'file-2', name: 'spreadsheet.xlsx' });
		});

		it('handles case-insensitive .xlsx extension', async () => {
			vi.mocked(graphRequest).mockResolvedValue({
				value: [
					{ id: 'file-1', name: 'report.XLSX', file: {} },
					{ id: 'file-2', name: 'data.Xlsx', file: {} },
					{ id: 'file-3', name: 'other.XlSx', file: {} },
				],
			});

			const mockFunctions = createMockExecuteFunctions();
			const context = createMockContext({
				resource: 'workbook',
				operation: 'getWorkbooks',
			});

			const result = await execute.call(mockFunctions, [{ json: {} }], context);

			expect(result).toHaveLength(3);
			expect(result[0].json.name).toBe('report.XLSX');
			expect(result[1].json.name).toBe('data.Xlsx');
			expect(result[2].json.name).toBe('other.XlSx');
		});
	});

	describe('edge cases', () => {
		it('returns message when no Excel files found', async () => {
			vi.mocked(graphRequest).mockResolvedValue({
				value: [
					{ id: 'file-1', name: 'document.pdf', file: {} },
					{ id: 'folder-1', name: 'Folder' },
				],
			});

			const mockFunctions = createMockExecuteFunctions();
			const context = createMockContext({
				resource: 'workbook',
				operation: 'getWorkbooks',
			});

			const result = await execute.call(mockFunctions, [{ json: {} }], context);

			expect(result).toHaveLength(1);
			expect(result[0].json).toEqual({ message: 'No Excel files found in drive' });
		});

		it('handles empty drive (no files at all)', async () => {
			vi.mocked(graphRequest).mockResolvedValue({ value: [] });

			const mockFunctions = createMockExecuteFunctions();
			const context = createMockContext({
				resource: 'workbook',
				operation: 'getWorkbooks',
			});

			const result = await execute.call(mockFunctions, [{ json: {} }], context);

			expect(result).toHaveLength(1);
			expect(result[0].json).toEqual({ message: 'No Excel files found in drive' });
		});

		it('handles response with missing value property', async () => {
			vi.mocked(graphRequest).mockResolvedValue({});

			const mockFunctions = createMockExecuteFunctions();
			const context = createMockContext({
				resource: 'workbook',
				operation: 'getWorkbooks',
			});

			const result = await execute.call(mockFunctions, [{ json: {} }], context);

			expect(result).toHaveLength(1);
			expect(result[0].json).toEqual({ message: 'No Excel files found in drive' });
		});
	});

	describe('error handling', () => {
		it('throws when graphRequest fails', async () => {
			const mockFunctions = createMockExecuteFunctions();
			const context = createMockContext({
				resource: 'workbook',
				operation: 'getWorkbooks',
			});

			vi.mocked(graphRequest).mockRejectedValue(
				new NodeOperationError(mockFunctions.getNode(), 'Graph API request failed: Access denied'),
			);

			await expect(execute.call(mockFunctions, [{ json: {} }], context)).rejects.toThrow(
				'Graph API request failed: Access denied',
			);
		});
	});
});
