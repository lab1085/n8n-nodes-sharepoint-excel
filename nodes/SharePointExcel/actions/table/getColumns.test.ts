import { describe, it, expect, vi, beforeEach } from 'vitest';
import { execute } from './getColumns';
import { createMockExecuteFunctions, createMockContext } from '../../test-utils/mocks';

vi.mock('../../api', () => ({
	graphRequest: vi.fn(),
}));

import { graphRequest } from '../../api';

describe('getColumns (table)', () => {
	beforeEach(() => {
		vi.clearAllMocks();
	});

	const mockColumnsResponse = (columns: Array<{ id: string; name: string; index: number }>) => ({
		value: columns,
	});

	describe('basic functionality', () => {
		it('returns columns with correct endpoint and structure', async () => {
			vi.mocked(graphRequest).mockResolvedValueOnce(
				mockColumnsResponse([
					{ id: 'col1', name: 'Name', index: 0 },
					{ id: 'col2', name: 'Email', index: 1 },
					{ id: 'col3', name: 'Age', index: 2 },
				]),
			);

			const mockFunctions = createMockExecuteFunctions({ tableName: 'Table1' });
			const context = createMockContext({
				resource: 'table',
				operation: 'getColumns',
				basePath: '/sites/site-id/drives/drive-id/items/file-id',
			});
			const result = await execute.call(mockFunctions, [], context);

			expect(graphRequest).toHaveBeenCalledWith(
				'GET',
				'/sites/site-id/drives/drive-id/items/file-id/workbook/tables/Table1/columns',
			);
			expect(result).toHaveLength(3);
			expect(result[0].json).toEqual({ id: 'col1', name: 'Name', index: 0 });
			expect(result[1].json).toEqual({ id: 'col2', name: 'Email', index: 1 });
			expect(result[2].json).toEqual({ id: 'col3', name: 'Age', index: 2 });
		});

		it('handles resourceLocator format for tableName', async () => {
			vi.mocked(graphRequest).mockResolvedValueOnce(
				mockColumnsResponse([{ id: 'col1', name: 'Name', index: 0 }]),
			);

			const mockFunctions = createMockExecuteFunctions({
				tableName: { mode: 'list', value: 'Table1' },
			});
			const context = createMockContext({ resource: 'table', operation: 'getColumns' });
			await execute.call(mockFunctions, [], context);

			expect(graphRequest).toHaveBeenCalledWith(
				'GET',
				expect.stringContaining('/workbook/tables/Table1/columns'),
			);
		});
	});

	describe('edge cases', () => {
		it('returns message when no columns found', async () => {
			vi.mocked(graphRequest).mockResolvedValueOnce(mockColumnsResponse([]));

			const mockFunctions = createMockExecuteFunctions({ tableName: 'Table1' });
			const context = createMockContext({ resource: 'table', operation: 'getColumns' });
			const result = await execute.call(mockFunctions, [], context);

			expect(result).toHaveLength(1);
			expect(result[0].json).toEqual({ message: 'No columns found in table' });
		});
	});

	describe('error handling', () => {
		it('throws error when graphRequest fails', async () => {
			vi.mocked(graphRequest).mockRejectedValueOnce(new Error('Graph API request failed'));

			const mockFunctions = createMockExecuteFunctions({ tableName: 'Table1' });
			const context = createMockContext({ resource: 'table', operation: 'getColumns' });

			await expect(execute.call(mockFunctions, [], context)).rejects.toThrow(
				'Graph API request failed',
			);
		});
	});
});
