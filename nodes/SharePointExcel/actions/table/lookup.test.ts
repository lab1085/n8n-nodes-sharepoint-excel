import { describe, it, expect, vi, beforeEach } from 'vitest';
import { execute } from './lookup';
import { createMockExecuteFunctions, createMockContext } from '../../test-utils/mocks';

vi.mock('../../api', () => ({
	graphRequest: vi.fn(),
}));

import { graphRequest } from '../../api';

describe('lookup (table)', () => {
	beforeEach(() => {
		vi.clearAllMocks();
	});

	const mockColumnsResponse = (columns: string[]) => ({
		value: columns.map((name, index) => ({ name, index })),
	});

	const mockRowsResponse = (rows: (string | number | boolean | null)[][]) => ({
		value: rows.map((values, index) => ({
			index,
			values: [values],
		})),
	});

	describe('basic functionality', () => {
		it('returns matching rows with correct endpoint and structure', async () => {
			vi.mocked(graphRequest)
				.mockResolvedValueOnce(mockColumnsResponse(['Name', 'Email', 'Age']))
				.mockResolvedValueOnce(
					mockRowsResponse([
						['John', 'john@test.com', 30],
						['Jane', 'jane@test.com', 25],
						['John', 'john2@test.com', 35],
					]),
				);

			const mockFunctions = createMockExecuteFunctions({
				tableName: 'Table1',
				lookupColumn: 'Name',
				lookupValue: 'John',
			});
			const context = createMockContext({
				resource: 'table',
				operation: 'lookup',
				basePath: '/sites/site-id/drives/drive-id/items/file-id',
			});
			const result = await execute.call(mockFunctions, [], context);

			expect(graphRequest).toHaveBeenCalledWith(
				'GET',
				'/sites/site-id/drives/drive-id/items/file-id/workbook/tables/Table1/columns',
			);
			expect(graphRequest).toHaveBeenCalledWith(
				'GET',
				'/sites/site-id/drives/drive-id/items/file-id/workbook/tables/Table1/rows',
			);
			expect(result).toHaveLength(2);
			expect(result[0].json).toEqual({ Name: 'John', Email: 'john@test.com', Age: 30 });
			expect(result[1].json).toEqual({ Name: 'John', Email: 'john2@test.com', Age: 35 });
		});

		it('handles resourceLocator format for tableName', async () => {
			vi.mocked(graphRequest)
				.mockResolvedValueOnce(mockColumnsResponse(['Name']))
				.mockResolvedValueOnce(mockRowsResponse([['John']]));

			const mockFunctions = createMockExecuteFunctions({
				tableName: { mode: 'list', value: 'Table1' },
				lookupColumn: 'Name',
				lookupValue: 'John',
			});
			const context = createMockContext({ resource: 'table', operation: 'lookup' });
			await execute.call(mockFunctions, [], context);

			expect(graphRequest).toHaveBeenCalledWith(
				'GET',
				expect.stringContaining('/workbook/tables/Table1/columns'),
			);
		});
	});

	describe('lookup matching behavior', () => {
		it('compares values as strings', async () => {
			vi.mocked(graphRequest)
				.mockResolvedValueOnce(mockColumnsResponse(['ID', 'Name']))
				.mockResolvedValueOnce(
					mockRowsResponse([
						[123, 'John'],
						[456, 'Jane'],
					]),
				);

			const mockFunctions = createMockExecuteFunctions({
				tableName: 'Table1',
				lookupColumn: 'ID',
				lookupValue: '123', // String lookup for numeric column
			});
			const context = createMockContext({ resource: 'table', operation: 'lookup' });
			const result = await execute.call(mockFunctions, [], context);

			expect(result).toHaveLength(1);
			expect(result[0].json).toEqual({ ID: 123, Name: 'John' });
		});
	});

	describe('edge cases', () => {
		it('returns message when no matches found', async () => {
			vi.mocked(graphRequest)
				.mockResolvedValueOnce(mockColumnsResponse(['Name']))
				.mockResolvedValueOnce(mockRowsResponse([['John'], ['Jane']]));

			const mockFunctions = createMockExecuteFunctions({
				tableName: 'Table1',
				lookupColumn: 'Name',
				lookupValue: 'Bob',
			});
			const context = createMockContext({ resource: 'table', operation: 'lookup' });
			const result = await execute.call(mockFunctions, [], context);

			expect(result).toHaveLength(1);
			expect(result[0].json).toEqual({
				message: 'No rows found where "Name" equals "Bob"',
			});
		});
	});

	describe('error handling', () => {
		it('throws error when lookup column not found', async () => {
			vi.mocked(graphRequest).mockResolvedValueOnce(mockColumnsResponse(['Name', 'Email']));

			const mockFunctions = createMockExecuteFunctions({
				tableName: 'Table1',
				lookupColumn: 'NonExistent',
				lookupValue: 'test',
			});
			const context = createMockContext({ resource: 'table', operation: 'lookup' });

			await expect(execute.call(mockFunctions, [], context)).rejects.toThrow(
				'Column "NonExistent" not found in table "Table1"',
			);
		});

		it('throws error when graphRequest fails', async () => {
			vi.mocked(graphRequest).mockRejectedValueOnce(new Error('Graph API request failed'));

			const mockFunctions = createMockExecuteFunctions({
				tableName: 'Table1',
				lookupColumn: 'Name',
				lookupValue: 'John',
			});
			const context = createMockContext({ resource: 'table', operation: 'lookup' });

			await expect(execute.call(mockFunctions, [], context)).rejects.toThrow(
				'Graph API request failed',
			);
		});
	});
});
