import { describe, it, expect, vi, beforeEach } from 'vitest';
import { execute } from './getRows';
import { createMockExecuteFunctions, createMockContext } from '../../test-utils/mocks';

// Mock the api module
vi.mock('../../api', () => ({
	graphRequest: vi.fn(),
}));

import { graphRequest } from '../../api';

describe('getRows (table)', () => {
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
		it('returns keyed objects with all columns by default', async () => {
			vi.mocked(graphRequest)
				.mockResolvedValueOnce(mockColumnsResponse(['Name', 'Email', 'Age']))
				.mockResolvedValueOnce(
					mockRowsResponse([
						['John', 'john@test.com', 30],
						['Jane', 'jane@test.com', 25],
					]),
				);

			const mockFunctions = createMockExecuteFunctions({
				tableName: 'Table1',
				returnAll: true,
				options: {},
			});

			const context = createMockContext({ resource: 'table', operation: 'getTableRows' });
			const result = await execute.call(mockFunctions, [], context);

			expect(result).toHaveLength(2);
			expect(result[0].json).toEqual({ Name: 'John', Email: 'john@test.com', Age: 30 });
			expect(result[1].json).toEqual({ Name: 'Jane', Email: 'jane@test.com', Age: 25 });
		});

		it('handles resourceLocator format for tableName', async () => {
			vi.mocked(graphRequest)
				.mockResolvedValueOnce(mockColumnsResponse(['Name']))
				.mockResolvedValueOnce(mockRowsResponse([['John']]));

			const mockFunctions = createMockExecuteFunctions({
				tableName: { mode: 'list', value: 'Table1' },
				returnAll: true,
				options: {},
			});

			const context = createMockContext({ resource: 'table', operation: 'getTableRows' });
			await execute.call(mockFunctions, [], context);

			expect(graphRequest).toHaveBeenCalledWith(
				'GET',
				expect.stringContaining('/workbook/tables/Table1/columns'),
			);
		});
	});

	describe('returnAll and limit', () => {
		it('returns all rows when returnAll is true', async () => {
			vi.mocked(graphRequest)
				.mockResolvedValueOnce(mockColumnsResponse(['Name']))
				.mockResolvedValueOnce(mockRowsResponse([['A'], ['B'], ['C'], ['D'], ['E']]));

			const mockFunctions = createMockExecuteFunctions({
				tableName: 'Table1',
				returnAll: true,
				options: {},
			});

			const context = createMockContext({ resource: 'table', operation: 'getTableRows' });
			const result = await execute.call(mockFunctions, [], context);

			expect(result).toHaveLength(5);
		});

		it('limits rows when returnAll is false', async () => {
			vi.mocked(graphRequest)
				.mockResolvedValueOnce(mockColumnsResponse(['Name']))
				.mockResolvedValueOnce(mockRowsResponse([['A'], ['B'], ['C'], ['D'], ['E']]));

			const mockFunctions = createMockExecuteFunctions({
				tableName: 'Table1',
				returnAll: false,
				limit: 2,
				options: {},
			});

			const context = createMockContext({ resource: 'table', operation: 'getTableRows' });
			const result = await execute.call(mockFunctions, [], context);

			expect(result).toHaveLength(2);
			expect(result[0].json).toEqual({ Name: 'A' });
			expect(result[1].json).toEqual({ Name: 'B' });
		});
	});

	describe('rawData mode', () => {
		it('returns arrays instead of objects when rawData is true', async () => {
			vi.mocked(graphRequest)
				.mockResolvedValueOnce(mockColumnsResponse(['Name', 'Email']))
				.mockResolvedValueOnce(
					mockRowsResponse([
						['John', 'john@test.com'],
						['Jane', 'jane@test.com'],
					]),
				);

			const mockFunctions = createMockExecuteFunctions({
				tableName: 'Table1',
				returnAll: true,
				options: { rawData: true },
			});

			const context = createMockContext({ resource: 'table', operation: 'getTableRows' });
			const result = await execute.call(mockFunctions, [], context);

			expect(result).toHaveLength(1);
			expect(result[0].json).toEqual({
				headers: ['Name', 'Email'],
				data: [
					['John', 'john@test.com'],
					['Jane', 'jane@test.com'],
				],
			});
		});

		it('uses custom dataProperty name', async () => {
			vi.mocked(graphRequest)
				.mockResolvedValueOnce(mockColumnsResponse(['Name']))
				.mockResolvedValueOnce(mockRowsResponse([['John']]));

			const mockFunctions = createMockExecuteFunctions({
				tableName: 'Table1',
				returnAll: true,
				options: { rawData: true, dataProperty: 'rows' },
			});

			const context = createMockContext({ resource: 'table', operation: 'getTableRows' });
			const result = await execute.call(mockFunctions, [], context);

			expect(result[0].json).toHaveProperty('rows');
			expect(result[0].json).not.toHaveProperty('data');
		});
	});

	describe('fields filter', () => {
		it('returns only specified columns', async () => {
			vi.mocked(graphRequest)
				.mockResolvedValueOnce(mockColumnsResponse(['Name', 'Email', 'Age', 'City']))
				.mockResolvedValueOnce(mockRowsResponse([['John', 'john@test.com', 30, 'NYC']]));

			const mockFunctions = createMockExecuteFunctions({
				tableName: 'Table1',
				returnAll: true,
				options: { fields: 'Name, Age' },
			});

			const context = createMockContext({ resource: 'table', operation: 'getTableRows' });
			const result = await execute.call(mockFunctions, [], context);

			expect(result[0].json).toEqual({ Name: 'John', Age: 30 });
			expect(result[0].json).not.toHaveProperty('Email');
			expect(result[0].json).not.toHaveProperty('City');
		});

		it('works with rawData mode and fields filter', async () => {
			vi.mocked(graphRequest)
				.mockResolvedValueOnce(mockColumnsResponse(['Name', 'Email', 'Age']))
				.mockResolvedValueOnce(mockRowsResponse([['John', 'john@test.com', 30]]));

			const mockFunctions = createMockExecuteFunctions({
				tableName: 'Table1',
				returnAll: true,
				options: { rawData: true, fields: 'Name, Age' },
			});

			const context = createMockContext({ resource: 'table', operation: 'getTableRows' });
			const result = await execute.call(mockFunctions, [], context);

			expect(result[0].json).toEqual({
				headers: ['Name', 'Age'],
				data: [['John', 30]],
			});
		});
	});

	describe('hasData filter', () => {
		it('skips rows with no data in normal mode', async () => {
			vi.mocked(graphRequest)
				.mockResolvedValueOnce(mockColumnsResponse(['Name', 'Email']))
				.mockResolvedValueOnce(
					mockRowsResponse([
						['John', 'john@test.com'],
						[null, null],
						['', ''],
						['Jane', 'jane@test.com'],
					]),
				);

			const mockFunctions = createMockExecuteFunctions({
				tableName: 'Table1',
				returnAll: true,
				options: {},
			});

			const context = createMockContext({ resource: 'table', operation: 'getTableRows' });
			const result = await execute.call(mockFunctions, [], context);

			expect(result).toHaveLength(2);
			expect(result[0].json).toEqual({ Name: 'John', Email: 'john@test.com' });
			expect(result[1].json).toEqual({ Name: 'Jane', Email: 'jane@test.com' });
		});

		it('skips rows with no data in rawData mode', async () => {
			vi.mocked(graphRequest)
				.mockResolvedValueOnce(mockColumnsResponse(['Name', 'Email']))
				.mockResolvedValueOnce(
					mockRowsResponse([
						['John', 'john@test.com'],
						[null, null],
						['Jane', 'jane@test.com'],
					]),
				);

			const mockFunctions = createMockExecuteFunctions({
				tableName: 'Table1',
				returnAll: true,
				options: { rawData: true },
			});

			const context = createMockContext({ resource: 'table', operation: 'getTableRows' });
			const result = await execute.call(mockFunctions, [], context);

			expect(result[0].json.data).toHaveLength(2);
		});
	});

	describe('edge cases', () => {
		it('returns message when no rows found', async () => {
			vi.mocked(graphRequest)
				.mockResolvedValueOnce(mockColumnsResponse(['Name']))
				.mockResolvedValueOnce(mockRowsResponse([]));

			const mockFunctions = createMockExecuteFunctions({
				tableName: 'Table1',
				returnAll: true,
				options: {},
			});

			const context = createMockContext({ resource: 'table', operation: 'getTableRows' });
			const result = await execute.call(mockFunctions, [], context);

			expect(result).toHaveLength(1);
			expect(result[0].json).toEqual({ message: 'No rows found in table' });
		});

		it('handles null values correctly', async () => {
			vi.mocked(graphRequest)
				.mockResolvedValueOnce(mockColumnsResponse(['Name', 'Email']))
				.mockResolvedValueOnce(mockRowsResponse([['John', null]]));

			const mockFunctions = createMockExecuteFunctions({
				tableName: 'Table1',
				returnAll: true,
				options: {},
			});

			const context = createMockContext({ resource: 'table', operation: 'getTableRows' });
			const result = await execute.call(mockFunctions, [], context);

			expect(result[0].json).toEqual({ Name: 'John', Email: null });
		});
	});
});
