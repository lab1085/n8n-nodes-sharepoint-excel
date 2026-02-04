import { describe, it, expect, vi, beforeEach } from 'vitest';
import { execute } from './readRows';
import {
	createMockExecuteFunctions,
	createMockContext,
	createMockWorksheet,
	createMockWorkbook,
} from '../../test-utils/mocks';

// Mock the api module
vi.mock('../../api', () => ({
	loadWorkbook: vi.fn(),
	getWorksheet: vi.fn(),
}));

import { loadWorkbook, getWorksheet } from '../../api';

describe('readRows', () => {
	beforeEach(() => {
		vi.clearAllMocks();
	});

	describe('basic functionality', () => {
		it('returns keyed objects with all columns by default', async () => {
			const worksheet = createMockWorksheet({
				headers: { 1: 'Name', 2: 'Email', 3: 'Age' },
				rows: [
					{ 1: 'John', 2: 'john@test.com', 3: 30 },
					{ 1: 'Jane', 2: 'jane@test.com', 3: 25 },
				],
			});
			const workbook = createMockWorkbook({ Sheet1: worksheet });

			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);
			vi.mocked(getWorksheet).mockReturnValue(worksheet as never);

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				returnAll: true,
				options: {},
			});

			const context = createMockContext();
			const result = await execute.call(mockFunctions, [], context);

			expect(result).toHaveLength(2);
			expect(result[0].json).toEqual({ Name: 'John', Email: 'john@test.com', Age: 30 });
			expect(result[1].json).toEqual({ Name: 'Jane', Email: 'jane@test.com', Age: 25 });
		});

		it('handles resourceLocator format for sheetName', async () => {
			const worksheet = createMockWorksheet({
				headers: { 1: 'Name' },
				rows: [{ 1: 'John' }],
			});
			const workbook = createMockWorkbook({ Sheet1: worksheet });

			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);
			vi.mocked(getWorksheet).mockReturnValue(worksheet as never);

			const mockFunctions = createMockExecuteFunctions({
				sheetName: { mode: 'list', value: 'Sheet1' },
				returnAll: true,
				options: {},
			});

			const context = createMockContext();
			await execute.call(mockFunctions, [], context);

			expect(getWorksheet).toHaveBeenCalledWith(workbook, 'Sheet1', mockFunctions);
		});
	});

	describe('returnAll and limit', () => {
		it('returns all rows when returnAll is true', async () => {
			const worksheet = createMockWorksheet({
				headers: { 1: 'Name' },
				rows: [{ 1: 'A' }, { 1: 'B' }, { 1: 'C' }, { 1: 'D' }, { 1: 'E' }],
			});
			const workbook = createMockWorkbook({ Sheet1: worksheet });

			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);
			vi.mocked(getWorksheet).mockReturnValue(worksheet as never);

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				returnAll: true,
				options: {},
			});

			const context = createMockContext();
			const result = await execute.call(mockFunctions, [], context);

			expect(result).toHaveLength(5);
		});

		it('limits rows when returnAll is false', async () => {
			const worksheet = createMockWorksheet({
				headers: { 1: 'Name' },
				rows: [{ 1: 'A' }, { 1: 'B' }, { 1: 'C' }, { 1: 'D' }, { 1: 'E' }],
			});
			const workbook = createMockWorkbook({ Sheet1: worksheet });

			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);
			vi.mocked(getWorksheet).mockReturnValue(worksheet as never);

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				returnAll: false,
				limit: 2,
				options: {},
			});

			const context = createMockContext();
			const result = await execute.call(mockFunctions, [], context);

			expect(result).toHaveLength(2);
			expect(result[0].json).toEqual({ Name: 'A' });
			expect(result[1].json).toEqual({ Name: 'B' });
		});
	});

	describe('rawData mode', () => {
		it('returns arrays instead of objects when rawData is true', async () => {
			const worksheet = createMockWorksheet({
				headers: { 1: 'Name', 2: 'Email' },
				rows: [
					{ 1: 'John', 2: 'john@test.com' },
					{ 1: 'Jane', 2: 'jane@test.com' },
				],
			});
			const workbook = createMockWorkbook({ Sheet1: worksheet });

			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);
			vi.mocked(getWorksheet).mockReturnValue(worksheet as never);

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				returnAll: true,
				options: { rawData: true },
			});

			const context = createMockContext();
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
			const worksheet = createMockWorksheet({
				headers: { 1: 'Name' },
				rows: [{ 1: 'John' }],
			});
			const workbook = createMockWorkbook({ Sheet1: worksheet });

			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);
			vi.mocked(getWorksheet).mockReturnValue(worksheet as never);

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				returnAll: true,
				options: { rawData: true, dataProperty: 'rows' },
			});

			const context = createMockContext();
			const result = await execute.call(mockFunctions, [], context);

			expect(result[0].json).toHaveProperty('rows');
			expect(result[0].json).not.toHaveProperty('data');
		});
	});

	describe('fields filter', () => {
		it('returns only specified columns', async () => {
			const worksheet = createMockWorksheet({
				headers: { 1: 'Name', 2: 'Email', 3: 'Age', 4: 'City' },
				rows: [{ 1: 'John', 2: 'john@test.com', 3: 30, 4: 'NYC' }],
			});
			const workbook = createMockWorkbook({ Sheet1: worksheet });

			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);
			vi.mocked(getWorksheet).mockReturnValue(worksheet as never);

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				returnAll: true,
				options: { fields: 'Name, Age' },
			});

			const context = createMockContext();
			const result = await execute.call(mockFunctions, [], context);

			expect(result[0].json).toEqual({ Name: 'John', Age: 30 });
			expect(result[0].json).not.toHaveProperty('Email');
			expect(result[0].json).not.toHaveProperty('City');
		});

		it('works with rawData mode and fields filter', async () => {
			const worksheet = createMockWorksheet({
				headers: { 1: 'Name', 2: 'Email', 3: 'Age' },
				rows: [{ 1: 'John', 2: 'john@test.com', 3: 30 }],
			});
			const workbook = createMockWorkbook({ Sheet1: worksheet });

			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);
			vi.mocked(getWorksheet).mockReturnValue(worksheet as never);

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				returnAll: true,
				options: { rawData: true, fields: 'Name, Age' },
			});

			const context = createMockContext();
			const result = await execute.call(mockFunctions, [], context);

			expect(result[0].json).toEqual({
				headers: ['Name', 'Age'],
				data: [['John', 30]],
			});
		});
	});

	describe('edge cases', () => {
		it('returns message when no data found', async () => {
			const worksheet = createMockWorksheet({
				headers: { 1: 'Name' },
				rows: [],
			});
			const workbook = createMockWorkbook({ Sheet1: worksheet });

			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);
			vi.mocked(getWorksheet).mockReturnValue(worksheet as never);

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				returnAll: true,
				options: {},
			});

			const context = createMockContext();
			const result = await execute.call(mockFunctions, [], context);

			expect(result).toHaveLength(1);
			expect(result[0].json).toEqual({ message: 'No data found in sheet' });
		});

		it('uses custom headerRow and startRow', async () => {
			const worksheet = createMockWorksheet({
				headers: { 1: 'Name', 2: 'Email' },
				rows: [
					{ 1: 'Skip1', 2: 'skip1@test.com' },
					{ 1: 'Skip2', 2: 'skip2@test.com' },
					{ 1: 'John', 2: 'john@test.com' },
				],
			});
			// Override getRow to handle custom header row
			worksheet.getRow = vi.fn((rowNum: number) => {
				if (rowNum === 2) {
					return {
						eachCell: vi.fn((opts, cb) => {
							cb({ value: 'Name' }, 1);
							cb({ value: 'Email' }, 2);
						}),
						getCell: vi.fn((col) => ({ value: col === 1 ? 'Name' : 'Email' })),
					};
				}
				if (rowNum === 4) {
					return {
						eachCell: vi.fn((opts, cb) => {
							cb({ value: 'John' }, 1);
							cb({ value: 'john@test.com' }, 2);
						}),
						getCell: vi.fn((col) => ({ value: col === 1 ? 'John' : 'john@test.com' })),
					};
				}
				return {
					eachCell: vi.fn(),
					getCell: vi.fn(() => ({ value: null })),
				};
			});
			worksheet.rowCount = 4;

			const workbook = createMockWorkbook({ Sheet1: worksheet });

			vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);
			vi.mocked(getWorksheet).mockReturnValue(worksheet as never);

			const mockFunctions = createMockExecuteFunctions({
				sheetName: 'Sheet1',
				returnAll: true,
				options: { headerRow: 2, startRow: 4 },
			});

			const context = createMockContext();
			const result = await execute.call(mockFunctions, [], context);

			expect(result).toHaveLength(1);
			expect(result[0].json).toEqual({ Name: 'John', Email: 'john@test.com' });
		});
	});
});
