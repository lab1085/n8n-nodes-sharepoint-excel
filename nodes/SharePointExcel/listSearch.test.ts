import { describe, it, expect, vi, beforeEach } from 'vitest';
import { searchSites, getDrives, getFiles, getSheets, getTables } from './listSearch';
import { createMockLoadOptionsFunctions } from './test-utils/mocks';

// Mock ExcelJS with a factory that returns a class
const mockWorkbookInstance = {
	xlsx: { load: vi.fn() },
	worksheets: [] as { name: string }[],
};

vi.mock('exceljs', () => ({
	Workbook: class MockWorkbook {
		xlsx = mockWorkbookInstance.xlsx;
		get worksheets() {
			return mockWorkbookInstance.worksheets;
		}
	},
}));

describe('listSearch', () => {
	beforeEach(() => {
		vi.clearAllMocks();
	});

	describe('searchSites', () => {
		describe('basic functionality', () => {
			it('returns sites from Graph API response', async () => {
				const mockFunctions = createMockLoadOptionsFunctions();
				mockFunctions._httpRequestWithAuthentication.mockResolvedValue({
					value: [
						{ id: 'site-1', displayName: 'Site One', webUrl: 'https://example.com/site1' },
						{ id: 'site-2', displayName: 'Site Two', webUrl: 'https://example.com/site2' },
					],
				});

				const result = await searchSites.call(mockFunctions);

				expect(result.results).toHaveLength(2);
				expect(result.results[0]).toEqual({
					name: 'Site One',
					value: 'site-1',
					url: 'https://example.com/site1',
				});
			});

			it('uses wildcard search when no filter provided', async () => {
				const mockFunctions = createMockLoadOptionsFunctions();
				mockFunctions._httpRequestWithAuthentication.mockResolvedValue({ value: [] });

				await searchSites.call(mockFunctions);

				expect(mockFunctions._httpRequestWithAuthentication).toHaveBeenCalledWith(
					mockFunctions,
					'microsoftGraphOAuth2Api',
					expect.objectContaining({
						url: expect.stringContaining('search=*'),
					}),
				);
			});

			it('uses filter term when provided', async () => {
				const mockFunctions = createMockLoadOptionsFunctions();
				mockFunctions._httpRequestWithAuthentication.mockResolvedValue({ value: [] });

				await searchSites.call(mockFunctions, 'marketing');

				expect(mockFunctions._httpRequestWithAuthentication).toHaveBeenCalledWith(
					mockFunctions,
					'microsoftGraphOAuth2Api',
					expect.objectContaining({
						url: expect.stringContaining('search=marketing'),
					}),
				);
			});
		});

		describe('edge cases', () => {
			it('returns empty results when API returns no value', async () => {
				const mockFunctions = createMockLoadOptionsFunctions();
				mockFunctions._httpRequestWithAuthentication.mockResolvedValue({});

				const result = await searchSites.call(mockFunctions);

				expect(result.results).toHaveLength(0);
			});
		});

		describe('error handling', () => {
			it('logs error and returns empty results on API failure', async () => {
				const mockFunctions = createMockLoadOptionsFunctions();
				mockFunctions._httpRequestWithAuthentication.mockRejectedValue(new Error('Network error'));

				const result = await searchSites.call(mockFunctions);

				expect(result.results).toHaveLength(0);
				expect(mockFunctions.logger.error).toHaveBeenCalledWith('Failed to search sites', {
					error: 'Network error',
				});
			});
		});
	});

	describe('getDrives', () => {
		describe('basic functionality', () => {
			it('returns drives for selected site', async () => {
				const mockFunctions = createMockLoadOptionsFunctions({
					siteId: 'site-123',
				});
				mockFunctions._httpRequestWithAuthentication.mockResolvedValue({
					value: [
						{ id: 'drive-1', name: 'Documents' },
						{ id: 'drive-2', name: 'Shared Files' },
					],
				});

				const result = await getDrives.call(mockFunctions);

				expect(result.results).toHaveLength(2);
				expect(result.results[0]).toEqual({
					name: 'Documents',
					value: 'drive-1',
				});
			});

			it('handles resourceLocator format for parameters', async () => {
				const mockFunctions = createMockLoadOptionsFunctions({
					siteId: { mode: 'list', value: 'site-456' },
				});
				mockFunctions._httpRequestWithAuthentication.mockResolvedValue({ value: [] });

				await getDrives.call(mockFunctions);

				expect(mockFunctions._httpRequestWithAuthentication).toHaveBeenCalledWith(
					mockFunctions,
					'microsoftGraphOAuth2Api',
					expect.objectContaining({
						url: expect.stringContaining('/sites/site-456/drives'),
					}),
				);
			});
		});

		describe('edge cases', () => {
			it('returns empty results when siteId is missing', async () => {
				const mockFunctions = createMockLoadOptionsFunctions({});

				const result = await getDrives.call(mockFunctions);

				expect(result.results).toHaveLength(0);
				expect(mockFunctions._httpRequestWithAuthentication).not.toHaveBeenCalled();
			});
		});

		describe('error handling', () => {
			it('logs error and returns empty results on API failure', async () => {
				const mockFunctions = createMockLoadOptionsFunctions({
					siteId: 'site-123',
				});
				mockFunctions._httpRequestWithAuthentication.mockRejectedValue(new Error('Network error'));

				const result = await getDrives.call(mockFunctions);

				expect(result.results).toHaveLength(0);
				expect(mockFunctions.logger.error).toHaveBeenCalledWith('Failed to get drives', {
					error: 'Network error',
				});
			});
		});
	});

	describe('getFiles', () => {
		describe('basic functionality', () => {
			it('returns only Excel files from drive', async () => {
				const mockFunctions = createMockLoadOptionsFunctions({
					driveId: 'drive-123',
				});
				mockFunctions._httpRequestWithAuthentication.mockResolvedValue({
					value: [
						{ id: 'file-1', name: 'Report.xlsx', file: {} },
						{ id: 'file-2', name: 'Document.docx', file: {} },
						{ id: 'file-3', name: 'Data.xlsx', file: {} },
						{ id: 'folder-1', name: 'Subfolder' }, // No file property = folder
					],
				});

				const result = await getFiles.call(mockFunctions);

				expect(result.results).toHaveLength(2);
				expect(result.results[0]).toEqual({ name: 'Report.xlsx', value: 'file-1' });
				expect(result.results[1]).toEqual({ name: 'Data.xlsx', value: 'file-3' });
			});

			it('handles resourceLocator format for driveId', async () => {
				const mockFunctions = createMockLoadOptionsFunctions({
					driveId: { mode: 'list', value: 'drive-456' },
				});
				mockFunctions._httpRequestWithAuthentication.mockResolvedValue({ value: [] });

				await getFiles.call(mockFunctions);

				expect(mockFunctions._httpRequestWithAuthentication).toHaveBeenCalledWith(
					mockFunctions,
					'microsoftGraphOAuth2Api',
					expect.objectContaining({
						url: expect.stringContaining('/drives/drive-456/root/children'),
					}),
				);
			});
		});

		describe('edge cases', () => {
			it('returns empty results when driveId is missing', async () => {
				const mockFunctions = createMockLoadOptionsFunctions({});

				const result = await getFiles.call(mockFunctions);

				expect(result.results).toHaveLength(0);
				expect(mockFunctions._httpRequestWithAuthentication).not.toHaveBeenCalled();
			});
		});

		describe('error handling', () => {
			it('logs error and returns empty results on API failure', async () => {
				const mockFunctions = createMockLoadOptionsFunctions({
					driveId: 'drive-123',
				});
				mockFunctions._httpRequestWithAuthentication.mockRejectedValue(new Error('Access denied'));

				const result = await getFiles.call(mockFunctions);

				expect(result.results).toHaveLength(0);
				expect(mockFunctions.logger.error).toHaveBeenCalledWith('Failed to get files', {
					error: 'Access denied',
				});
			});
		});
	});

	describe('getSheets', () => {
		beforeEach(() => {
			mockWorkbookInstance.xlsx.load.mockReset();
			mockWorkbookInstance.worksheets = [];
		});

		describe('basic functionality', () => {
			it('returns worksheet names from Excel file', async () => {
				const mockFunctions = createMockLoadOptionsFunctions({
					siteId: 'site-123',
					driveId: 'drive-123',
					fileId: 'file-123',
				});
				mockFunctions._httpRequestWithAuthentication.mockResolvedValue(
					Buffer.from('mock-excel-data'),
				);
				mockWorkbookInstance.worksheets = [
					{ name: 'Sheet1' },
					{ name: 'Sheet2' },
					{ name: 'Data' },
				];

				const result = await getSheets.call(mockFunctions);

				expect(result.results).toHaveLength(3);
				expect(result.results[0]).toEqual({ name: 'Sheet1', value: 'Sheet1' });
				expect(result.results[1]).toEqual({ name: 'Sheet2', value: 'Sheet2' });
				expect(result.results[2]).toEqual({ name: 'Data', value: 'Data' });
			});

			it('downloads file content and parses with ExcelJS', async () => {
				const mockFunctions = createMockLoadOptionsFunctions({
					siteId: 'site-123',
					driveId: 'drive-123',
					fileId: 'file-123',
				});
				const fileContent = Buffer.from('mock-excel-data');
				mockFunctions._httpRequestWithAuthentication.mockResolvedValue(fileContent);
				mockWorkbookInstance.worksheets = [];

				await getSheets.call(mockFunctions);

				expect(mockFunctions._httpRequestWithAuthentication).toHaveBeenCalledWith(
					mockFunctions,
					'microsoftGraphOAuth2Api',
					expect.objectContaining({
						url: expect.stringContaining('/sites/site-123/drives/drive-123/items/file-123/content'),
						encoding: 'arraybuffer',
					}),
				);
				expect(mockWorkbookInstance.xlsx.load).toHaveBeenCalledWith(fileContent);
			});
		});

		describe('edge cases', () => {
			it('returns empty results when required params are missing', async () => {
				const mockFunctions = createMockLoadOptionsFunctions({
					siteId: 'site-123',
					fileId: 'file-123',
					// driveId missing
				});

				const result = await getSheets.call(mockFunctions);

				expect(result.results).toHaveLength(0);
				expect(mockFunctions._httpRequestWithAuthentication).not.toHaveBeenCalled();
			});
		});

		describe('error handling', () => {
			it('logs error and returns empty results on Excel parse failure', async () => {
				const mockFunctions = createMockLoadOptionsFunctions({
					siteId: 'site-123',
					driveId: 'drive-123',
					fileId: 'file-123',
				});
				mockFunctions._httpRequestWithAuthentication.mockResolvedValue(Buffer.from('invalid-data'));
				mockWorkbookInstance.xlsx.load.mockRejectedValue(new Error('Invalid Excel format'));

				const result = await getSheets.call(mockFunctions);

				expect(result.results).toHaveLength(0);
				expect(mockFunctions.logger.error).toHaveBeenCalledWith('Failed to get sheets', {
					error: 'Invalid Excel format',
				});
			});
		});
	});

	describe('getTables', () => {
		describe('basic functionality', () => {
			it('returns tables from Graph API', async () => {
				const mockFunctions = createMockLoadOptionsFunctions({
					siteId: 'site-123',
					driveId: 'drive-123',
					fileId: 'file-123',
				});
				mockFunctions._httpRequestWithAuthentication.mockResolvedValue({
					value: [
						{ id: 'table-1', name: 'SalesData' },
						{ id: 'table-2', name: 'Inventory' },
					],
				});

				const result = await getTables.call(mockFunctions);

				expect(result.results).toHaveLength(2);
				expect(result.results[0]).toEqual({ name: 'SalesData', value: 'SalesData' });
				expect(result.results[1]).toEqual({ name: 'Inventory', value: 'Inventory' });
			});

			it('uses workbook/tables Graph API endpoint', async () => {
				const mockFunctions = createMockLoadOptionsFunctions({
					siteId: 'site-123',
					driveId: 'drive-123',
					fileId: 'file-123',
				});
				mockFunctions._httpRequestWithAuthentication.mockResolvedValue({ value: [] });

				await getTables.call(mockFunctions);

				expect(mockFunctions._httpRequestWithAuthentication).toHaveBeenCalledWith(
					mockFunctions,
					'microsoftGraphOAuth2Api',
					expect.objectContaining({
						url: expect.stringContaining('/workbook/tables'),
					}),
				);
			});
		});

		describe('edge cases', () => {
			it('returns empty results when required params are missing', async () => {
				const mockFunctions = createMockLoadOptionsFunctions({
					siteId: 'site-123',
					driveId: 'drive-123',
					// fileId missing
				});

				const result = await getTables.call(mockFunctions);

				expect(result.results).toHaveLength(0);
				expect(mockFunctions._httpRequestWithAuthentication).not.toHaveBeenCalled();
			});
		});

		describe('error handling', () => {
			it('logs error and returns empty results on API failure', async () => {
				const mockFunctions = createMockLoadOptionsFunctions({
					siteId: 'site-123',
					driveId: 'drive-123',
					fileId: 'file-123',
				});
				mockFunctions._httpRequestWithAuthentication.mockRejectedValue(
					new Error('Table access denied'),
				);

				const result = await getTables.call(mockFunctions);

				expect(result.results).toHaveLength(0);
				expect(mockFunctions.logger.error).toHaveBeenCalledWith('Failed to get tables', {
					error: 'Table access denied',
				});
			});
		});
	});
});
