# Testing Standards

This document defines testing conventions for the n8n-nodes-sharepoint-excel project.

## Test File Location

Tests are co-located with source files:

```
nodes/SharePointExcel/actions/sheet/
├── clearSheet.ts
├── clearSheet.test.ts
├── appendRows.ts
└── appendRows.test.ts
```

## Test File Structure

```typescript
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { execute } from './operationName';
import {} from /* shared mocks */ '../../test-utils/mocks';

// 1. Module mocks
vi.mock('../../api', () => ({
	loadWorkbook: vi.fn(),
	saveWorkbook: vi.fn(),
	getWorksheet: vi.fn(),
}));

import { loadWorkbook, saveWorkbook, getWorksheet } from '../../api';

// 2. Local helpers (if needed for operation-specific setup)

// 3. Test suite
describe('operationName', () => {
	beforeEach(() => {
		vi.clearAllMocks();
	});

	describe('basic functionality', () => {
		/* ... */
	});
	describe('operation-specific behavior', () => {
		/* ... */
	});
	describe('edge cases', () => {
		/* ... */
	});
	describe('error handling', () => {
		/* ... */
	});
});
```

## Describe Block Organization

Group tests by feature/behavior, not by implementation:

```typescript
describe('clearSheet', () => {
  describe('basic functionality', () => {
    it('clears all rows and saves the workbook', ...);
    it('handles resourceLocator format for sheetName', ...);
    it('uses correct basePath from context', ...);
  });

  describe('row clearing behavior', () => {
    it('handles empty sheet (0 rows)', ...);
    it('handles sheet with single row', ...);
    it('handles large sheets', ...);
  });

  describe('edge cases', () => {
    it('ignores input items (operation is sheet-level)', ...);
  });

  describe('error handling', () => {
    it('throws error when sheet not found', ...);
    it('throws error when loadWorkbook fails', ...);
    it('throws error when saveWorkbook fails', ...);
  });
});
```

## Mock Patterns

### Shared Mocks (test-utils/mocks.ts)

Generic, reusable mock creators:

- `createMockWorksheet(options)` - Creates worksheet mock
- `createMockWorkbook(sheets)` - Creates workbook mock
- `createMockExecuteFunctions(params)` - Creates n8n IExecuteFunctions mock
- `createMockContext(overrides)` - Creates OperationContext mock
- `setupSheetMocks(options)` - Combines above for sheet operations

### When to Use Shared vs Local

| Location              | Use Case                              |
| --------------------- | ------------------------------------- |
| `test-utils/mocks.ts` | Reusable across multiple test files   |
| Local in test file    | One-off setup specific to single test |

### vi.mocked() Setup

Standard pattern for sheet operations:

```typescript
vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);
vi.mocked(getWorksheet).mockReturnValue(worksheet as never);
vi.mocked(saveWorkbook).mockResolvedValue(undefined);
```

## Test Coverage Requirements

### Must Test

1. **Happy path** - Basic functionality works
2. **Parameter formats** - String and ResourceLocator formats
3. **Context usage** - basePath passed correctly to API functions
4. **Edge cases** - Empty data, single item, large data sets
5. **Return structure** - Correct shape and values
6. **Error propagation** - API and helper errors bubble up correctly

### Verify Function Calls

```typescript
// Verify called
expect(loadWorkbook).toHaveBeenCalledTimes(1);

// Verify called with correct arguments
expect(spliceRows).toHaveBeenCalledWith(1, rowCount);

// Verify NOT called when appropriate
expect(spliceRows).not.toHaveBeenCalled();
```

### Verify Outcomes

```typescript
// Verify return structure
expect(result[0].json).toEqual({
	success: true,
	sheet: 'Sheet1',
	clearedRows: 3,
	clearedColumns: 2,
});

// Or verify specific fields
expect(result[0].json.clearedRows).toBe(3);
```

## Naming Conventions

### Test Names

Use descriptive names that explain behavior:

```typescript
// Good
it('handles empty sheet (0 rows)', ...);
it('writes data to correct columns (A, B) not shifted (B, C)', ...);
it('throws error when columns value is empty', ...);

// Avoid
it('test empty', ...);
it('should work', ...);
```

### Describe Block Names

Use feature/behavior names, not implementation:

```typescript
// Good
describe('row clearing behavior', () => {});
describe('column mapping with existing headers', () => {});

// Avoid
describe('spliceRows tests', () => {});
describe('implementation details', () => {});
```

## Error Testing

### Validation Errors

When testing validation errors in the action itself:

```typescript
it('throws error when columns value is empty', async () => {
	// setup...

	await expect(execute.call(mockFunctions, [{ json: {} }], context)).rejects.toThrow(
		'No column values provided in manual mapping mode',
	);
});
```

### Error Propagation from API Layer

Test that errors from dependencies (API calls, helpers) propagate correctly. Import `NodeOperationError` to create realistic error mocks:

```typescript
import { NodeOperationError } from 'n8n-workflow';

it('throws error when sheet not found', async () => {
	const workbook = createMockWorkbook({});
	const mockFunctions = createMockExecuteFunctions({ sheetName: 'NonExistent' });
	const context = createMockContext({ operation: 'clearSheet' });

	vi.mocked(loadWorkbook).mockResolvedValue(workbook as never);
	vi.mocked(getWorksheet).mockImplementation(() => {
		throw new NodeOperationError(
			mockFunctions.getNode(),
			'Sheet "NonExistent" not found in workbook',
		);
	});

	await expect(execute.call(mockFunctions, [{ json: {} }], context)).rejects.toThrow(
		'Sheet "NonExistent" not found in workbook',
	);
});

it('throws error when loadWorkbook fails', async () => {
	const mockFunctions = createMockExecuteFunctions({ sheetName: 'Sheet1' });
	const context = createMockContext({ operation: 'clearSheet' });

	vi.mocked(loadWorkbook).mockRejectedValue(
		new NodeOperationError(mockFunctions.getNode(), 'Graph API request failed: Access denied'),
	);

	await expect(execute.call(mockFunctions, [{ json: {} }], context)).rejects.toThrow(
		'Graph API request failed: Access denied',
	);
});
```

## Anti-Patterns

### Avoid Testing Implementation Details

```typescript
// Bad - tests how, not what
expect(spliceRowsCalls).toHaveLength(4);
expect(spliceRowsCalls[0]).toEqual({ start: 4, count: 1 });
expect(spliceRowsCalls[1]).toEqual({ start: 3, count: 1 });

// Good - tests outcome
expect(spliceRows).toHaveBeenCalledWith(1, 4);
expect(result[0].json.clearedRows).toBe(4);
```

### Avoid Redundant Tests

Don't test the same thing multiple times in different describe blocks.

Tests with different input values that exercise the same code path are redundant. For example, testing `rowCount: 3`, `rowCount: 4`, and `rowCount: 10000` all verify the same logic. Instead, keep:

- **Boundary cases** (0, 1) - test edge behavior
- **One representative case** - covered by happy path test
- **Scale test** (large values) - only if testing performance or limits

### Avoid Over-Mocking

Only mock what's necessary. If a function doesn't affect the test, don't mock it.

## Running Tests

```bash
# Run all tests
bun run test

# Run specific test file
bun run test -- path/to/file.test.ts

# Watch mode
bun run test:watch

# Coverage report
bun run test:coverage
```
