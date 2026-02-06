# ExcelJS Table Corruption Limitation

**Date:** 2026-02-06

## Summary

Write operations on files containing Excel Tables can corrupt those tables due to a bug in the exceljs library.

## What is an Excel Table?

An Excel Table is a formal feature (Insert → Table or Ctrl+T), not just data in cells. Tables have:

- Named references (e.g., "Table1")
- Filter dropdown buttons on headers
- Banded row colors
- Auto-expansion when adding data
- Structured formula references like `=SUM(Table1[Sales])`

Most users have regular data in cells, not formal Tables.

## The Problem

When exceljs opens and saves a file that contains Excel Tables, it corrupts the table XML structure. Microsoft Excel then shows:

> "We found a problem with some content. Do you want us to try to recover as much as we can?"

Recovery removes the Table features (filters, structured references) from the file.

## Root Cause

exceljs writes malformed `autoFilter` XML elements that don't match the OpenXML standard format expected by Excel.

Reference: https://github.com/exceljs/exceljs/issues/2585

## Affected Operations

**Safe (read-only, no file save):**

- getSheets
- readRows
- getWorkbooks
- getColumns (table)
- getTableRows (table)
- lookup (table)

**Risk of corruption (saves file back):**

- appendRows
- updateCell
- upsertRows
- clearSheet
- deleteSheet
- addSheet

## Why This Node Uses exceljs Anyway

The native n8n Microsoft Excel node uses Graph API Excel endpoints with WAC (Web Application Companion) tokens. These endpoints handle Tables correctly server-side but have token/authentication issues that fail for some SharePoint configurations.

This node downloads the entire file, modifies it with exceljs, and re-uploads. This bypasses WAC token issues but introduces the table corruption risk.

## Recommendations

1. **Document in README** - Warn users about this limitation
2. **Target use case** - Position node for files without formal Excel Tables
3. **Consider detection** - Could check for tables before write and warn user

## Potential Future Solutions

- Fork `@nbelyh/exceljs` reportedly preserves tables better (but has issues with column name changes)
- Monitor exceljs issue #2585 for upstream fix
- Hybrid approach: use Graph API Excel endpoints for table operations, exceljs for sheet operations
