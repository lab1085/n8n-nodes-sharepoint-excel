# Changelog

<!-- markdownlint-disable MD024 MD013 -->

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.2.0] - 2026-02-04

### Added

- Data mode pattern for appendRows and upsertRows operations ([ff7bc20](../../commit/ff7bc20))
- Read options (headerRow, startRow, maxRows) to readRows and getTableRows ([85a0865](../../commit/85a0865))
- Comprehensive test suite with vitest ([876c357](../../commit/876c357))

### Changed

- Simplified OAuth2 credentials by extending microsoftOAuth2Api ([5de59f1](../../commit/5de59f1))

### Fixed

- Various bug fixes and improved error handling ([876c357](../../commit/876c357))

## [0.1.1] - 2026-02-03

Initial release of the Microsoft SharePoint Excel node for n8n. This node provides read and write operations for Excel files stored in SharePoint document libraries, bypassing WAC token limitations by using a download-modify-upload pattern with the exceljs library.

### Added

#### Core Node

- Initial SharePoint Excel node implementation with Microsoft Graph API integration ([10dc09f](../../commit/10dc09f))
- Dynamic dropdowns for Site, Drive, File, and Sheet selection with searchable site picker ([4f000b9](../../commit/4f000b9))
- Resource selector organizing operations into Sheet, Table, and Workbook categories ([f2460c1](../../commit/f2460c1))
- Microsoft Graph OAuth2 credentials for authentication ([10dc09f](../../commit/10dc09f))

#### Sheet Operations

- **Get Sheets** - List all worksheets in a workbook
- **Get Rows** - Read rows with configurable header row, start row, and max rows
- **Append Rows** - Add new rows to the end of a sheet
- **Update** - Update a specific cell by reference (e.g., A1, B5)
- **Upsert Rows** - Insert or update rows based on a key column
- **Clear** - Clear all data from a sheet
- **Delete** - Delete a sheet from the workbook

#### Table Operations

- **Get Rows** - Retrieve all rows from an Excel table
- **Get Columns** - Get column definitions from a table
- **Lookup** - Find a row by column value

#### Workbook Operations

- **Get Sheets** - List all sheets in the workbook
- **Add Sheet** - Create a new sheet in the workbook
- **Delete** - Delete the workbook file
- **Get Workbooks** - List all Excel files in the drive

### Fixed

- Improved error handling and binary upload support for large files ([b7ae6f4](../../commit/b7ae6f4))
- Use wildcard search to list all SharePoint sites in dropdown ([4be5fc3](../../commit/4be5fc3))

### Changed

- Modularized node structure into organized actions, types, and descriptions ([e5ead3f](../../commit/e5ead3f))
- Removed OneDrive support to focus on SharePoint (OneDrive works with native n8n Excel node) ([045f059](../../commit/045f059))
- Removed source selector dropdown (now SharePoint-only) ([d0abf73](../../commit/d0abf73))
- Renamed node to "Microsoft SharePoint Excel" for better discoverability ([94a9b1a](../../commit/94a9b1a))
- Replaced Biome with Prettier for code formatting ([faf9745](../../commit/faf9745))

[unreleased]: ../../compare/v0.1.1...HEAD
[0.1.1]: ../../releases/tag/v0.1.1
