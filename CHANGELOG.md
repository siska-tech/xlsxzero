# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Initial implementation of Excel to Markdown converter
- Support for XLSX file format parsing
- Streaming architecture for memory-efficient processing
- Cell merging support with multiple strategies (DataDuplication, HtmlFallback)
- Date/time conversion with ISO 8601 and custom formats
- Formula support (cached values and formula strings)
- Sheet selection (by index, name, or all sheets)
- Range limiting for partial sheet conversion
- Hidden element filtering (sheets only in Phase I)
- Comprehensive error handling with structured error types
- Builder pattern API for flexible configuration

### Documentation
- Complete API documentation with examples
- Architecture documentation
- Interface design documentation
- Test specification
- README with usage examples
- Contributing guidelines

## [0.1.0] - 2025-01-27

### Added
- Initial release
- Core conversion functionality
- Basic cell formatting
- Markdown table output
- HTML table output for merged cells

[Unreleased]: https://github.com/your-org/xlsxzero/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/your-org/xlsxzero/releases/tag/v0.1.0

