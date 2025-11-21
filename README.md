# xlsxzero

[![License: MIT OR Apache-2.0](https://img.shields.io/badge/license-MIT%2FApache--2.0-blue.svg)](LICENSE-MIT)
[![Rust](https://github.com/your-org/xlsxzero/actions/workflows/test.yml/badge.svg)](https://github.com/your-org/xlsxzero/actions)

Pure-Rust Excel parser and Markdown converter for RAG systems.

## Overview

xlsxzero is a high-performance, memory-efficient Rust crate designed to parse Excel files (XLSX format) and convert them into structured Markdown format. It is optimized for RAG (Retrieval-Augmented Generation) systems that require efficient processing of large Excel files.

## Features

- **Pure Rust Implementation**: No dependencies on C/C++ libraries
- **Streaming Architecture**: Process large Excel files with minimal memory footprint
- **Structured Markdown Output**: Convert Excel tables to GitHub Flavored Markdown
- **Cell Merging Support**: Handle merged cells with multiple strategies
- **Date/Time Conversion**: Accurate conversion of Excel serial dates
- **Formula Support**: Extract cached values or formula strings

## Installation

Add this to your `Cargo.toml`:

```toml
[dependencies]
xlsxzero = "0.1.0"
```

## Quick Start

### Basic Usage

```rust
use std::fs::File;
use xlsxzero::ConverterBuilder;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let input = File::open("example.xlsx")?;
    let output = File::create("output.md")?;
    
    ConverterBuilder::new()
        .build()?
        .convert(input, output)?;
    
    Ok(())
}
```

### Custom Configuration

```rust
use std::fs::File;
use xlsxzero::{ConverterBuilder, SheetSelector, MergeStrategy, DateFormat};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let converter = ConverterBuilder::new()
        .with_sheet_selector(SheetSelector::Index(0))  // First sheet only
        .with_merge_strategy(MergeStrategy::HtmlFallback)  // HTML for merged cells
        .with_date_format(DateFormat::Custom("%Y年%m月%d日".to_string()))  // Japanese format
        .build()?;
    
    let input = File::open("example.xlsx")?;
    let output = File::create("output.md")?;
    converter.convert(input, output)?;
    
    Ok(())
}
```

### Convert to String

```rust
use std::fs::File;
use xlsxzero::ConverterBuilder;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let converter = ConverterBuilder::new().build()?;
    let input = File::open("example.xlsx")?;
    let markdown = converter.convert_to_string(input)?;
    println!("{}", markdown);
    Ok(())
}
```

## Examples

The repository includes several example programs demonstrating different use cases:

- **Basic Conversion** (`examples/basic_conversion.rs`): Simple file-to-file conversion
- **Custom Configuration** (`examples/custom_config.rs`): Using advanced configuration options
- **CLI Tool** (`examples/cli_tool.rs`): Building a command-line tool

Run an example:

```bash
cargo run --example basic_conversion -- input.xlsx output.md
cargo run --example custom_config -- input.xlsx output.md
cargo run --example cli_tool -- input.xlsx output.md --sheet-index 0
```

## Status

This project is currently in early development. Phase I features are being implemented.

## Documentation

- [Requirements](docs/requirements.md)
- [Architecture](docs/architecture.md)
- [Interface Design](docs/interface.md)
- [Detailed Design](docs/detailed_design.md)
- [Test Specification](docs/test_specification.md)

## License

Licensed under either of:

- Apache License, Version 2.0 ([LICENSE-APACHE](LICENSE-APACHE) or <http://www.apache.org/licenses/LICENSE-2.0>)
- MIT license ([LICENSE-MIT](LICENSE-MIT) or <http://opensource.org/licenses/MIT>)

at your option.

## Contributing

Contributions are welcome! Please see [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines and the [issue list](issues/issue_list.md) for development tasks.

## API Documentation

Full API documentation is available at [docs.rs/xlsxzero](https://docs.rs/xlsxzero) (when published) or by running:

```bash
cargo doc --open
```
