# Contributing to xlsxzero

Thank you for your interest in contributing to xlsxzero! This document provides guidelines and instructions for contributing to the project.

## Code of Conduct

This project adheres to the Rust Code of Conduct. Please be respectful and constructive in all interactions.

## Getting Started

### Prerequisites

- Rust 1.70 or later
- `cargo` (comes with Rust)
- Git

### Setting Up the Development Environment

1. Fork the repository on GitHub
2. Clone your fork:
   ```bash
   git clone https://github.com/your-username/xlsxzero.git
   cd xlsxzero
   ```
3. Build the project:
   ```bash
   cargo build
   ```
4. Run tests:
   ```bash
   cargo test
   ```

## Development Workflow

### Finding Work

Check the [issue list](issues/issue_list.md) for available tasks. Issues are organized by:
- **Phase**: Phase I (Core), Phase II (Data Integrity), Phase III (Robustness)
- **Priority**: P0 (Critical), P1 (High), P2 (Medium), P3 (Low)
- **Category**: Core, API, Documentation, Testing, etc.

### Making Changes

1. Create a new branch from `master`:
   ```bash
   git checkout -b issue-XX-description
   ```
   Use the issue number in the branch name (e.g., `issue-24-documentation`).

2. Make your changes following the coding standards below.

3. Write or update tests for your changes.

4. Ensure all tests pass:
   ```bash
   cargo test
   ```

5. Run the linter:
   ```bash
   cargo clippy
   ```

6. Format your code:
   ```bash
   cargo fmt
   ```

7. Update documentation if needed:
   ```bash
   cargo doc --no-deps --open
   ```

### Coding Standards

- **Rust Style**: Follow the standard Rust style guide. Use `cargo fmt` to format code.
- **Documentation**: All public APIs must have documentation comments using `///`.
- **Error Handling**: Use the `XlsxToMdError` type for all errors.
- **Testing**: Write unit tests for new functionality. Use property-based testing with `proptest` where appropriate.
- **Naming**: Follow Rust naming conventions:
  - Types: `UpperCamelCase`
  - Functions and variables: `snake_case`
  - Constants: `SCREAMING_SNAKE_CASE`

### Commit Messages

Write clear, descriptive commit messages:

```
Short summary (50 chars or less)

More detailed explanation if needed. Wrap at 72 characters.
Explain what and why, not how.

- Bullet points are fine
- Use present tense ("Add feature" not "Added feature")
- Reference issue numbers: Fixes #24
```

### Submitting Changes

1. Push your branch to your fork:
   ```bash
   git push origin issue-XX-description
   ```

2. Create a Pull Request on GitHub:
   - Provide a clear description of your changes
   - Reference the related issue number
   - Include any relevant test results or screenshots

3. Respond to review feedback promptly.

## Testing

### Running Tests

```bash
# Run all tests
cargo test

# Run tests for a specific module
cargo test --lib builder

# Run tests with output
cargo test -- --nocapture
```

### Writing Tests

- Unit tests should be in the same file as the code they test (in `#[cfg(test)]` modules)
- Integration tests go in `tests/` directory
- Use descriptive test names that explain what is being tested
- Follow the pattern: `test_<functionality>_<scenario>`

### Property-Based Testing

For complex logic, consider using property-based testing with `proptest`:

```rust
use proptest::prelude::*;

proptest! {
    #[test]
    fn test_property(input in any::<InputType>()) {
        // Test property holds for all inputs
    }
}
```

## Documentation

### API Documentation

All public APIs must have documentation comments:

```rust
/// Brief description.
///
/// More detailed explanation.
///
/// # Examples
///
/// ```rust,no_run
/// // Example code
/// ```
pub fn public_function() -> Result<(), Error> {
    // ...
}
```

### Building Documentation

```bash
# Build documentation
cargo doc --no-deps

# Open documentation in browser
cargo doc --no-deps --open
```

### Documentation Standards

- Use `///` for public API documentation
- Use `//!` for module-level documentation
- Include examples in code blocks with `rust,no_run` or `rust`
- Document all parameters, return values, and possible errors
- Use markdown formatting for better readability

## Project Structure

```
xlsxzero/
├── src/              # Source code
│   ├── api.rs        # Public API types
│   ├── builder.rs    # Builder pattern implementation
│   ├── error.rs      # Error types
│   ├── formatter.rs  # Cell formatting
│   ├── grid.rs       # Grid construction
│   ├── parser.rs     # Excel parsing
│   └── lib.rs        # Library entry point
├── tests/            # Integration tests
├── examples/         # Example programs
├── benches/          # Benchmark tests
├── docs/             # Design documentation
└── issues/           # Issue tracking
```

## Issue Management

### Creating Issues

When creating a new issue:
- Use a clear, descriptive title
- Provide detailed description
- Include steps to reproduce (for bugs)
- Reference related documentation
- Assign appropriate priority and phase

### Working on Issues

1. Comment on the issue to claim it
2. Create a branch named `issue-XX-description`
3. Make your changes
4. Update the issue file with implementation results
5. Update `issue_list.md` when the issue is complete

## Release Process

Releases follow semantic versioning:
- **MAJOR**: Breaking API changes
- **MINOR**: New features (backward compatible)
- **PATCH**: Bug fixes

Before release:
1. Update `CHANGELOG.md`
2. Update version in `Cargo.toml`
3. Run full test suite
4. Update documentation
5. Create release tag

## Questions?

If you have questions or need help:
- Check the [documentation](docs/)
- Review existing issues
- Open a new issue with the "question" label

Thank you for contributing to xlsxzero!

