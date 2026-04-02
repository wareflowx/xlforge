# xlforge

A command-line interface for Excel manipulation. Designed for agents and scripts.

## Overview

xlforge is a CLI tool that provides comprehensive Excel operations through a set of stateless commands. Every command includes the source file as a parameter, making it perfect for automation, scripts, and AI integrations.

```
xlforge cell get report.xlsx "Data!A1"
xlforge cell set report.xlsx "Data!A1" "Hello World"
xlforge sql query "SELECT * FROM 'report.xlsx!Data'" --to output.csv
```

## Features

- **Cell operations** - get, set, formula, clear, copy, bulk operations
- **Sheet management** - create, delete, rename, copy, use
- **Formatting** - cell styles, number formats, borders, colors
- **Data import/export** - CSV, Excel tables, database sync
- **Charts and validations** - create and manage Excel objects
- **SQL Bridge** - Query Excel, CSV, and databases using DuckDB
- **Semantic search** - AI-powered cell finding by meaning
- **Macro recorder** - Transform user actions into reusable scripts
- **Checkpoint versioning** - Git-like versioning for Excel files
- **Batch execution** - Run scripts in a single COM session

## Architecture

xlforge follows a layered architecture: **API -> SDK -> CLI**

See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for details.

## Project Structure

```
xlforge/
├── docs/                    # Documentation
│   ├── ARCHITECTURE.md      # Layered architecture design
│   ├── testing/README.md    # Testing strategy
│   ├── README.md            # Overview and quick reference
│   ├── context.md           # Context management
│   ├── batch.md             # Batch execution
│   ├── examples.md          # Complete workflow examples
│   ├── reference.md         # Error codes, global flags
│   └── commands/            # Command documentation
├── xlforge/                 # Source code
│   ├── __init__.py          # CLI entry point
│   └── ...
├── tests/                   # Test suite
│   ├── conftest.py          # Pytest fixtures
│   └── ...
└── pyproject.toml           # Project configuration
```

## Quick Start

### Installation

```bash
pip install xlforge
```

### Basic Usage

```bash
# Read a cell
xlforge cell get report.xlsx "A1"

# Write to a cell
xlforge cell set report.xlsx "A1" "Hello World"

# Set a formula
xlforge cell formula report.xlsx "B1" "=SUM(A:A)"

# List sheets
xlforge sheet list report.xlsx

# Save with output to new file
xlforge file save report.xlsx --output "backup.xlsx"
```

### Global Flags

All commands support these flags:

| Flag | Description |
|------|-------------|
| `--json` | JSON output |
| `--json-errors` | Return errors as JSON |
| `--dry-run` | Preview without executing |
| `--engine <name>` | Force engine (xlwings\|openpyxl) |
| `--verbose` | Verbose logging |

## Development

### Setup

```bash
# Install dependencies
uv sync

# Run tests
uv run pytest tests/ -v

# Run linter
uv run ruff check .

# Format code
uv run ruff format .
```

### Running the CLI

```bash
# Via uv
uv run xlforge --help
uv run xlforge ping

# Via installed package
xlforge --help
```

### Architecture Layers

| Layer | Location | Purpose |
|-------|----------|---------|
| **CLI** | `xlforge/` | Typer commands, argument parsing, output formatting |
| **SDK** | `xlforge/core.py` | Business logic, validation, orchestration |
| **API** | `xlforge/engines/` | Excel interaction (xlwings, openpyxl, duckdb) |

See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for full details.

## Testing

xlforge uses a multi-layer testing approach:

- **Unit tests** - Pure functions, SDK logic (mocked engines)
- **Integration tests** - CLI commands with mocked SDK
- **E2E tests** - Real Excel/xlwings (requires Windows + Excel)

See [docs/testing/README.md](docs/testing/README.md) for details.

## Error Codes

xlforge returns specific exit codes for different errors:

| Code | Meaning |
|------|---------|
| 0 | Success |
| 1 | General error |
| 2 | File not found |
| 3 | Sheet not found |
| 4 | Cell not found |
| 6 | File locked |
| 9 | Feature unavailable (headless) |

See [docs/reference.md](docs/reference.md) for all 127 error codes.

## Commands

| Category | Commands |
|----------|----------|
| **File** | open, save, close, info, kill, recover, check, monitor |
| **Sheet** | list, create, delete, rename, copy, use |
| **Cell** | get, set, formula, clear, copy, bulk, search, fill |
| **Format** | cell, range, apply |
| **Data** | import csv, export csv |
| **Table** | create, link, sync-schema, refresh |
| **Chart** | create |
| **Validation** | create |
| **Protection** | freeze, protect, unprotect |
| **App** | visible, calculate, focus, alert, wait-idle |
| **Checkpoint** | create, list, restore, delete |
| **SQL** | query, push, pull, connect |
| **Semantic** | index, query, describe |
| **Watch** | start, stop |

See [docs/commands/](docs/commands/) for detailed command documentation.

## License

MIT License
