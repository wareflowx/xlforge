<p align="center">
  <img src="public/banner.jpg" alt="xlforge Logo" width="100%">
</p>

<h1 align="center">xlforge</h1>

<p align="center">
  <a href="https://github.com/wareflowx/wareflow/stargazers">
    <img src="https://img.shields.io/github/stars/wareflowx/wareflow" alt="Stars">
  </a>
  <a href="https://github.com/wareflowx/wareflow/fork">
    <img src="https://img.shields.io/github/forks/wareflowx/wareflow" alt="Forks">
  </a>
  <a href="https://github.com/wareflowx/wareflow/blob/main/LICENSE">
    <img src="https://img.shields.io/github/license/wareflowx/wareflow" alt="License">
  </a>
</p>

> A command-line interface for Excel manipulation. Designed for agents and scripts.

## About

xlforge is a powerful CLI for Excel manipulation built with an agent-first philosophy. Every command is stateless and includes the source file as a parameter, making it perfect for automation, scripts, and AI integrations.

Similar to how a robust build tool provides declarative commands for complex tasks, xlforge provides a comprehensive set of commands for Excel operations - from basic cell manipulation to advanced features like SQL queries, semantic search, and macro recording.

xlforge uses a **Hybrid Engine** architecture that automatically selects the best approach:
- **xlwings** when Excel is installed - full control including cells, formatting, charts, and macros
- **openpyxl** for headless/Docker environments - read/write cells, sheet operations, and basic formatting

## Philosophy

- **Agent-first** - Every command is a standalone shell operation
- **Stateless** - No interactive mode, no shell persistence
- **JSON everywhere** - All commands support JSON output; errors can be returned as JSON
- **Fail fast with retry** - Exponential backoff on file lock (max 3 attempts)
- **Transaction-safe** - Batch operations can be wrapped in transactions

## Features

- Cell operations - get, set, formula, clear, copy, bulk operations
- Sheet management - create, delete, rename, copy, use
- Formatting - cell styles, number formats, borders, colors
- Data import/export - CSV, Excel tables, database sync
- Charts and validations - create and manage Excel objects
- SQL Bridge - Query Excel, CSV, and databases using DuckDB
- Semantic search - AI-powered cell finding by meaning
- Macro recorder - Transform user actions into reusable scripts
- Checkpoint versioning - Git-like versioning for Excel files
- Batch execution - Run scripts in a single COM session

## Quick Start

```bash
# Install xlforge
pip install xlforge

# Read a cell
xlforge cell get report.xlsx "A1"

# Write to a cell
xlforge cell set report.xlsx "A1" "Hello World"

# Set a formula
xlforge cell formula report.xlsx "B1" "=SUM(A:A)"

# Set context to avoid repeating filename
xlforge use report.xlsx --sheet Data
cell set A1 "Value"  # Operates on Data!A1 in report.xlsx

# Run a batch script
xlforge run script.xlf --transaction
```

## Documentation

- [Getting Started](docs/README.md) - Overview and architecture
- [Context Management](docs/context.md) - Environment variables and default context
- [Batch Execution](docs/batch.md) - Run scripts with transaction support
- [Examples](docs/examples.md) - Complete workflow examples

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

MIT License - see the [LICENSE](LICENSE) file for details.
