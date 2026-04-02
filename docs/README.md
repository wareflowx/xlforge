# xlforge CLI

Command-line interface for Excel manipulation. Designed for agents and scripts.

## Overview

Every command is stateless and includes the source file as a parameter. No interactive mode, no shell persistence.

```bash
xlforge <command> <file.xlsx> [options]
```

---

## Architecture

xlforge uses a **Hybrid Engine** architecture:

| Engine | When Used | Capabilities |
|--------|-----------|--------------|
| **xlwings** | Excel installed | Full control: cells, formatting, charts, macros |
| **openpyxl** | Headless/Docker | Read/write cells, sheet ops, basic formatting |

The engine is auto-detected. Use `--engine <name>` to force a specific engine.

```bash
xlforge cell get report.xlsx "A1"              # Auto-detect
xlforge cell get report.xlsx "A1" --engine xlwings    # Force xlwings
xlforge cell get report.xlsx "A1" --engine openpyxl   # Force openpyxl
```

**Note:** Some commands (chart, validation) require xlwings. In headless mode, these return error code 9: `Feature requires Excel`.

---

## Design Principles

1. **Agent-first** - Every command is a standalone shell operation
2. **Auto-save with safety** - Changes save automatically (use `--dry-run` to preview)
3. **JSON everywhere** - All commands support `--json`; errors can be returned as JSON via `--json-errors`
4. **Fail fast with retry** - Exponential backoff on file lock (max 3 attempts)
5. **Context-aware** - Support for default file/sheet via environment or `use` command
6. **Transaction-safe** - Batch operations can be wrapped in transactions

---

## Quick Reference

```bash
# Essential commands
xlforge file open <file>                      # Open/create file
xlforge file save <file>                     # Save changes
xlforge file info <file>                     # Show metadata + PID

# Cell operations
xlforge cell get <file> <cell> [--json]      # Read cell
xlforge cell set <file> <cell> <value>       # Write cell
xlforge cell formula <file> <cell> <f>       # Set formula

# Context (reduces repetition)
xlforge use <file> [--sheet <name>]          # Set default context
xlforge context                               # Show current context

# SQL Bridge (DuckDB-powered)
xlforge sql query "<query>"                   # Query Excel/CSV/DB
xlforge sql push "<query>" --db <url> --to <file> <table>
xlforge sql pull <file> <range> --into <url>

# Batch (fast, single COM session)
xlforge run <script.xlf> [--dry-run] [--transaction]

# AI Context (semantic search + describe)
xlforge index create <file> --engine local --privacy-check  # Local-first indexing
xlforge query <file> "Net Profit" --coordinate              # Find by meaning
xlforge describe <file> <range> --schema-only --json       # LLM-optimized schema

# Macro Recorder (transform user actions into scripts)
xlforge record start <file> --interactive                   # Teacher mode
xlforge record stop --clean                                  # Normalize + save
```

---

## Documentation Structure

### Getting Started
- [Context Management](./context.md) - Environment variables and `use` command
- [Batch Execution](./batch.md) - Run scripts in single COM session
- [Examples](./examples.md) - Complete workflow examples

### Commands
- [File Commands](./commands/file.md) - open, save, close, info, kill
- [Sheet Commands](./commands/sheet.md) - list, create, delete, rename, copy, use
- [Cell Commands](./commands/cell.md) - get, set, formula, clear, copy, bulk
- [Format Commands](./commands/format.md) - cell, range, apply
- [Column & Row Commands](./commands/column-row.md) - width, auto-fit, height
- [Data Commands](./commands/data.md) - import csv, export csv
- [Table Commands](./commands/table.md) - create, link, sync-schema, refresh
- [Chart Commands](./commands/chart.md) - create
- [Validation Commands](./commands/validation.md) - create
- [Protection Commands](./commands/protection.md) - freeze, protect, unprotect
- [App Commands](./commands/app.md) - visible, calculate, focus, alert, wait-idle, screen-update
- [Checkpoint Commands](./commands/checkpoint.md) - Git-like versioning for Excel
- [Semantic Commands](./commands/semantic.md) - index, query, describe (AI context)
- [Watch Commands](./commands/watch.md) - Reactive triggers
- [SQL Commands](./commands/sql.md) - query, push, pull, connect

### Reference
- [Reference](./reference.md) - Error codes, global flags, retry mechanism

---

## Version

v1.1.0 - 2026-03-31 (AI-native features: semantic search, describe with sampling, macro recorder)
