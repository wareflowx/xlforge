# Context Management

## Environment Variables

```bash
export XLFORGE_FILE=report.xlsx        # Default file for all commands
export XLFORGE_SHEET=Data             # Default sheet
export XLFORGE_VISIBLE=true           # Show/hide Excel window
export XLFORGE_BACKUP=true            # Auto-backup before changes
export XLFORGE_ENGINE=xlwings        # Force engine (xlwings|openpyxl)
```

When set, you can omit the file argument:
```bash
xlforge cell get "Data!A1"           # Uses $XLFORGE_FILE
xlforge sheet list                     # Lists sheets in $XLFORGE_FILE
```

---

## use command

Sets the active context for subsequent commands:

```bash
xlforge use <file.xlsx>              # Set default file (absolute path stored)
xlforge use <file.xlsx> --sheet Data  # Set both file and sheet
xlforge context                       # Show current context
xlforge clear-use                     # Clear context
```

Context stores **absolute paths** to avoid issues when terminal changes directories.

After `xlforge use report.xlsx`, subsequent commands only need the cell/range:
```bash
xlforge use report.xlsx --sheet Data
cell set A1 "Hello"                   # Sets Data!A1 in report.xlsx
format cell B2 --bold                 # Formats Data!B2
```
