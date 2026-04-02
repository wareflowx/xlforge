"""Error codes and exceptions for xlforge."""

from __future__ import annotations

from enum import IntEnum


class ErrorCode(IntEnum):
    """All xlforge error codes."""

    # Success
    SUCCESS = 0

    # General errors (1-15)
    GENERAL_ERROR = 1
    FILE_NOT_FOUND = 2
    ENGINE_MISMATCH = 50
    FILE_CORRUPTED = 51
    CANNOT_KILL = 52
    TEMPLATE_NOT_FOUND = 53
    RECOVERY_FAILED = 54
    SHEET_NOT_FOUND = 3
    CELL_NOT_FOUND = 4
    INVALID_SYNTAX = 5
    FILE_LOCKED = 6
    COM_ERROR = 7
    EXCEL_BUSY = 8
    FEATURE_UNAVAILABLE = 9
    EXCEL_HUNG = 10
    TYPE_COERCION_FAILED = 11
    RANGE_TOO_LARGE = 12
    CHART_NOT_FOUND = 13
    INVALID_CHART_TYPE = 14
    CHART_EXISTS = 15

    # Checkpoint errors (20-24)
    CHECKPOINT_NOT_FOUND = 20
    CHECKPOINT_RESTORE_FAILED = 21
    BRANCH_NOT_FOUND = 22
    BRANCH_MERGE_CONFLICT = 23
    CANNOT_DELETE_ACTIVE_BRANCH = 24

    # Column/Row errors (30-33)
    COLUMN_NOT_FOUND = 30
    ROW_NOT_FOUND = 31
    INVALID_UNIT = 32
    COLUMN_ROW_HIDDEN = 33

    # CSV errors (40-45)
    CSV_NOT_FOUND = 40
    ENCODING_ERROR = 41
    CSV_TYPE_COERCION_FAILED = 42
    HEADER_MISMATCH = 43
    SHEET_NOT_FOUND_DURING_EXPORT = 44
    INVALID_CSV_FORMAT = 45

    # Style errors (60-64)
    INVALID_STYLE_STRING = 60
    INVALID_NUMBER_FORMAT = 61
    NAMED_STYLE_NOT_FOUND = 62
    CONDITIONAL_FORMAT_NOT_SUPPORTED = 63
    RANGE_TOO_COMPLEX_FOR_CONDITIONAL = 64

    # Protection errors (70-79)
    SHEET_PROTECTED = 70
    PASSWORD_REQUIRED = 71
    INVALID_PASSWORD = 72
    WORKBOOK_PROTECTED = 73
    CANNOT_UNHIDE_VERY_HIDDEN_SHEET = 74
    CELL_LOCKED = 75
    INVALID_PROTECTION_OPTION = 76
    CANNOT_DELETE_LAST_SHEET = 77
    CIRCULAR_SHEET_REFERENCE = 78
    CANNOT_MOVE_SHEET = 79

    # Semantic/AI errors (80-88)
    INDEX_NOT_FOUND = 80
    LLM_PROVIDER_ERROR = 81
    PRIVACY_CHECK_FAILED = 82
    CONFIDENCE_BELOW_THRESHOLD = 83
    RECORDING_ALREADY_ACTIVE = 84
    NO_ACTIVE_RECORDING = 85
    WATCHDOG_TIMEOUT = 86
    INVALID_RULE_SYNTAX = 87
    TYPE_COERCION_FAILED_STRICT = 88

    # Database errors (89-96)
    DATABASE_CONNECTION_FAILED = 89
    QUERY_TIMEOUT = 90
    FILE_IS_LOCKED = 91
    UPSERT_KEY_COLUMN_NOT_FOUND = 92
    SCHEMA_MISMATCH = 93
    EXTENSION_NOT_AVAILABLE = 94
    VIRTUAL_VIEW_CONNECTION_FAILED = 95
    PIVOT_REFRESH_FAILED = 96

    # Table errors (100-116)
    TABLE_NOT_FOUND = 100
    TABLE_ALREADY_EXISTS = 101
    INVALID_TABLE_NAME = 102
    LINK_CONNECTION_FAILED = 103
    REFRESH_TIMEOUT = 104
    WRITEBACK_KEY_COLUMN_NOT_FOUND = 105
    NO_DIRTY_ROWS = 106
    SCHEMA_DRIFT_DETECTED = 107
    PIVOT_CREATION_FAILED = 108
    FORMULA_COLUMN_SYNTAX_ERROR = 109
    VALUE_VIOLATES_VALIDATION = 110
    VALIDATION_TYPE_NOT_SUPPORTED = 111
    INVALID_FORMULA_SYNTAX = 112
    DEPENDENT_VALIDATION_MAP_NOT_FOUND = 113
    PARENT_CELL_VALIDATION_NOT_FOUND = 114
    CIRCULAR_DEPENDENCY_IN_VALIDATION = 115
    VALIDATION_RANGE_TOO_LARGE = 116

    # Watcher errors (120-127)
    WATCHER_ALREADY_ACTIVE = 120
    NO_ACTIVE_WATCHER = 121
    WATCHER_PID_NOT_FOUND = 122
    HEADLESS_MODE_NOT_SUPPORTED = 123
    FILE_DOES_NOT_EXIST = 124
    CONDITION_SYNTAX_ERROR = 125
    WATCHER_TIMEOUT = 126
    APP_READY_CHECK_FAILED = 127


ERROR_MESSAGES: dict[int, str] = {
    # Success
    ErrorCode.SUCCESS: "Success",
    # General errors
    ErrorCode.GENERAL_ERROR: "General error",
    ErrorCode.FILE_NOT_FOUND: "File not found",
    ErrorCode.ENGINE_MISMATCH: "Engine mismatch",
    ErrorCode.FILE_CORRUPTED: "File corrupted",
    ErrorCode.CANNOT_KILL: "Cannot kill (file in use)",
    ErrorCode.TEMPLATE_NOT_FOUND: "Template not found",
    ErrorCode.RECOVERY_FAILED: "Recovery failed",
    ErrorCode.SHEET_NOT_FOUND: "Sheet not found",
    ErrorCode.CELL_NOT_FOUND: "Cell not found",
    ErrorCode.INVALID_SYNTAX: "Invalid syntax",
    ErrorCode.FILE_LOCKED: "File is locked (retry with backoff)",
    ErrorCode.COM_ERROR: "COM error",
    ErrorCode.EXCEL_BUSY: "Excel is busy (timeout in app idle)",
    ErrorCode.FEATURE_UNAVAILABLE: "Feature unavailable (e.g., chart in headless mode)",
    ErrorCode.EXCEL_HUNG: "Excel is hung (use app recover)",
    ErrorCode.TYPE_COERCION_FAILED: "Type coercion failed",
    ErrorCode.RANGE_TOO_LARGE: "Range too large (use cell bulk instead)",
    ErrorCode.CHART_NOT_FOUND: "Chart not found",
    ErrorCode.INVALID_CHART_TYPE: "Invalid chart type",
    ErrorCode.CHART_EXISTS: "Chart name already exists (use --replace)",
    # Checkpoint errors
    ErrorCode.CHECKPOINT_NOT_FOUND: "Checkpoint not found",
    ErrorCode.CHECKPOINT_RESTORE_FAILED: "Checkpoint restore failed",
    ErrorCode.BRANCH_NOT_FOUND: "Branch not found",
    ErrorCode.BRANCH_MERGE_CONFLICT: "Branch merge conflict",
    ErrorCode.CANNOT_DELETE_ACTIVE_BRANCH: "Cannot delete active branch",
    # Column/Row errors
    ErrorCode.COLUMN_NOT_FOUND: "Column not found",
    ErrorCode.ROW_NOT_FOUND: "Row not found",
    ErrorCode.INVALID_UNIT: "Invalid unit (use px, pt, or excel)",
    ErrorCode.COLUMN_ROW_HIDDEN: "Column/row is hidden",
    # CSV errors
    ErrorCode.CSV_NOT_FOUND: "CSV not found",
    ErrorCode.ENCODING_ERROR: "Encoding error",
    ErrorCode.CSV_TYPE_COERCION_FAILED: "Type coercion failed",
    ErrorCode.HEADER_MISMATCH: "Header mismatch",
    ErrorCode.SHEET_NOT_FOUND_DURING_EXPORT: "Sheet not found during export",
    ErrorCode.INVALID_CSV_FORMAT: "Invalid CSV format",
    # Style errors
    ErrorCode.INVALID_STYLE_STRING: "Invalid style string",
    ErrorCode.INVALID_NUMBER_FORMAT: "Invalid number format",
    ErrorCode.NAMED_STYLE_NOT_FOUND: "Named style not found",
    ErrorCode.CONDITIONAL_FORMAT_NOT_SUPPORTED: "Conditional format not supported",
    ErrorCode.RANGE_TOO_COMPLEX_FOR_CONDITIONAL: "Range too complex for conditional format",
    # Protection errors
    ErrorCode.SHEET_PROTECTED: "Sheet is protected",
    ErrorCode.PASSWORD_REQUIRED: "Password required",
    ErrorCode.INVALID_PASSWORD: "Invalid password",
    ErrorCode.WORKBOOK_PROTECTED: "Workbook is protected",
    ErrorCode.CANNOT_UNHIDE_VERY_HIDDEN_SHEET: "Cannot unhide very-hidden sheet",
    ErrorCode.CELL_LOCKED: "Cell is locked",
    ErrorCode.INVALID_PROTECTION_OPTION: "Invalid protection option",
    ErrorCode.CANNOT_DELETE_LAST_SHEET: "Cannot delete last sheet (workbook must have at least one)",
    ErrorCode.CIRCULAR_SHEET_REFERENCE: "Circular sheet reference in move",
    ErrorCode.CANNOT_MOVE_SHEET: "Cannot move sheet that doesn't exist",
    # Semantic/AI errors
    ErrorCode.INDEX_NOT_FOUND: "Index not found (run index create first)",
    ErrorCode.LLM_PROVIDER_ERROR: "LLM provider error (quota, network)",
    ErrorCode.PRIVACY_CHECK_FAILED: "Privacy check failed (sensitive data detected)",
    ErrorCode.CONFIDENCE_BELOW_THRESHOLD: "Confidence below threshold",
    ErrorCode.RECORDING_ALREADY_ACTIVE: "Recording already active",
    ErrorCode.NO_ACTIVE_RECORDING: "No active recording to stop",
    ErrorCode.WATCHDOG_TIMEOUT: "Watchdog timeout (Excel connection lost)",
    ErrorCode.INVALID_RULE_SYNTAX: "Invalid rule syntax for semantic-check",
    ErrorCode.TYPE_COERCION_FAILED_STRICT: "Type coercion failed (use --strict for details)",
    # Database errors
    ErrorCode.DATABASE_CONNECTION_FAILED: "Database connection failed",
    ErrorCode.QUERY_TIMEOUT: "Query timeout",
    ErrorCode.FILE_IS_LOCKED: "File is locked (Excel has it open)",
    ErrorCode.UPSERT_KEY_COLUMN_NOT_FOUND: "Upsert key column not found",
    ErrorCode.SCHEMA_MISMATCH: "Schema mismatch (use --strict or manual CAST)",
    ErrorCode.EXTENSION_NOT_AVAILABLE: "Extension not available (e.g., postgres_scanner)",
    ErrorCode.VIRTUAL_VIEW_CONNECTION_FAILED: "Virtual view connection failed",
    ErrorCode.PIVOT_REFRESH_FAILED: "Pivot refresh failed (no matching pivot found)",
    # Table errors
    ErrorCode.TABLE_NOT_FOUND: "Table not found",
    ErrorCode.TABLE_ALREADY_EXISTS: "Table already exists (use --replace)",
    ErrorCode.INVALID_TABLE_NAME: "Invalid table name",
    ErrorCode.LINK_CONNECTION_FAILED: "Link connection failed",
    ErrorCode.REFRESH_TIMEOUT: "Refresh timeout (data not loaded in time)",
    ErrorCode.WRITEBACK_KEY_COLUMN_NOT_FOUND: "Writeback key column not found",
    ErrorCode.NO_DIRTY_ROWS: "No dirty rows to writeback",
    ErrorCode.SCHEMA_DRIFT_DETECTED: "Schema drift detected (use --strict or --prune)",
    ErrorCode.PIVOT_CREATION_FAILED: "Pivot creation failed (source table empty)",
    ErrorCode.FORMULA_COLUMN_SYNTAX_ERROR: "Formula column syntax error",
    ErrorCode.VALUE_VIOLATES_VALIDATION: "Value violates validation (strict mode)",
    ErrorCode.VALIDATION_TYPE_NOT_SUPPORTED: "Validation type not supported",
    ErrorCode.INVALID_FORMULA_SYNTAX: "Invalid formula syntax",
    ErrorCode.DEPENDENT_VALIDATION_MAP_NOT_FOUND: "Dependent validation map not found",
    ErrorCode.PARENT_CELL_VALIDATION_NOT_FOUND: "Parent cell validation not found",
    ErrorCode.CIRCULAR_DEPENDENCY_IN_VALIDATION: "Circular dependency in dependent validation",
    ErrorCode.VALIDATION_RANGE_TOO_LARGE: "Validation range is too large",
    # Watcher errors
    ErrorCode.WATCHER_ALREADY_ACTIVE: "Watcher already active for this file",
    ErrorCode.NO_ACTIVE_WATCHER: "No active watcher to stop",
    ErrorCode.WATCHER_PID_NOT_FOUND: "Watcher PID not found (stale pid file)",
    ErrorCode.HEADLESS_MODE_NOT_SUPPORTED: "Headless mode not supported on this platform",
    ErrorCode.FILE_DOES_NOT_EXIST: "File does not exist",
    ErrorCode.CONDITION_SYNTAX_ERROR: "Condition syntax error",
    ErrorCode.WATCHER_TIMEOUT: "Watcher timeout (Excel closed, no re-open)",
    ErrorCode.APP_READY_CHECK_FAILED: "App.Ready check failed (Excel in bad state)",
}


class XlforgeError(Exception):
    """Base exception for xlforge errors.

    Attributes:
        code: The error code from ErrorCode enum.
        message: Human-readable error message.
    """

    def __init__(
        self,
        code: ErrorCode,
        message: str | None = None,
        details: dict | None = None,
    ):
        self.code = code
        self.message = message or ERROR_MESSAGES.get(int(code), "Unknown error")
        self.details = details or {}
        super().__init__(self.message)

    def __repr__(self) -> str:
        return f"XlforgeError({self.code.name}, {self.message!r})"

    def __str__(self) -> str:
        if self.details:
            return f"[{self.code.name}] {self.message} ({self.details})"
        return f"[{self.code.name}] {self.message}"

    def to_dict(self) -> dict:
        """Convert error to dictionary for JSON output."""
        return {
            "success": False,
            "code": int(self.code),
            "error": self.code.name,
            "message": self.message,
            "details": self.details,
        }


def get_error_message(code: ErrorCode | int) -> str:
    """Get the human-readable message for an error code."""
    return ERROR_MESSAGES.get(int(code), "Unknown error")


def is_success(code: ErrorCode | int) -> bool:
    """Check if an error code represents success."""
    return int(code) == ErrorCode.SUCCESS
