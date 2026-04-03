"""Engine selector for automatic engine selection."""

from __future__ import annotations

from importlib.util import find_spec
from pathlib import Path

from xlforge.core.engines.base import Engine
from xlforge.core.engines.openpyxl_engine import OpenpyxlEngine


class EngineSelector:
    """Select appropriate engine based on environment."""

    @classmethod
    def for_path(cls, path: Path) -> Engine:
        """Select engine for a file path.

        Args:
            path: Path to the workbook file.

        Returns:
            Engine instance - XlwingsEngine if Excel is available,
            otherwise OpenpyxlEngine.
        """
        if cls._is_excel_available():
            return cls._get_xlwings_engine()()
        return OpenpyxlEngine()

    @classmethod
    def for_engine_name(cls, name: str) -> Engine:
        """Select engine by name.

        Args:
            name: Engine name ('xlwings' or 'openpyxl').

        Returns:
            Engine instance.

        Raises:
            ValueError: If engine name is unknown.
        """
        engines: dict[str, type[Engine]] = {
            "openpyxl": OpenpyxlEngine,
        }
        if name == "xlwings":
            engines["xlwings"] = cls._get_xlwings_engine()
        elif name not in engines:
            available = list(engines.keys()) + (
                ["xlwings"] if cls._xlwings_engine_available() else []
            )
            raise ValueError(f"Unknown engine: {name}. Available: {available}")
        return engines[name]()

    @staticmethod
    def _is_excel_available() -> bool:
        """Check if Excel is available via xlwings.

        Returns:
            True if xlwings module is available.
        """
        return find_spec("xlwings") is not None

    @classmethod
    def _xlwings_engine_available(cls) -> bool:
        """Check if xlwings_engine module is available.

        Returns:
            True if xlwings_engine module can be imported.
        """
        return find_spec("xlforge.core.engines.xlwings_engine") is not None

    @classmethod
    def _get_xlwings_engine(cls) -> type[Engine]:
        """Get XlwingsEngine class.

        Returns:
            XlwingsEngine class.

        Raises:
            ImportError: If xlwings or xlwings_engine is not available.
        """
        if not cls._xlwings_engine_available():
            msg = (
                "xlwings engine is not available. Install xlforge with xlwings support."
            )
            raise ImportError(msg)
        from xlforge.core.engines.xlwings_engine import XlwingsEngine  # type: ignore[import-not-found]

        return XlwingsEngine
