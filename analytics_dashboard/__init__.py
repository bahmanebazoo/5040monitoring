"""
5040 Monitoring Analytics — Public API
=======================================
backward-compatible: هم تابع هم کلاس.
"""

from pathlib import Path

from .dashboard_generator import DashboardGenerator
from .data_preparator import DataPreparator, PreparedData
from .style_manager import StyleManager
from .column_resolver import ColumnResolver


def generate_analytics_dashboard(
    enriched_path: str | Path,
    output_path: str | Path | None = None,
) -> Path:
    """
    تابع سازگار با main.py قبلی.
    """
    generator = DashboardGenerator()
    return generator.generate(enriched_path, output_path)


__all__ = [
    "generate_analytics_dashboard",
    "DashboardGenerator",
    "DataPreparator",
    "PreparedData",
    "StyleManager",
    "ColumnResolver",
]
